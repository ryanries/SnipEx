// SnipExHijack.c
// Author: Joseph Ryan Ries, 2017-2020
// Implements the Snipping Tool replacement feature for both classic (IFEO) and modern (MSIX) Windows.

#ifndef UNICODE
#define UNICODE
#endif
#ifndef _UNICODE
#define _UNICODE
#endif

#pragma warning(push, 0)
#include <windows.h>
#include <shobjidl_core.h>
#include <shellapi.h>
#include <shlwapi.h>
#include <stdio.h>
#pragma warning(pop)

#pragma comment(lib, "Ole32.lib")
#pragma comment(lib, "Shell32.lib")
#pragma comment(lib, "Shlwapi.lib")
#pragma comment(lib, "Advapi32.lib")

#pragma warning(disable: 4820)
#pragma warning(disable: 4710)
#pragma warning(disable: 5045)

#include "SnipEx.h"
#include "SnipExHijack.h"


// Function pointers for dynamically resolved Win8+ package APIs from kernel32.dll.
typedef LONG (WINAPI *PFN_GetPackagesByPackageFamily)(
    _In_                                PCWSTR  packageFamilyName,
    _Inout_                             UINT32* count,
    _Out_writes_opt_(*count)            PWSTR*  packageFullNames,
    _Inout_                             UINT32* bufferLength,
    _Out_writes_opt_(*bufferLength)     WCHAR*  buffer);

typedef LONG (WINAPI *PFN_GetPackageFamilyName)(
    _In_                                HANDLE  hProcess,
    _Inout_                             UINT32* packageFamilyNameLength,
    _Out_writes_opt_(*packageFamilyNameLength) PWSTR packageFamilyName);


static const wchar_t* gIfeoParentKeyPath = L"SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Image File Execution Options";

static const wchar_t* gIfeoSubkeyName    = L"SnippingTool.exe";


// Opens an HKLM key with KEY_WOW64_64KEY so 32-bit and 64-bit builds use the same view.
static LSTATUS OpenHklmKey64(_In_ const wchar_t* SubKeyPath, _In_ REGSAM DesiredAccess, _Out_ HKEY* Key)
{
    return RegOpenKeyExW(HKEY_LOCAL_MACHINE, SubKeyPath, 0, DesiredAccess | KEY_WOW64_64KEY, Key);
}


static LSTATUS CreateHklmKey64(_In_ const wchar_t* SubKeyPath, _In_ REGSAM DesiredAccess, _Out_ HKEY* Key, _Out_opt_ DWORD* Disposition)
{
    DWORD LocalDisposition = 0;

    LSTATUS Status = RegCreateKeyExW(
        HKEY_LOCAL_MACHINE, SubKeyPath, 0, NULL, 0,
        DesiredAccess | KEY_WOW64_64KEY, NULL, Key, &LocalDisposition);

    if (Disposition != NULL)
    {
        *Disposition = LocalDisposition;
    }

    return Status;
}


// Gets the path where SnipEx should be registered when hooking the Snipping Tool.
// If the running executable is already under a protected location, returns its path.
// Otherwise returns %ProgramFiles%\SnipEx\SnipEx.exe.
static BOOL GetProtectedSnipExPath(_Out_writes_(MAX_PATH) wchar_t* ProtectedPath, _Out_ BOOL* NeedsCopy)
{
    wchar_t CurrentPath[MAX_PATH] = { 0 };

    *NeedsCopy = FALSE;

    ProtectedPath[0] = L'\0';

    if (GetModuleFileNameW(NULL, CurrentPath, _countof(CurrentPath)) == 0)
    {
        return FALSE;
    }

    wchar_t ProgramFilesPath[MAX_PATH] = { 0 };

    wchar_t ProgramFilesX86Path[MAX_PATH] = { 0 };

    wchar_t SystemRootPath[MAX_PATH] = { 0 };

    GetEnvironmentVariableW(L"ProgramFiles", ProgramFilesPath, _countof(ProgramFilesPath));

    GetEnvironmentVariableW(L"ProgramFiles(x86)", ProgramFilesX86Path, _countof(ProgramFilesX86Path));

    GetEnvironmentVariableW(L"SystemRoot", SystemRootPath, _countof(SystemRootPath));

    BOOL IsProtected = FALSE;

    if (wcslen(ProgramFilesPath) > 0 && _wcsnicmp(CurrentPath, ProgramFilesPath, wcslen(ProgramFilesPath)) == 0)
    {
        IsProtected = TRUE;
    }
    else if (wcslen(ProgramFilesX86Path) > 0 && _wcsnicmp(CurrentPath, ProgramFilesX86Path, wcslen(ProgramFilesX86Path)) == 0)
    {
        IsProtected = TRUE;
    }
    else if (wcslen(SystemRootPath) > 0 && _wcsnicmp(CurrentPath, SystemRootPath, wcslen(SystemRootPath)) == 0)
    {
        IsProtected = TRUE;
    }

    if (IsProtected)
    {
        wcscpy_s(ProtectedPath, MAX_PATH, CurrentPath);

        return TRUE;
    }

    // Build a path under %ProgramFiles%\SnipEx
    if (wcslen(ProgramFilesPath) == 0)
    {
        return FALSE;
    }

    swprintf_s(ProtectedPath, MAX_PATH, L"%s\\SnipEx\\SnipEx.exe", ProgramFilesPath);

    *NeedsCopy = TRUE;

    return TRUE;
}


// Copies the running SnipEx.exe to the protected installation path.
static BOOL CopySnipExToProtectedLocation(_In_ const wchar_t* DestinationPath)
{
    wchar_t CurrentPath[MAX_PATH] = { 0 };

    if (GetModuleFileNameW(NULL, CurrentPath, _countof(CurrentPath)) == 0)
    {
        return FALSE;
    }

    // Create the directory if needed.
    wchar_t Directory[MAX_PATH] = { 0 };

    wcscpy_s(Directory, _countof(Directory), DestinationPath);

    PathRemoveFileSpecW(Directory);

    if (GetFileAttributesW(Directory) == INVALID_FILE_ATTRIBUTES)
    {
        if (!CreateDirectoryW(Directory, NULL))
        {
            return FALSE;
        }
    }

    return CopyFileW(CurrentPath, DestinationPath, FALSE);
}


// Resolves Snipping Tool packages for the current user.
static BOOL EnumerateSnippingToolPackages(_Out_ SNIPPINGTOOLPACKAGESET* PackageSet)
{
    PackageSet->Count = 0;

    HMODULE Kernel32 = GetModuleHandleW(L"kernel32.dll");

    if (Kernel32 == NULL)
    {
        return FALSE;
    }

    PFN_GetPackagesByPackageFamily GetPackagesByPackageFamilyFn =
        (PFN_GetPackagesByPackageFamily)(void*)GetProcAddress(Kernel32, "GetPackagesByPackageFamily");

    if (GetPackagesByPackageFamilyFn == NULL)
    {
        return FALSE;
    }

    UINT32 Count = 0;

    UINT32 BufferLength = 0;

    LONG Status = GetPackagesByPackageFamilyFn(
        SNIPPINGTOOL_PACKAGE_FAMILY_NAME, &Count, NULL, &BufferLength, NULL);

    if (Status != ERROR_INSUFFICIENT_BUFFER || Count == 0)
    {
        return FALSE;
    }

    if (Count > MAX_SNIPPINGTOOL_PACKAGES)
    {
        Count = MAX_SNIPPINGTOOL_PACKAGES;
    }

    PWSTR* FullNamePointers = (PWSTR*)HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, Count * sizeof(PWSTR));

    WCHAR* Buffer = (WCHAR*)HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, BufferLength * sizeof(WCHAR));

    if (FullNamePointers == NULL || Buffer == NULL)
    {
        if (FullNamePointers) HeapFree(GetProcessHeap(), 0, FullNamePointers);

        if (Buffer) HeapFree(GetProcessHeap(), 0, Buffer);

        return FALSE;
    }

    Status = GetPackagesByPackageFamilyFn(
        SNIPPINGTOOL_PACKAGE_FAMILY_NAME, &Count, FullNamePointers, &BufferLength, Buffer);

    if (Status == ERROR_SUCCESS)
    {
        for (UINT32 Index = 0; Index < Count && PackageSet->Count < MAX_SNIPPINGTOOL_PACKAGES; Index++)
        {
            if (FullNamePointers[Index] != NULL && wcslen(FullNamePointers[Index]) <= MAX_PACKAGE_FULL_NAME_LENGTH)
            {
                wcscpy_s(PackageSet->FullNames[PackageSet->Count], MAX_PACKAGE_FULL_NAME_LENGTH + 1, FullNamePointers[Index]);

                PackageSet->Count++;
            }
        }
    }

    HeapFree(GetProcessHeap(), 0, FullNamePointers);

    HeapFree(GetProcessHeap(), 0, Buffer);

    return PackageSet->Count > 0;
}


// Reads the REG_MULTI_SZ "HookedPackages" journal from HKLM\SOFTWARE\SnipEx.
static BOOL ReadHookedPackagesJournal(_Out_ SNIPPINGTOOLPACKAGESET* PackageSet)
{
    PackageSet->Count = 0;

    HKEY SnipExKey = NULL;

    if (OpenHklmKey64(SNIPEX_MACHINE_KEY_PATH, KEY_READ, &SnipExKey) != ERROR_SUCCESS)
    {
        return FALSE;
    }

    WCHAR Buffer[4096] = { 0 };

    DWORD BufferSize = sizeof(Buffer);

    DWORD ValueType = 0;

    LSTATUS Status = RegQueryValueExW(SnipExKey, L"HookedPackages", NULL, &ValueType, (LPBYTE)Buffer, &BufferSize);

    RegCloseKey(SnipExKey);

    if (Status != ERROR_SUCCESS || ValueType != REG_MULTI_SZ)
    {
        return FALSE;
    }

    const wchar_t* Current = Buffer;

    while (*Current != L'\0' && PackageSet->Count < MAX_SNIPPINGTOOL_PACKAGES)
    {
        if (wcslen(Current) <= MAX_PACKAGE_FULL_NAME_LENGTH)
        {
            wcscpy_s(PackageSet->FullNames[PackageSet->Count], MAX_PACKAGE_FULL_NAME_LENGTH + 1, Current);

            PackageSet->Count++;
        }

        Current += wcslen(Current) + 1;
    }

    return PackageSet->Count > 0;
}


// Writes or appends to the REG_MULTI_SZ HookedPackages journal.
static LSTATUS WriteHookedPackagesJournal(_In_ const SNIPPINGTOOLPACKAGESET* PackageSet)
{
    HKEY SnipExKey = NULL;

    LSTATUS Status = CreateHklmKey64(SNIPEX_MACHINE_KEY_PATH, KEY_SET_VALUE, &SnipExKey, NULL);

    if (Status != ERROR_SUCCESS)
    {
        return Status;
    }

    // Build REG_MULTI_SZ: each string NUL-terminated, double-NUL at the end.
    WCHAR Buffer[4096] = { 0 };

    DWORD Offset = 0;

    for (UINT32 Index = 0; Index < PackageSet->Count; Index++)
    {
        size_t Length = wcslen(PackageSet->FullNames[Index]);

        if (Offset + Length + 2 >= _countof(Buffer))
        {
            break;
        }

        wcscpy_s(&Buffer[Offset], _countof(Buffer) - Offset, PackageSet->FullNames[Index]);

        Offset += (DWORD)(Length + 1);
    }

    Buffer[Offset] = L'\0';

    DWORD DataSize = (Offset + 1) * sizeof(WCHAR);

    Status = RegSetValueExW(SnipExKey, L"HookedPackages", 0, REG_MULTI_SZ, (const BYTE*)Buffer, DataSize);

    RegCloseKey(SnipExKey);

    return Status;
}


// Stores the registered debugger command line in the journal.
static LSTATUS WriteRegisteredDebuggerPath(_In_ const wchar_t* QuotedPath)
{
    HKEY SnipExKey = NULL;

    LSTATUS Status = CreateHklmKey64(SNIPEX_MACHINE_KEY_PATH, KEY_SET_VALUE, &SnipExKey, NULL);

    if (Status != ERROR_SUCCESS)
    {
        return Status;
    }

    DWORD DataSize = (DWORD)((wcslen(QuotedPath) + 1) * sizeof(WCHAR));

    Status = RegSetValueExW(SnipExKey, L"RegisteredDebugger", 0, REG_SZ, (const BYTE*)QuotedPath, DataSize);

    RegCloseKey(SnipExKey);

    return Status;
}


// Reads the registered debugger command line from the journal.
static BOOL ReadRegisteredDebuggerPath(_Out_writes_(MAX_PATH) wchar_t* Path)
{
    Path[0] = L'\0';

    HKEY SnipExKey = NULL;

    if (OpenHklmKey64(SNIPEX_MACHINE_KEY_PATH, KEY_READ, &SnipExKey) != ERROR_SUCCESS)
    {
        return FALSE;
    }

    DWORD BufferSize = MAX_PATH * sizeof(WCHAR);

    DWORD ValueType = 0;

    LSTATUS Status = RegQueryValueExW(SnipExKey, L"RegisteredDebugger", NULL, &ValueType, (LPBYTE)Path, &BufferSize);

    RegCloseKey(SnipExKey);

    return (Status == ERROR_SUCCESS && ValueType == REG_SZ && wcslen(Path) > 0);
}


// Deletes the SnipEx journal key and all values.
static void DeleteSnipExJournal(void)
{
    HKEY ParentKey = NULL;

    if (RegOpenKeyExW(HKEY_LOCAL_MACHINE, L"SOFTWARE", 0, KEY_WRITE | KEY_WOW64_64KEY, &ParentKey) == ERROR_SUCCESS)
    {
        RegDeleteKeyExW(ParentKey, L"SnipEx", KEY_WOW64_64KEY, 0);

        RegCloseKey(ParentKey);
    }
}


// Checks if the classic System32\SnippingTool.exe still exists (Win7/8/older Win10).
static BOOL ClassicSnippingToolExists(void)
{
    wchar_t SnippingToolPath[MAX_PATH] = { 0 };

    // Use Sysnative so a 32-bit build checks the real System32 on 64-bit Windows.
    wchar_t SystemDir[MAX_PATH] = { 0 };

    GetEnvironmentVariableW(L"SystemRoot", SystemDir, _countof(SystemDir));

    swprintf_s(SnippingToolPath, _countof(SnippingToolPath), L"%s\\Sysnative\\SnippingTool.exe", SystemDir);

    if (GetFileAttributesW(SnippingToolPath) != INVALID_FILE_ATTRIBUTES)
    {
        return TRUE;
    }

    swprintf_s(SnippingToolPath, _countof(SnippingToolPath), L"%s\\System32\\SnippingTool.exe", SystemDir);

    return (GetFileAttributesW(SnippingToolPath) != INVALID_FILE_ATTRIBUTES);
}


// Reads the current IFEO Debugger value for SnippingTool.exe and extracts the executable path.
static BOOL ReadIfeoDebuggerValue(_Out_writes_(MAX_PATH) wchar_t* DebuggerPath)
{
    DebuggerPath[0] = L'\0';

    wchar_t IfeoSubkeyPath[MAX_PATH] = { 0 };

    swprintf_s(IfeoSubkeyPath, _countof(IfeoSubkeyPath),
        L"%s\\%s", gIfeoParentKeyPath, gIfeoSubkeyName);

    HKEY IfeoKey = NULL;

    if (OpenHklmKey64(IfeoSubkeyPath, KEY_READ, &IfeoKey) != ERROR_SUCCESS)
    {
        return FALSE;
    }

    wchar_t RawValue[MAX_PATH] = { 0 };

    DWORD ValueSize = sizeof(RawValue);

    DWORD ValueType = 0;

    LSTATUS Status = RegQueryValueExW(IfeoKey, L"Debugger", NULL, &ValueType, (LPBYTE)RawValue, &ValueSize);

    RegCloseKey(IfeoKey);

    if (Status != ERROR_SUCCESS || ValueType != REG_SZ || wcslen(RawValue) == 0)
    {
        return FALSE;
    }

    // Parse the value through CommandLineToArgvW to handle quoted paths correctly.
    int ArgumentCount = 0;

    LPWSTR* Arguments = CommandLineToArgvW(RawValue, &ArgumentCount);

    if (Arguments == NULL || ArgumentCount < 1)
    {
        if (Arguments) LocalFree(Arguments);

        return FALSE;
    }

    wcscpy_s(DebuggerPath, MAX_PATH, Arguments[0]);

    LocalFree(Arguments);

    return TRUE;
}


// Checks if a given debugger path belongs to SnipEx (matches our journaled registration).
static BOOL IsOurDebuggerValue(_In_ const wchar_t* DebuggerPath)
{
    if (DebuggerPath == NULL || wcslen(DebuggerPath) == 0)
    {
        return FALSE;
    }

    // Check against the journaled registered path first.
    wchar_t JournaledPath[MAX_PATH] = { 0 };

    if (ReadRegisteredDebuggerPath(JournaledPath))
    {
        // Parse argv[0] from the journaled path (it may be quoted).
        int ArgumentCount = 0;

        LPWSTR* Arguments = CommandLineToArgvW(JournaledPath, &ArgumentCount);

        if (Arguments != NULL && ArgumentCount >= 1)
        {
            if (_wcsicmp(DebuggerPath, Arguments[0]) == 0)
            {
                LocalFree(Arguments);

                return TRUE;
            }

            LocalFree(Arguments);
        }
    }

    // Fallback: check if the basename is SnipEx.exe (for legacy pre-journal installations).
    const wchar_t* BaseName = wcsrchr(DebuggerPath, L'\\');

    if (BaseName != NULL)
    {
        BaseName++;
    }
    else
    {
        BaseName = DebuggerPath;
    }

    return (_wcsicmp(BaseName, L"SnipEx.exe") == 0);
}


// Enables the IPackageDebugSettings debugger hook for a single package.
static HRESULT EnablePackageDebugging(_In_ const wchar_t* PackageFullName, _In_ const wchar_t* DebuggerCommandLine)
{
    IPackageDebugSettings* DebugSettings = NULL;

    HRESULT Result = CoCreateInstance(
        &CLSID_PackageDebugSettings, NULL, CLSCTX_ALL,
        &IID_IPackageDebugSettings, (void**)&DebugSettings);

    if (FAILED(Result))
    {
        return Result;
    }

    Result = DebugSettings->lpVtbl->EnableDebugging(DebugSettings, PackageFullName, DebuggerCommandLine, NULL);

    DebugSettings->lpVtbl->Release(DebugSettings);

    return Result;
}


// Disables the IPackageDebugSettings debugger hook for a single package.
static HRESULT DisablePackageDebugging(_In_ const wchar_t* PackageFullName)
{
    IPackageDebugSettings* DebugSettings = NULL;

    HRESULT Result = CoCreateInstance(
        &CLSID_PackageDebugSettings, NULL, CLSCTX_ALL,
        &IID_IPackageDebugSettings, (void**)&DebugSettings);

    if (FAILED(Result))
    {
        return Result;
    }

    Result = DebugSettings->lpVtbl->DisableDebugging(DebugSettings, PackageFullName);

    DebugSettings->lpVtbl->Release(DebugSettings);

    return Result;
}


// Installs the legacy IFEO Debugger hook so classic SnippingTool.exe launches SnipEx.
static LSTATUS InstallLegacyIfeoHook(_In_ const wchar_t* QuotedDebuggerPath)
{
    wchar_t IfeoSubkeyPath[MAX_PATH] = { 0 };

    swprintf_s(IfeoSubkeyPath, _countof(IfeoSubkeyPath),
        L"%s\\%s", gIfeoParentKeyPath, gIfeoSubkeyName);

    HKEY SnippingToolKey = NULL;

    LSTATUS Status = CreateHklmKey64(IfeoSubkeyPath, KEY_SET_VALUE, &SnippingToolKey, NULL);

    if (Status != ERROR_SUCCESS)
    {
        return Status;
    }

    DWORD DataSize = (DWORD)((wcslen(QuotedDebuggerPath) + 1) * sizeof(WCHAR));

    Status = RegSetValueExW(SnippingToolKey, L"Debugger", 0, REG_SZ, (const BYTE*)QuotedDebuggerPath, DataSize);

    RegCloseKey(SnippingToolKey);

    return Status;
}


// Removes the legacy IFEO Debugger hook only if it belongs to SnipEx.
static LSTATUS RemoveLegacyIfeoHook(void)
{
    wchar_t DebuggerPath[MAX_PATH] = { 0 };

    if (!ReadIfeoDebuggerValue(DebuggerPath))
    {
        return ERROR_SUCCESS;
    }

    if (!IsOurDebuggerValue(DebuggerPath))
    {
        return ERROR_ACCESS_DENIED;
    }

    wchar_t IfeoSubkeyPath[MAX_PATH] = { 0 };

    swprintf_s(IfeoSubkeyPath, _countof(IfeoSubkeyPath),
        L"%s\\%s", gIfeoParentKeyPath, gIfeoSubkeyName);

    HKEY IfeoKey = NULL;

    if (OpenHklmKey64(IfeoSubkeyPath, KEY_SET_VALUE, &IfeoKey) != ERROR_SUCCESS)
    {
        return GetLastError();
    }

    LSTATUS Status = RegDeleteValueW(IfeoKey, L"Debugger");

    RegCloseKey(IfeoKey);

    if (Status == ERROR_SUCCESS)
    {
        // Try to delete the now-empty subkey (non-fatal if it has other values).
        HKEY ParentKey = NULL;

        if (OpenHklmKey64(gIfeoParentKeyPath, KEY_WRITE, &ParentKey) == ERROR_SUCCESS)
        {
            RegDeleteKeyExW(ParentKey, gIfeoSubkeyName, KEY_WOW64_64KEY, 0);

            RegCloseKey(ParentKey);
        }
    }

    return Status;
}


// Checks whether SnipEx's own filename is SnippingTool.exe (anti-recursion guard).
static BOOL IsOwnModuleNamedSnippingTool(void)
{
    wchar_t ModulePath[MAX_PATH] = { 0 };

    GetModuleFileNameW(NULL, ModulePath, _countof(ModulePath));

    const wchar_t* BaseName = wcsrchr(ModulePath, L'\\');

    if (BaseName != NULL)
    {
        BaseName++;
    }
    else
    {
        BaseName = ModulePath;
    }

    return (_wcsicmp(BaseName, L"SnippingTool.exe") == 0);
}


SNIPPINGTOOLHOOKSTATE GetSnippingToolHookState(void)
{
    // Check the modern hook (journal-based).
    SNIPPINGTOOLPACKAGESET JournaledPackages = { 0 };

    SNIPPINGTOOLPACKAGESET InstalledPackages = { 0 };

    if (ReadHookedPackagesJournal(&JournaledPackages) && EnumerateSnippingToolPackages(&InstalledPackages))
    {
        for (UINT32 Journaled = 0; Journaled < JournaledPackages.Count; Journaled++)
        {
            for (UINT32 Installed = 0; Installed < InstalledPackages.Count; Installed++)
            {
                if (_wcsicmp(JournaledPackages.FullNames[Journaled], InstalledPackages.FullNames[Installed]) == 0)
                {
                    return SNIPPINGTOOLHOOKSTATE_REPLACED;
                }
            }
        }
    }

    // Check the legacy hook (IFEO).
    wchar_t DebuggerPath[MAX_PATH] = { 0 };

    if (ReadIfeoDebuggerValue(DebuggerPath))
    {
        if (IsOurDebuggerValue(DebuggerPath))
        {
            DWORD Attributes = GetFileAttributesW(DebuggerPath);

            if (Attributes != INVALID_FILE_ATTRIBUTES && !(Attributes & FILE_ATTRIBUTE_DIRECTORY))
            {
                return SNIPPINGTOOLHOOKSTATE_REPLACED;
            }
        }
        else
        {
            return SNIPPINGTOOLHOOKSTATE_FOREIGNDEBUGGER;
        }
    }

    return SNIPPINGTOOLHOOKSTATE_NOTREPLACED;
}


HIJACKEDLAUNCHRESULT TerminateHijackedSnippingToolProcess(void)
{
    int ArgumentCount = 0;

    LPWSTR* Arguments = CommandLineToArgvW(GetCommandLineW(), &ArgumentCount);

    if (Arguments == NULL)
    {
        return HIJACKEDLAUNCHRESULT_NOTHIJACKEDLAUNCH;
    }

    DWORD TargetPid = 0;

    BOOL FoundPid = FALSE;

    BOOL FoundTid = FALSE;

    for (int Index = 1; Index < ArgumentCount - 1; Index++)
    {
        if (_wcsicmp(Arguments[Index], L"-p") == 0)
        {
            wchar_t* EndPointer = NULL;

            unsigned long Parsed = wcstoul(Arguments[Index + 1], &EndPointer, 10);

            if (EndPointer != NULL && *EndPointer == L'\0' && Parsed > 0 && Parsed <= 0xFFFFFFFF)
            {
                TargetPid = (DWORD)Parsed;

                FoundPid = TRUE;
            }
        }
        else if (_wcsicmp(Arguments[Index], L"-tid") == 0)
        {
            wchar_t* EndPointer = NULL;

            unsigned long Parsed = wcstoul(Arguments[Index + 1], &EndPointer, 10);

            if (EndPointer != NULL && *EndPointer == L'\0' && Parsed > 0)
            {
                FoundTid = TRUE;
            }
        }
    }

    LocalFree(Arguments);

    if (!FoundPid || !FoundTid || TargetPid == 0)
    {
        return HIJACKEDLAUNCHRESULT_NOTHIJACKEDLAUNCH;
    }

    // Open the process and verify its package identity.
    HANDLE ProcessHandle = OpenProcess(
        PROCESS_TERMINATE | SYNCHRONIZE | PROCESS_QUERY_LIMITED_INFORMATION,
        FALSE, TargetPid);

    if (ProcessHandle == NULL)
    {
        return HIJACKEDLAUNCHRESULT_REJECTEDFOREIGNPROCESS;
    }

    // Dynamically resolve GetPackageFamilyName (Win8+).
    HMODULE Kernel32 = GetModuleHandleW(L"kernel32.dll");

    PFN_GetPackageFamilyName GetPackageFamilyNameFn = NULL;

    if (Kernel32 != NULL)
    {
        GetPackageFamilyNameFn = (PFN_GetPackageFamilyName)(void*)GetProcAddress(Kernel32, "GetPackageFamilyName");
    }

    if (GetPackageFamilyNameFn == NULL)
    {
        CloseHandle(ProcessHandle);

        return HIJACKEDLAUNCHRESULT_REJECTEDFOREIGNPROCESS;
    }

    wchar_t FamilyName[256] = { 0 };

    UINT32 FamilyNameLength = _countof(FamilyName);

    LONG PackageStatus = GetPackageFamilyNameFn(ProcessHandle, &FamilyNameLength, FamilyName);

    if (PackageStatus != ERROR_SUCCESS || _wcsicmp(FamilyName, SNIPPINGTOOL_PACKAGE_FAMILY_NAME) != 0)
    {
        CloseHandle(ProcessHandle);

        return HIJACKEDLAUNCHRESULT_REJECTEDFOREIGNPROCESS;
    }

    // Verified as a Snipping Tool process. End it so SnipEx can take over.
    if (!TerminateProcess(ProcessHandle, 0))
    {
        CloseHandle(ProcessHandle);

        return HIJACKEDLAUNCHRESULT_TERMINATIONFAILED;
    }

    WaitForSingleObject(ProcessHandle, 5000);

    CloseHandle(ProcessHandle);

    return HIJACKEDLAUNCHRESULT_SNIPPINGTOOLTERMINATED;
}


ELEVATEDHIJACKREQUEST GetElevatedHijackRequest(void)
{
    int ArgumentCount = 0;

    LPWSTR* Arguments = CommandLineToArgvW(GetCommandLineW(), &ArgumentCount);

    if (Arguments == NULL)
    {
        return ELEVATEDHIJACKREQUEST_NONE;
    }

    ELEVATEDHIJACKREQUEST Request = ELEVATEDHIJACKREQUEST_NONE;

    for (int Index = 1; Index < ArgumentCount; Index++)
    {
        if (_wcsicmp(Arguments[Index], REPLACE_SNIPPINGTOOL_VERB) == 0)
        {
            Request = ELEVATEDHIJACKREQUEST_REPLACE;

            break;
        }
        else if (_wcsicmp(Arguments[Index], RESTORE_SNIPPINGTOOL_VERB) == 0)
        {
            Request = ELEVATEDHIJACKREQUEST_RESTORE;

            break;
        }
    }

    LocalFree(Arguments);

    return Request;
}


BOOL ReplaceSnippingToolWithSnipEx(_In_opt_ HWND OwnerWindow)
{
    MyOutputDebugStringW(L"[%s] Line %d: Entered.\n", __FUNCTIONW__, __LINE__);

    if (IsOwnModuleNamedSnippingTool())
    {
        MessageBoxW(OwnerWindow,
            L"SnipEx cannot replace the Snipping Tool when its own executable is named SnippingTool.exe. "
            L"Please rename SnipEx.exe first.",
            L"Error", MB_OK | MB_ICONERROR | MB_SYSTEMMODAL);

        return FALSE;
    }

    // Determine the protected path to register.
    wchar_t ProtectedPath[MAX_PATH] = { 0 };

    BOOL NeedsCopy = FALSE;

    if (!GetProtectedSnipExPath(ProtectedPath, &NeedsCopy))
    {
        MessageBoxW(OwnerWindow,
            L"Could not determine a protected installation path for SnipEx.",
            L"Error", MB_OK | MB_ICONERROR | MB_SYSTEMMODAL);

        return FALSE;
    }

    if (NeedsCopy)
    {
        int UserChoice = MessageBoxW(OwnerWindow,
            L"SnipEx is running from an unprotected location. For security, it needs to be "
            L"copied to Program Files before it can replace the Snipping Tool.\n\n"
            L"Copy SnipEx.exe to Program Files and continue?",
            L"Security", MB_OKCANCEL | MB_ICONINFORMATION | MB_SYSTEMMODAL);

        if (UserChoice != IDOK)
        {
            return FALSE;
        }

        if (!CopySnipExToProtectedLocation(ProtectedPath))
        {
            MessageBoxW(OwnerWindow,
                L"Failed to copy SnipEx.exe to Program Files.",
                L"Error", MB_OK | MB_ICONERROR | MB_SYSTEMMODAL);

            return FALSE;
        }
    }

    // Build the quoted debugger command line that will be registered.
    wchar_t QuotedDebuggerPath[MAX_PATH + 4] = { 0 };

    swprintf_s(QuotedDebuggerPath, _countof(QuotedDebuggerPath), L"\"%s\"", ProtectedPath);

    BOOL LegacyApplied = FALSE;

    BOOL ModernApplied = FALSE;

    BOOL LegacyApplicable = ClassicSnippingToolExists();

    SNIPPINGTOOLPACKAGESET InstalledPackages = { 0 };

    BOOL ModernApplicable = EnumerateSnippingToolPackages(&InstalledPackages);

    if (!LegacyApplicable && !ModernApplicable)
    {
        MessageBoxW(OwnerWindow,
            L"The Windows Snipping Tool was not found on this system. "
            L"Neither the classic desktop version nor the packaged Store version could be located.",
            L"Not Found", MB_OK | MB_ICONWARNING | MB_SYSTEMMODAL);

        return FALSE;
    }

    // Initialize COM for the modern hook.
    BOOL ComInitializedByUs = FALSE;

    HRESULT ComResult = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);

    if (ComResult == S_OK)
    {
        ComInitializedByUs = TRUE;
    }
    else if (ComResult == S_FALSE)
    {
        ComInitializedByUs = TRUE;
    }
    else if (ComResult != RPC_E_CHANGED_MODE)
    {
        if (ModernApplicable)
        {
            MessageBoxW(OwnerWindow,
                L"Failed to initialize COM. The modern Snipping Tool hook could not be installed.",
                L"Error", MB_OK | MB_ICONERROR | MB_SYSTEMMODAL);

            if (!LegacyApplicable)
            {
                return FALSE;
            }
        }
    }

    // Install the modern hook.
    if (ModernApplicable)
    {
        // Check if already hooked with the same path (idempotent).
        SNIPPINGTOOLPACKAGESET JournaledPackages = { 0 };

        ReadHookedPackagesJournal(&JournaledPackages);

        wchar_t JournaledDebugger[MAX_PATH] = { 0 };

        BOOL AlreadyCorrect = FALSE;

        if (JournaledPackages.Count > 0 && ReadRegisteredDebuggerPath(JournaledDebugger))
        {
            if (_wcsicmp(JournaledDebugger, QuotedDebuggerPath) == 0)
            {
                // Check if all current packages are in the journal.
                BOOL AllPresent = TRUE;

                for (UINT32 Installed = 0; Installed < InstalledPackages.Count && AllPresent; Installed++)
                {
                    BOOL Found = FALSE;

                    for (UINT32 Journaled = 0; Journaled < JournaledPackages.Count; Journaled++)
                    {
                        if (_wcsicmp(InstalledPackages.FullNames[Installed], JournaledPackages.FullNames[Journaled]) == 0)
                        {
                            Found = TRUE;

                            break;
                        }
                    }

                    if (!Found) AllPresent = FALSE;
                }

                AlreadyCorrect = AllPresent;
            }
        }

        if (!AlreadyCorrect)
        {
            // Write journal before enabling (crash recovery).
            WriteHookedPackagesJournal(&InstalledPackages);

            WriteRegisteredDebuggerPath(QuotedDebuggerPath);

            BOOL AllSucceeded = TRUE;

            for (UINT32 Index = 0; Index < InstalledPackages.Count; Index++)
            {
                HRESULT EnableResult = EnablePackageDebugging(InstalledPackages.FullNames[Index], QuotedDebuggerPath);

                if (FAILED(EnableResult))
                {
                    MyOutputDebugStringW(L"[%s] Line %d: EnableDebugging failed for %s: 0x%08lX\n",
                        __FUNCTIONW__, __LINE__, InstalledPackages.FullNames[Index], EnableResult);

                    AllSucceeded = FALSE;
                }
            }

            if (!AllSucceeded)
            {
                // Roll back: disable any that succeeded and clear journal.
                for (UINT32 Index = 0; Index < InstalledPackages.Count; Index++)
                {
                    DisablePackageDebugging(InstalledPackages.FullNames[Index]);
                }

                DeleteSnipExJournal();
            }
            else
            {
                ModernApplied = TRUE;
            }
        }
        else
        {
            ModernApplied = TRUE;
        }
    }

    // Install the legacy hook.
    if (LegacyApplicable)
    {
        LSTATUS IfeoStatus = InstallLegacyIfeoHook(QuotedDebuggerPath);

        if (IfeoStatus == ERROR_SUCCESS)
        {
            LegacyApplied = TRUE;
        }
        else
        {
            MyOutputDebugStringW(L"[%s] Line %d: Legacy IFEO hook failed: %ld\n", __FUNCTIONW__, __LINE__, IfeoStatus);
        }
    }

    if (ComInitializedByUs)
    {
        CoUninitialize();
    }

    // Report results.
    BOOL OverallSuccess = TRUE;

    if (LegacyApplicable && !LegacyApplied)
    {
        OverallSuccess = FALSE;
    }

    if (ModernApplicable && !ModernApplied)
    {
        OverallSuccess = FALSE;
    }

    if (OverallSuccess)
    {
        MessageBoxW(OwnerWindow,
            L"The Windows Snipping Tool has been replaced with SnipEx.",
            L"Success", MB_OK | MB_ICONINFORMATION | MB_SYSTEMMODAL);
    }
    else if (LegacyApplied || ModernApplied)
    {
        MessageBoxW(OwnerWindow,
            L"The Snipping Tool was partially replaced. Some launch methods may still open the original app.",
            L"Partial Success", MB_OK | MB_ICONWARNING | MB_SYSTEMMODAL);
    }
    else
    {
        MessageBoxW(OwnerWindow,
            L"Failed to replace the Snipping Tool. Check that SnipEx is running with administrator privileges.",
            L"Error", MB_OK | MB_ICONERROR | MB_SYSTEMMODAL);
    }

    return (LegacyApplied || ModernApplied);
}


BOOL RestoreWindowsSnippingTool(_In_opt_ HWND OwnerWindow)
{
    MyOutputDebugStringW(L"[%s] Line %d: Entered.\n", __FUNCTIONW__, __LINE__);

    BOOL ComInitializedByUs = FALSE;

    HRESULT ComResult = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);

    if (ComResult == S_OK || ComResult == S_FALSE)
    {
        ComInitializedByUs = TRUE;
    }

    BOOL ModernRestored = TRUE;

    BOOL LegacyRestored = TRUE;

    // Disable modern hooks.
    SNIPPINGTOOLPACKAGESET JournaledPackages = { 0 };

    if (ReadHookedPackagesJournal(&JournaledPackages))
    {
        for (UINT32 Index = 0; Index < JournaledPackages.Count; Index++)
        {
            HRESULT DisableResult = DisablePackageDebugging(JournaledPackages.FullNames[Index]);

            if (FAILED(DisableResult))
            {
                MyOutputDebugStringW(L"[%s] Line %d: DisableDebugging failed for %s: 0x%08lX\n",
                    __FUNCTIONW__, __LINE__, JournaledPackages.FullNames[Index], DisableResult);

                ModernRestored = FALSE;
            }
        }
    }

    // Remove legacy hook.
    LSTATUS IfeoStatus = RemoveLegacyIfeoHook();

    if (IfeoStatus != ERROR_SUCCESS)
    {
        if (IfeoStatus == ERROR_ACCESS_DENIED)
        {
            MyOutputDebugStringW(L"[%s] Line %d: IFEO hook belongs to another tool, leaving it alone.\n", __FUNCTIONW__, __LINE__);
        }
        else
        {
            LegacyRestored = FALSE;
        }
    }

    // Clean up journal.
    if (ModernRestored)
    {
        DeleteSnipExJournal();
    }

    if (ComInitializedByUs)
    {
        CoUninitialize();
    }

    if (ModernRestored && LegacyRestored)
    {
        MessageBoxW(OwnerWindow,
            L"The Windows Snipping Tool has been restored.",
            L"Success", MB_OK | MB_ICONINFORMATION | MB_SYSTEMMODAL);
    }
    else
    {
        MessageBoxW(OwnerWindow,
            L"Some hooks could not be fully removed. You may need to run SnipEx as administrator again, "
            L"or manually delete the HKLM\\SOFTWARE\\SnipEx registry key and the "
            L"HKLM\\...\\Image File Execution Options\\SnippingTool.exe\\Debugger value.",
            L"Partial Restore", MB_OK | MB_ICONWARNING | MB_SYSTEMMODAL);
    }

    return (ModernRestored && LegacyRestored);
}


BOOL RelaunchElevatedForHijackRequest(_In_opt_ HWND OwnerWindow, _In_ ELEVATEDHIJACKREQUEST Request)
{
    wchar_t ModulePath[MAX_PATH] = { 0 };

    if (GetModuleFileNameW(NULL, ModulePath, _countof(ModulePath)) == 0)
    {
        return FALSE;
    }

    const wchar_t* Verb = NULL;

    if (Request == ELEVATEDHIJACKREQUEST_REPLACE)
    {
        Verb = REPLACE_SNIPPINGTOOL_VERB;
    }
    else if (Request == ELEVATEDHIJACKREQUEST_RESTORE)
    {
        Verb = RESTORE_SNIPPINGTOOL_VERB;
    }
    else
    {
        return FALSE;
    }

    // Resolve package names now, in the original user's context, and pass them on the command line.
    wchar_t CommandLine[4096] = { 0 };

    wcscpy_s(CommandLine, _countof(CommandLine), Verb);

    SNIPPINGTOOLPACKAGESET InstalledPackages = { 0 };

    if (Request == ELEVATEDHIJACKREQUEST_REPLACE && EnumerateSnippingToolPackages(&InstalledPackages))
    {
        for (UINT32 Index = 0; Index < InstalledPackages.Count; Index++)
        {
            wcscat_s(CommandLine, _countof(CommandLine), L" \"");

            wcscat_s(CommandLine, _countof(CommandLine), InstalledPackages.FullNames[Index]);

            wcscat_s(CommandLine, _countof(CommandLine), L"\"");
        }
    }

    SHELLEXECUTEINFOW ShellExecuteInfo = { sizeof(SHELLEXECUTEINFOW) };

    ShellExecuteInfo.lpVerb       = L"runas";

    ShellExecuteInfo.lpFile       = ModulePath;

    ShellExecuteInfo.lpParameters = CommandLine;

    ShellExecuteInfo.hwnd         = OwnerWindow;

    ShellExecuteInfo.nShow        = SW_NORMAL;

    ShellExecuteInfo.fMask        = SEE_MASK_NOCLOSEPROCESS;

    if (!ShellExecuteExW(&ShellExecuteInfo))
    {
        DWORD Error = GetLastError();

        if (Error != ERROR_CANCELLED)
        {
            MessageBoxW(OwnerWindow,
                L"Failed to relaunch SnipEx with administrator privileges.",
                L"Error", MB_OK | MB_ICONERROR | MB_SYSTEMMODAL);
        }

        return FALSE;
    }

    // Wait for the elevated instance to finish its work.
    if (ShellExecuteInfo.hProcess != NULL)
    {
        WaitForSingleObject(ShellExecuteInfo.hProcess, 30000);

        CloseHandle(ShellExecuteInfo.hProcess);
    }

    return TRUE;
}
