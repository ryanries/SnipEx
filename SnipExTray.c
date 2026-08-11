// SnipExTray.c
// Author: Joseph Ryan Ries, 2017-2020
// Background/tray mode with low-level keyboard hook for Win+Shift+S interception on Windows 10.

#ifndef UNICODE
#define UNICODE
#endif
#ifndef _UNICODE
#define _UNICODE
#endif

#pragma warning(push, 0)
#include <windows.h>
#include <shellapi.h>
#include <VersionHelpers.h>
#pragma warning(pop)

#pragma warning(disable: 4820)
#pragma warning(disable: 4710)
#pragma warning(disable: 5045)

#include "SnipEx.h"
#include "SnipExTray.h"

extern HWND gMainWindowHandle;

static HHOOK gKeyboardHook;

static NOTIFYICONDATAW gTrayIconData;

static BOOL gWinKeyDown;

static BOOL gShiftKeyDown;


static LRESULT CALLBACK LowLevelKeyboardProc(_In_ int Code, _In_ WPARAM WParam, _In_ LPARAM LParam)
{
    if (Code == HC_ACTION)
    {
        KBDLLHOOKSTRUCT* KeyData = (KBDLLHOOKSTRUCT*)LParam;

        if (WParam == WM_KEYDOWN || WParam == WM_SYSKEYDOWN)
        {
            if (KeyData->vkCode == VK_LWIN || KeyData->vkCode == VK_RWIN)
            {
                gWinKeyDown = TRUE;
            }
            else if (KeyData->vkCode == VK_SHIFT || KeyData->vkCode == VK_LSHIFT || KeyData->vkCode == VK_RSHIFT)
            {
                gShiftKeyDown = TRUE;
            }
            else if (KeyData->vkCode == 'S' && gWinKeyDown && gShiftKeyDown)
            {
                PostMessageW(gMainWindowHandle, WM_HOTKEY_INTERCEPTED, 0, 0);

                gWinKeyDown = FALSE;

                gShiftKeyDown = FALSE;

                return 1;
            }
        }
        else if (WParam == WM_KEYUP || WParam == WM_SYSKEYUP)
        {
            if (KeyData->vkCode == VK_LWIN || KeyData->vkCode == VK_RWIN)
            {
                gWinKeyDown = FALSE;
            }
            else if (KeyData->vkCode == VK_SHIFT || KeyData->vkCode == VK_LSHIFT || KeyData->vkCode == VK_RSHIFT)
            {
                gShiftKeyDown = FALSE;
            }
        }
    }

    return CallNextHookEx(gKeyboardHook, Code, WParam, LParam);
}


BOOL IsHotkeyInterceptNeeded(void)
{
    // On Windows 11 (build 22000+), our IPackageDebugSettings hook intercepts Win+Shift+S
    // because the Snipping Tool package is activated directly. On Windows 10, the shell
    // handles Win+Shift+S internally and only launches the package after capture completes,
    // so the hook doesn't intercept the shortcut.

    OSVERSIONINFOEXW VersionInfo = { sizeof(OSVERSIONINFOEXW) };

    typedef NTSTATUS (WINAPI *PFN_RtlGetVersion)(PRTL_OSVERSIONINFOW);

    HMODULE Ntdll = GetModuleHandleW(L"ntdll.dll");

    if (Ntdll != NULL)
    {
        PFN_RtlGetVersion RtlGetVersionFn = (PFN_RtlGetVersion)(void*)GetProcAddress(Ntdll, "RtlGetVersion");

        if (RtlGetVersionFn != NULL)
        {
            RtlGetVersionFn((PRTL_OSVERSIONINFOW)&VersionInfo);

            // Windows 11 is build 22000+. Windows 10 is anything below that.
            if (VersionInfo.dwBuildNumber >= 22000)
            {
                return FALSE;
            }
        }
    }

    return TRUE;
}


BOOL StartHotkeyIntercept(void)
{
    if (gKeyboardHook != NULL)
    {
        return TRUE;
    }

    gKeyboardHook = SetWindowsHookExW(WH_KEYBOARD_LL, LowLevelKeyboardProc, GetModuleHandleW(NULL), 0);

    if (gKeyboardHook == NULL)
    {
        return FALSE;
    }

    // Add notification area icon.
    ZeroMemory(&gTrayIconData, sizeof(gTrayIconData));

    gTrayIconData.cbSize = sizeof(NOTIFYICONDATAW);

    gTrayIconData.hWnd = gMainWindowHandle;

    gTrayIconData.uID = TRAY_ICON_ID;

    gTrayIconData.uFlags = NIF_ICON | NIF_MESSAGE | NIF_TIP;

    gTrayIconData.uCallbackMessage = WM_TRAYICON;

    gTrayIconData.hIcon = (HICON)GetClassLongPtrW(gMainWindowHandle, GCLP_HICON);

    wcscpy_s(gTrayIconData.szTip, _countof(gTrayIconData.szTip), L"SnipEx - Win+Shift+S intercept active");

    Shell_NotifyIconW(NIM_ADD, &gTrayIconData);

    return TRUE;
}


void StopHotkeyIntercept(void)
{
    if (gKeyboardHook != NULL)
    {
        UnhookWindowsHookEx(gKeyboardHook);

        gKeyboardHook = NULL;
    }

    Shell_NotifyIconW(NIM_DELETE, &gTrayIconData);
}


BOOL SetAutoStart(_In_ BOOL Enable)
{
    HKEY RunKey = NULL;

    LSTATUS Status = RegOpenKeyExW(HKEY_CURRENT_USER, SNIPEX_AUTOSTART_KEY, 0, KEY_SET_VALUE, &RunKey);

    if (Status != ERROR_SUCCESS)
    {
        return FALSE;
    }

    if (Enable)
    {
        wchar_t ModulePath[MAX_PATH] = { 0 };

        GetModuleFileNameW(NULL, ModulePath, _countof(ModulePath));

        // Add --minimized so SnipEx starts in tray mode.
        wchar_t CommandLine[MAX_PATH + 32] = { 0 };

        swprintf_s(CommandLine, _countof(CommandLine), L"\"%s\" --minimized", ModulePath);

        DWORD DataSize = (DWORD)((wcslen(CommandLine) + 1) * sizeof(WCHAR));

        Status = RegSetValueExW(RunKey, SNIPEX_AUTOSTART_VALUE, 0, REG_SZ, (const BYTE*)CommandLine, DataSize);
    }
    else
    {
        Status = RegDeleteValueW(RunKey, SNIPEX_AUTOSTART_VALUE);

        if (Status == ERROR_FILE_NOT_FOUND)
        {
            Status = ERROR_SUCCESS;
        }
    }

    RegCloseKey(RunKey);

    return (Status == ERROR_SUCCESS);
}


BOOL IsAutoStartEnabled(void)
{
    HKEY RunKey = NULL;

    if (RegOpenKeyExW(HKEY_CURRENT_USER, SNIPEX_AUTOSTART_KEY, 0, KEY_READ, &RunKey) != ERROR_SUCCESS)
    {
        return FALSE;
    }

    DWORD ValueType = 0;

    LSTATUS Status = RegQueryValueExW(RunKey, SNIPEX_AUTOSTART_VALUE, NULL, &ValueType, NULL, NULL);

    RegCloseKey(RunKey);

    return (Status == ERROR_SUCCESS);
}


void HandleTrayIconMessage(_In_ HWND Window, _In_ WPARAM WParam, _In_ LPARAM LParam)
{
    UNREFERENCED_PARAMETER(WParam);

    if (LOWORD(LParam) == WM_LBUTTONDBLCLK)
    {
        ShowWindow(Window, SW_RESTORE);

        SetForegroundWindow(Window);
    }
    else if (LOWORD(LParam) == WM_RBUTTONUP)
    {
        HMENU PopupMenu = CreatePopupMenu();

        if (PopupMenu != NULL)
        {
            AppendMenuW(PopupMenu, MF_STRING, 1, L"Show SnipEx");

            AppendMenuW(PopupMenu, MF_STRING, 2, L"Exit");

            POINT CursorPosition = { 0 };

            GetCursorPos(&CursorPosition);

            SetForegroundWindow(Window);

            UINT Command = (UINT)TrackPopupMenu(PopupMenu, TPM_RETURNCMD | TPM_NONOTIFY, CursorPosition.x, CursorPosition.y, 0, Window, NULL);

            DestroyMenu(PopupMenu);

            if (Command == 1)
            {
                ShowWindow(Window, SW_RESTORE);

                SetForegroundWindow(Window);
            }
            else if (Command == 2)
            {
                PostMessageW(Window, WM_CLOSE, 0, 0);
            }
        }
    }
}
