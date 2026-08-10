// SnipExHijack.h
// Author: Joseph Ryan Ries, 2017-2020
// Taking over the Windows Snipping Tool, on both classic and MSIX-packaged versions of Windows.

#pragma once

// The modern Snipping Tool ships inside an MSIX package that is still named "Screen Sketch" for
// historical reasons, so the family name below covers Snip & Sketch on Windows 10 and the Snipping
// Tool on Windows 11.
#define SNIPPINGTOOL_PACKAGE_FAMILY_NAME  L"Microsoft.ScreenSketch_8wekyb3d8bbwe"

#define SNIPEX_MACHINE_KEY_PATH           L"SOFTWARE\\SnipEx"

#define REPLACE_SNIPPINGTOOL_VERB         L"--replace-snipping-tool"

#define RESTORE_SNIPPINGTOOL_VERB         L"--restore-snipping-tool"

#define MAX_SNIPPINGTOOL_PACKAGES         16

#define MAX_PACKAGE_FULL_NAME_LENGTH      127


// Whether the Windows Snipping Tool currently launches SnipEx instead of itself.
typedef enum SNIPPINGTOOLHOOKSTATE
{
	// The Snipping Tool launches normally.
	SNIPPINGTOOLHOOKSTATE_NOTREPLACED,

	// At least one launch surface currently redirects to SnipEx.
	SNIPPINGTOOLHOOKSTATE_REPLACED,

	// Some other program already registered itself as the classic Snipping Tool's debugger.
	// SnipEx refuses to overwrite or delete another program's hook.
	SNIPPINGTOOLHOOKSTATE_FOREIGNDEBUGGER

} SNIPPINGTOOLHOOKSTATE;


// What happened when Windows handed a Snipping Tool activation over to this SnipEx process.
typedef enum HIJACKEDLAUNCHRESULT
{
	// This process was started normally, not by the Snipping Tool package hook.
	HIJACKEDLAUNCHRESULT_NOTHIJACKEDLAUNCH,

	// The suspended Snipping Tool process was verified and terminated. SnipEx takes over.
	HIJACKEDLAUNCHRESULT_SNIPPINGTOOLTERMINATED,

	// The process named on the command line was not a Snipping Tool process, so it was left alone.
	HIJACKEDLAUNCHRESULT_REJECTEDFOREIGNPROCESS,

	// The Snipping Tool process was identified but could not be terminated. It may still be suspended.
	HIJACKEDLAUNCHRESULT_TERMINATIONFAILED

} HIJACKEDLAUNCHRESULT;


// An operation that an unelevated SnipEx asked this elevated SnipEx to carry out.
typedef enum ELEVATEDHIJACKREQUEST
{
	ELEVATEDHIJACKREQUEST_NONE,

	ELEVATEDHIJACKREQUEST_REPLACE,

	ELEVATEDHIJACKREQUEST_RESTORE

} ELEVATEDHIJACKREQUEST;


// A set of Snipping Tool package full names. Package full names are per-user and change every time
// the package is updated, so they are always resolved fresh rather than hard-coded.
typedef struct SNIPPINGTOOLPACKAGESET
{
	UINT32  Count;

	wchar_t FullNames[MAX_SNIPPINGTOOL_PACKAGES][MAX_PACKAGE_FULL_NAME_LENGTH + 1];

} SNIPPINGTOOLPACKAGESET;


// Reports whether the Snipping Tool currently launches SnipEx, so the caller can decide whether to
// offer "Replace" or "Restore". Derived from machine state rather than from any in-memory flag, so a
// freshly downloaded copy of SnipEx sitting in a different folder still reports the truth.
SNIPPINGTOOLHOOKSTATE GetSnippingToolHookState(void);

// Call this before doing anything else in WinMain. When Windows launches SnipEx in place of the
// Snipping Tool it appends "-p <ProcessId> -tid <ThreadId>" and creates the real Snipping Tool
// process suspended, waiting for a debugger. This verifies that the process really does belong to
// the Snipping Tool package and then terminates it, so it is never left suspended forever.
HIJACKEDLAUNCHRESULT TerminateHijackedSnippingToolProcess(void);

// Returns the operation that an unelevated SnipEx asked this process to perform after elevating.
ELEVATEDHIJACKREQUEST GetElevatedHijackRequest(void);

// Makes the Windows Snipping Tool launch SnipEx. Requires elevation. Reports its own results to the
// user. Returns TRUE only if the Snipping Tool now launches SnipEx.
BOOL ReplaceSnippingToolWithSnipEx(_In_opt_ HWND OwnerWindow);

// Undoes ReplaceSnippingToolWithSnipEx. Requires elevation. Reports its own results to the user.
BOOL RestoreWindowsSnippingTool(_In_opt_ HWND OwnerWindow);

// Relaunches SnipEx elevated to carry out Request. Package full names are resolved here, in the
// original user's context, because package registrations are per-user and the elevated process may
// belong to a different account entirely.
BOOL RelaunchElevatedForHijackRequest(_In_opt_ HWND OwnerWindow, _In_ ELEVATEDHIJACKREQUEST Request);
