// SnipExTray.h
// Author: Joseph Ryan Ries, 2017-2020
// Implements background/tray mode for intercepting Win+Shift+S on Windows 10 where the
// IPackageDebugSettings hook cannot intercept the shell-handled screen capture overlay.

#pragma once

#define WM_TRAYICON            (WM_APP + 100)

#define WM_HOTKEY_INTERCEPTED  (WM_APP + 101)

#define TRAY_ICON_ID           1

#define REG_HOTKEYINTERCEPTNAME L"HotkeyIntercept"

#define SNIPEX_AUTOSTART_KEY   L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run"

#define SNIPEX_AUTOSTART_VALUE L"SnipEx"

#define SYSCMD_HOTKEY          20008


// Returns TRUE if this system needs the low-level keyboard hook to intercept Win+Shift+S.
// This is the case on Windows 10 where Win+Shift+S is handled by the shell rather than
// by launching the Screen Sketch package.
BOOL IsHotkeyInterceptNeeded(void);

// Installs the low-level keyboard hook and notification area icon.
BOOL StartHotkeyIntercept(void);

// Removes the hook and tray icon.
void StopHotkeyIntercept(void);

// Adds or removes SnipEx from the Windows auto-start (Run key).
BOOL SetAutoStart(_In_ BOOL Enable);

// Returns TRUE if SnipEx is in the auto-start Run key.
BOOL IsAutoStartEnabled(void);

// Call from the main message loop when WM_TRAYICON is received.
void HandleTrayIconMessage(_In_ HWND Window, _In_ WPARAM WParam, _In_ LPARAM LParam);
