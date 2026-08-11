SnipEx aims to be a slightly better version of Windows' built-in Snipping Tool.

The project began when I grew tired of trying to hilight text in a screenshot, and the freehand-style hilighter would wander all over the image, making the screenshot look as if a small child had taken a hilighter to it. 
I wanted a hilighter that would travel in a perfectly straight line, for professional-looking hilighted text.

So I reinvented the wheel and the project grew from there. 

Contact ryanries09@gmail.com for any questions or bug reports or feature requests, or hit me up on Twitter @JosephRyanRies.


Shortcut keys:
-------------
 - Escape cancels an active screen capture, or exits the app.
 - N = New Snip
 - D = Delay Snip
 - S = Save Snip
 - C = Copy Snip
 - H = Hilight Snip (Toggle High-Contrast Hilighter)
 - B = Draw rectangle
 - A = Draw arrow
 - R = Redact (black marker)
 - T = Text
 - Ctrl+Z = Undo the last change


Right click on the tool buttons to cycle through different colors.

Right click on the Text button to change the font, size and color.
 
Pictures:
------------- 

The toolbar:

![SnipEx 1](./pictures/snipex1.png) 

The highlighter:

![SnipEx 2](./pictures/snipex2.png)

The highlighter works via correlation function. The "brighter" a pixel is, where "brightness" is defined as the sum of R, G and B, the more darkly it will be highlighted. So black pixels won't be touched, while white pixels will be highlighted strongly.

Replace Windows Snipping Tool with SnipEx:

![SnipEx 3](./pictures/replace.png)

On older versions of Windows (7, 8, 8.1, and early builds of Windows 10), SnipEx uses the classic Image File Execution Options (IFEO) "Debugger" registry hack to intercept SnippingTool.exe launches -- the same technique Sysinternals Process Explorer uses to replace Task Manager.

On Windows 10 (recent builds) and Windows 11, Microsoft converted the Snipping Tool into an MSIX-packaged app (Microsoft.ScreenSketch). The classic IFEO hook no longer intercepts it because the packaged app is not launched through CreateProcess in the traditional way. SnipEx now uses the IPackageDebugSettings COM interface to register itself as the package's debugger, which intercepts all Snipping Tool launch surfaces: the Start menu tile, the Win+Shift+S hotkey, the Print Screen key, the app execution alias, and the ms-screenclip/ms-screensketch protocol activations.

Both mechanisms are applied simultaneously when both are applicable (e.g. on Windows 10 builds that have both the classic exe and the packaged app), so no launch surface is missed.

**Security note:** When you click "Replace Windows Snipping Tool with SnipEx," the path to SnipEx.exe is registered as a machine-wide launch hook. If SnipEx is running from a user-writable location (Downloads, Desktop, a USB stick, etc.), a standard user could replace the binary and gain code execution as any other user who opens the Snipping Tool. To prevent this, SnipEx will warn you and offer to copy itself into Program Files before registering the hook. If you decline the copy, the Replace operation is aborted.

With other snipping tools, this looks bad:

![Ugly](./pictures/ugly.png)

This looks better:

![Better](./pictures/better.png)




History:
-------
Update 8/10/2026:
- Version 1.4.31
  - Fixed the "Replace Windows Snipping Tool with SnipEx" feature on Windows 10 and Windows 11.
    The Snipping Tool became an MSIX-packaged app starting in Windows 10, and the classic Image File
    Execution Options (IFEO) hook no longer intercepts it. SnipEx now uses the IPackageDebugSettings
    COM interface to hook the modern packaged Snipping Tool while keeping the legacy IFEO method for
    older versions of Windows where the classic SnippingTool.exe still exists.
  - The Replace feature now copies SnipEx.exe into Program Files if it is running from an unprotected
    location (e.g. Downloads or Desktop), preventing a privilege escalation where a standard user could
    swap the registered binary.
  - Fixed a pre-existing bug where the IFEO Debugger value was written without a terminating NUL and
    without quoting, which broke paths containing spaces.
  - Manual recovery: if SnipEx is uninstalled while the hook is active, run the following from an
    elevated command prompt to restore the Snipping Tool:
      reg delete "HKLM\SOFTWARE\SnipEx" /f
      reg delete "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\SnippingTool.exe" /v Debugger /f

Update 8/11/2026:
- Version 1.4.33
  - Added "Intercept Win+Shift+S (stay in background)" option for Windows 10 users. On Windows 10,
    the Win+Shift+S shortcut is handled directly by the shell (explorer.exe) rather than by launching
    the Snipping Tool package, so the IPackageDebugSettings hook alone cannot intercept it. When this
    option is enabled, SnipEx installs a low-level keyboard hook to catch Win+Shift+S before the shell
    sees it, minimizes to the notification area (system tray), and registers itself to auto-start with
    Windows. This option is only shown on Windows 10 systems where it is needed; on Windows 11 the
    package hook already intercepts Win+Shift+S without needing a background process.
  - Added "Automatically save screen captures" option. When enabled, a folder picker dialog lets
    you choose a destination folder, and every new snip is automatically saved as a timestamped PNG
    file (e.g. SnipEx_2026-08-11_14-30-05-123.png) in that folder.
  - Fixed an issue where SnipEx could spawn behind other windows on some PCs when launched via the
    Snipping Tool package hook (Win+Shift+S / Start menu).

Update 9/30/2020:
- Version 1.3.30
  - Removed bad "custom DPI/scaling" code and replaced it with a simpler algorithm that should solve for any custom DPI.
  
Update 8/24/2020:
- Version 1.3.27
  - Added a "Automatically copy snip to clipboard" option to the drop-down menu.
  
Update 3/19/2020:
- Version 1.2.25
  - Added a new Text button!
  - Added a "Remember previously used tool" option to the drop-down menu.


Update 2/26/2020:
- Version 1.1.12
  - Added new colors for the rectangle, hilighter, and arrow tools! Just right-click the buttons to cycle through different colors!
  - Added an "Undo (Ctrl+Z)" to the drop-down menu. Ctrl+Z always worked, but the menu item might help people who don't know the keyboard shortcut.

Update 7/12/2017:
- Version 1.0.11
  - Added support for custom scaling levels up to 200%, which is the default scaling level of a Surface Pro 4. (High-DPI)
  

Update 7/11/2017:
- Version 1.0.10
  - Removed the "toggleable" high-contrast hilighter and just made that the only hilighter. Improved the hilighter quality to probably about as good as it can be, by using an correlation function. The darker a pixel is, the less hilighted it gets.
  - Added support for custom Windows GUI scaling up to 175%. If you use scaling > 175%, let me know and I will add support for you. You must be using a football stadium display as a monitor or something.


Update 7/4/2017:
- Version 1.0.9
  - Added a "high-contrast" hilighter. Click the hilighter tool multiple times to toggle it. It helps make hilighted text look better under most conditions.

  
Update 6/21/2017:
- Version 1.0.8
  - Added an elegant drop shadow effect which can be toggled on and off. You can find it in the drop-down menu in the top-left corner of the app. Your preference is saved in the registry.

  
Update 6/17/2017:
- Version 1.0.7
  - Added support for saving images in PNG format.

  
Update 6/13/2017:
- Version 1.0.6
  - Fixed an issue where if the user set a custom scaling factor, the title bar would be larger than expected and cut into the window's client area, clipping off the bottom of the buttons.

  
Update 6/12/2017:
- Version 1.0.5
  - Fixed button font size and face to avoid issues when user's DEFAULT_GUI_FONT is larger than expected.
  
  
Update 6/11/2017:
- Version 1.0.4
  - Added "Redact" drawing tool.
  - Fixed a bug where I assumed that the top-left coordinate of the user's viewing area was always (0,0), but that is not true if the user has multiple monitors and they are arranged in an unusual order.
  - Signed the binary.


Update 5/30/2017:
- Version 1.0.3
  - Added the "Arrow" drawing tool.
  - Added a UAC icon to the "Replace Snipping Tool" menu item.

  
Update 5/24/2017:
- Version 1.0.1
  - Added a "Replace Snipping Tool with SnipEx" and "Restore Snipping Tool" option from the application's system menu in the top-left corner. This works like Sysinternals' Process Explorer's "Replace Task Manager".
  - Added a missing call to UpdateWindow that should make drawing with the hilighter look a little smoother. Increased the amount of pixels that can be drawn over before the hilighted area begins to become more opaque.
 
 
Contact ryanries09@gmail.com for any questions or bug reports or feature requests, or hit me up on Twitter @JosephRyanRies.