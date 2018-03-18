// SnipEx.c
// Author: Joseph Ryan Ries, 2017
// The snip.exe that comes bundled with Microsoft Windows is *almost* good enough. So I made one just a little better.

#pragma warning(push, 0)			// Temporarily disable warnings in header files over which I have no control.
#include <windowsx.h>				// Windows API
#include <ShObjIdl.h>				// For save dialog
#include <ShellScalingApi.h>		// For detecting monitor DPI (Win 8.1 or above)
#include <initguid.h>				// For doing stuff with GUIDs (Needed for COM interop)
#pragma warning(pop)

#pragma comment(lib, "Msimg32.lib") // For TransparentBlt
#pragma comment(lib, "Shcore.lib")	// For detecting monitor DPI (Win 8.1 or above)

#pragma warning(disable: 4820)		// Disable compiler warning about padding bytes being added to structs
#pragma warning(disable: 4710)		// Disable compiler warning about functions not being inlined

#include <stdio.h>					// For doing stuff with strings
#include <math.h>					// Needed to introduce some math to draw the arrow head for the arrow tool

#include "resource.h"				// Images, cursors, etc.
#include "SnipEx.h"					// My custom definitions
#include "ButtonDefs.h"				// Buttons!
#include "GdiPlusInterop.h"			// Stuff to make GDI+ work in pure C

APPSTATE gAppState = APPSTATE_BEFORECAPTURE;

// Set this to FALSE to exit the app immediately.
BOOL gMainWindowIsRunning;
                           
HWND gMainWindowHandle;
HWND gCaptureWindowHandle;

const INT8 gStartingDelayCountdown = 6;
INT8 gCurrentDelayCountdown = 6;

UINT16 gDisplayWidth;
UINT16 gDisplayHeight;
INT16  gDisplayLeft;
INT16  gDisplayTop;

UINT16 gStartingMainWindowWidth  = 593;
UINT16 gStartingMainWindowHeight = 92;

HBITMAP gCleanScreenShot;

// For use during drawing.
HBITMAP gScratchBitmap;

RECT gCaptureSelectionRectangle;

int gCaptureWidth;
int gCaptureHeight;

BOOL gLeftMouseButtonIsDown;

// An array of bitmap states that we can revert back to, so we can undo changes. ctrl-z.
// Snip state 0 will always be the unaltered clip right as the user first took it.

UINT8   gCurrentSnipState;
HBITMAP gSnipStates[32];

HBITMAP gUACIcon;

BOOL gShouldAddDropShadow;


// Application entry-point.
// MSDN says that if WinMain fails before reaching the message loop, we should return zero.
int CALLBACK WinMain(_In_ HINSTANCE Instance, _In_opt_ HINSTANCE PreviousInstance, _In_ LPSTR CommandLine, _In_ int WindowShowCode)
{
	MyOutputDebugStringA("[%s] Line %d: Entered.\n", __FUNCTION__, __LINE__);

	UNREFERENCED_PARAMETER(PreviousInstance);
	UNREFERENCED_PARAMETER(CommandLine);
	UNREFERENCED_PARAMETER(WindowShowCode);	

	// NOTE: This width and height are the bounding box that includes all of the user's monitors.
	// I have only tested with two monitors of equal dimensions, so this may not work correctly
	// if the user has several irregularly-arranged monitors. Also, if the monitors are arranged
	// in reverse order, the left-most X coordinate may be e.g. negative 1920! In other words, 0,0 may not 
	// necessarily be the top-left corner of the user's viewing area.
	gDisplayWidth  = (UINT16)GetSystemMetrics(SM_CXVIRTUALSCREEN);
	gDisplayHeight = (UINT16)GetSystemMetrics(SM_CYVIRTUALSCREEN);
	gDisplayLeft   = (INT16)GetSystemMetrics(SM_XVIRTUALSCREEN);
	gDisplayTop    = (INT16)GetSystemMetrics(SM_YVIRTUALSCREEN);

	if (gDisplayWidth == 0 || gDisplayHeight == 0)
	{
		MyMessageBoxA(NULL, "Error", MB_OK | MB_ICONERROR | MB_SYSTEMMODAL, "Failed to retrieve display area! Error code: 0x%lx", GetLastError());
		return(0);
	}

	MyOutputDebugStringA("[%s] Line %d: Detected a screen area of %dx%d.\n", __FUNCTION__, __LINE__, gDisplayWidth, gDisplayHeight);

	// If the user had a custom scaling/DPI level set, this app was originally not high-DPI aware, and so what would happen
	// is that the buttons would get cropped as the non-client area of the window got bigger, and the screen would "zoom in"
	// whenever the user clicked "New"!
	if (AdjustForCustomScaling() == FALSE)
	{
		return(0);
	}

	WNDCLASSEX CaptureWindowClass = { sizeof(WNDCLASSEX) };

	CaptureWindowClass.style         = CS_HREDRAW | CS_VREDRAW;
	CaptureWindowClass.hInstance     = Instance;
	CaptureWindowClass.lpszClassName = "SnipExCaptureWindowClass";
	CaptureWindowClass.hCursor       = LoadCursor(NULL, IDC_CROSS);
	CaptureWindowClass.hbrBackground = GetSysColorBrush(COLOR_DESKTOP);
	CaptureWindowClass.lpfnWndProc   = CaptureWindowCallback;

	if (RegisterClassEx(&CaptureWindowClass) == 0)
	{
		MyMessageBoxA(NULL, "Error", MB_OK | MB_ICONERROR | MB_SYSTEMMODAL, "Failed to register CaptureWindowClass with error 0x%lx!", GetLastError());
		return(0);
	}

	// This window will capture the entire display surface, including multiple monitors. It will then display an exact
	// screenshot of the desktop over the real desktop. Then the user will be able to select a subsection of that using
	// the mouse. This could also be used to play a prank on somebody and make them think their desktop was hung.
	gCaptureWindowHandle = CreateWindowEx(
		WS_EX_TOPMOST | WS_EX_TOOLWINDOW,	// No taskbar icon
		CaptureWindowClass.lpszClassName,
		"",									// No title
		0,									// Not visible
		CW_USEDEFAULT,
		CW_USEDEFAULT,
		0,									// Will size it later
		0,
		NULL,
		NULL,
		Instance,
		NULL);

	if (gCaptureWindowHandle == NULL)
	{
		MyMessageBoxA(NULL, "Error", MB_OK | MB_ICONERROR | MB_SYSTEMMODAL, "Failed to create Capture Window with error 0x%lx!", GetLastError());		
		return(0);
	}

	// Remove all window style, including title bar. We're trying to make the "capture window" indistinguishable from the real desktop that lay underneath it.
	if (SetWindowLongPtr(gCaptureWindowHandle, GWL_STYLE, 0) == 0)
	{
		MyMessageBoxA(NULL, "Error", MB_OK | MB_ICONERROR | MB_SYSTEMMODAL, "SetWindowLongPtr failed with error 0x%lx!", GetLastError());
		return(0);
	}

	// The main window class that appears when the application is first launched.
	WNDCLASSEX MainWindowClass = { sizeof(WNDCLASSEX) };

	MainWindowClass.style         = CS_HREDRAW | CS_VREDRAW;
	MainWindowClass.hInstance     = Instance;
	MainWindowClass.lpszClassName = "SnipExWindowClass";
	MainWindowClass.hCursor       = LoadCursor(NULL, IDC_ARROW);
	MainWindowClass.hbrBackground = GetSysColorBrush(COLOR_WINDOW);
	MainWindowClass.lpfnWndProc   = MainWindowCallback;
	MainWindowClass.hIcon         = LoadIcon(Instance, MAKEINTRESOURCE(IDI_MAINAPPICON));
	MainWindowClass.hIconSm       = LoadIcon(Instance, MAKEINTRESOURCE(IDI_MAINAPPICON));

	if (RegisterClassEx(&MainWindowClass) == 0)
	{
		MyMessageBoxA(NULL, "Error", MB_OK | MB_ICONERROR | MB_SYSTEMMODAL, "Failed to register MainWindowClass with error 0x%lx!", GetLastError());		
		return(0);
	}

	gMainWindowHandle = CreateWindowEx(
		0, //WS_EX_TOPMOST,
		MainWindowClass.lpszClassName,
		"SnipEx",
		WS_VISIBLE | (WS_OVERLAPPEDWINDOW ^ (WS_MAXIMIZEBOX | WS_THICKFRAME)),
		CW_USEDEFAULT,
		CW_USEDEFAULT,
		gStartingMainWindowWidth,
		gStartingMainWindowHeight,
		NULL,
		NULL,
		Instance,
		NULL);	

	if (gMainWindowHandle == NULL)
	{
		MyMessageBoxA(NULL, "Error", MB_OK | MB_ICONERROR | MB_SYSTEMMODAL, "Failed to create Main Window with error 0x%lx!", GetLastError());
		return(0);
	}

	#if _DEBUG
	char TitleBarBuffer[64] = { 0 };

	GetWindowText(gMainWindowHandle, TitleBarBuffer, sizeof(TitleBarBuffer));
	strcat_s(TitleBarBuffer, sizeof(TitleBarBuffer), " - *DEBUG BUILD*");
	SetWindowText(gMainWindowHandle, TitleBarBuffer);
	MyOutputDebugStringA("[%s] Line %d: Setting window text for *DEBUG BUILD*.\n", __FUNCTION__, __LINE__);
	#endif

	// Create all the buttons.
	for (UINT8 Counter = 0; Counter < ArrayCount(gButtons); Counter++)
	{
		HWND ButtonHandle = CreateWindowEx(
			0,
			"BUTTON",
			gButtons[Counter]->Caption,
			BS_OWNERDRAW | WS_VISIBLE | WS_CHILD,
			gButtons[Counter]->Rectangle.left,
			gButtons[Counter]->Rectangle.top,
			gButtons[Counter]->Rectangle.right - gButtons[Counter]->Rectangle.left,
			gButtons[Counter]->Rectangle.bottom,
			gMainWindowHandle,
			(HMENU)gButtons[Counter]->Id,
			(HINSTANCE)GetWindowLongPtr(gMainWindowHandle, GWLP_HINSTANCE),
			NULL);

		if (ButtonHandle == NULL)
		{
			MyMessageBoxA(NULL, "Error", MB_OK | MB_ICONERROR | MB_SYSTEMMODAL, "Failed to create \"%s\" button with error 0x%lx!", gButtons[Counter]->Caption, GetLastError());
			return(0);
		}

		gButtons[Counter]->Handle = ButtonHandle;

		if (gButtons[Counter]->EnabledIconId > 0)
		{			
			gButtons[Counter]->EnabledIcon = (HBITMAP)LoadImage(Instance, MAKEINTRESOURCE(gButtons[Counter]->EnabledIconId), IMAGE_BITMAP, 0, 0, 0);

			if (gButtons[Counter]->EnabledIcon == NULL)
			{
				MyOutputDebugStringA("[%s] Line %d: Loading resource %d failed! Error: 0x%lx\n", __FUNCTION__, __LINE__, gButtons[Counter]->EnabledIconId, GetLastError());
			}
		}

		if (gButtons[Counter]->DisabledIconId > 0)
		{
			gButtons[Counter]->DisabledIcon = (HBITMAP)LoadImage(Instance, MAKEINTRESOURCE(gButtons[Counter]->DisabledIconId), IMAGE_BITMAP, 0, 0, 0);

			if (gButtons[Counter]->DisabledIcon == NULL)
			{
				MyOutputDebugStringA("[%s] Line %d: Loading resource %d failed! Error: 0x%lx\n", __FUNCTION__, __LINE__, gButtons[Counter]->DisabledIconId, GetLastError());
			}
		}

		if (gButtons[Counter]->CursorId > 0)
		{
			gButtons[Counter]->Cursor = LoadCursor(Instance, MAKEINTRESOURCE(gButtons[Counter]->CursorId));

			if (gButtons[Counter]->Cursor == NULL)
			{
				MyOutputDebugStringA("[%s] Line %d: Loading resource %d failed! Error: 0x%lx\n", __FUNCTION__, __LINE__, gButtons[Counter]->CursorId, GetLastError());
			}
		}

		MyOutputDebugStringA("[%s] Line %d: %s button created.\n", __FUNCTION__, __LINE__, gButtons[Counter]->Caption);
	}

	// Add custom menu items to the form's system menu in the top left

	AddReplaceSnippingToolMenuItem(Instance);
	
	AddDropShadowToolMenuItem();


	gMainWindowIsRunning = TRUE;

	MSG MainWindowMessage    = { 0 };
	MSG CaptureWindowMessage = { 0 };	

	MyOutputDebugStringA("[%s] Line %d: Setup is finished. Entering message loop.\n", __FUNCTION__, __LINE__);

	//////////////////
	//              //
	// MESSAGE LOOP //
	//              //
	//////////////////

	while (gMainWindowIsRunning)
	{
		// Drain message queue for main window.
		while (PeekMessage(&MainWindowMessage, gMainWindowHandle, 0, 0, PM_REMOVE))
		{
			DispatchMessage(&MainWindowMessage);
		}

		// Drain message queue for capture window.
		while (PeekMessage(&CaptureWindowMessage, gCaptureWindowHandle, 0, 0, PM_REMOVE))
		{
			DispatchMessage(&CaptureWindowMessage);
		}

		if ((gAppState == APPSTATE_AFTERCAPTURE) && gLeftMouseButtonIsDown)
		{
			if (GetAsyncKeyState(VK_LBUTTON) == 0)
			{
				gLeftMouseButtonIsDown = FALSE;
				SendMessage(gMainWindowHandle, WM_LBUTTONUP, 0, 0);
			}
		}
			
		Sleep(1); // Could be anywhere from 0.5ms to 15.6ms
	}

	return(0);
}

// Handles window messages for the main window.
LRESULT CALLBACK MainWindowCallback(_In_ HWND Window, _In_ UINT Message, _In_ WPARAM WParam, _In_ LPARAM LParam)
{
	LRESULT Result = 0;
	
	static BOOL CurrentlyDrawing;
	static POINT MousePosWhenDrawingStarted;
	static POINT PreviousMousePos;
	static SMALLPOINT HilighterPixelsAlreadyDrawn[32768] = { 0 }; // 128k of memory

	static UINT16 HilighterPixelsAlreadyDrawnCounter = 0;

	switch (Message)
	{
		case WM_CLOSE:
		{
			PostQuitMessage(0);
			gMainWindowIsRunning = FALSE;
			break;
		}
		case WM_KEYDOWN:
		{
			// Ctrl+Z, Undo
			if ((WParam == 0x5A) && GetKeyState(VK_CONTROL) && (gAppState == APPSTATE_AFTERCAPTURE))
			{
				if (gCurrentSnipState > 0)
				{
					gCurrentSnipState--;

					MyOutputDebugStringA("[%s] Line %d: Deleting gSnipStates[gCurrentSnipState + 1]\n", __FUNCTION__, __LINE__);

					memset(HilighterPixelsAlreadyDrawn, 0, sizeof(HilighterPixelsAlreadyDrawn));
					HilighterPixelsAlreadyDrawnCounter = 0;

					if (DeleteObject(gSnipStates[gCurrentSnipState + 1]) == 0)
					{
						MyOutputDebugStringA("[%s] Line %d: DeleteObject failed!\n", __FUNCTION__, __LINE__);
					}
				}
			}

			// Allow Escape to terminate the app
			if ((WParam == VK_ESCAPE) && ((gAppState == APPSTATE_AFTERCAPTURE) || (gAppState == APPSTATE_BEFORECAPTURE)))
			{
				MyOutputDebugStringA("[%s] Line %d: Escape key was pressed. Exiting application.\n", __FUNCTION__, __LINE__);
				PostQuitMessage(0);
				gMainWindowIsRunning = FALSE;
			}

			for (UINT8 Counter = 0; Counter < ArrayCount(gButtons); Counter++)
			{
				if (WParam == gButtons[Counter]->Hotkey && gButtons[Counter]->Enabled == TRUE)
				{								
					gButtons[Counter]->State = BUTTONSTATE_PRESSED;
				}
			}

			if ((WParam == VK_ESCAPE) && (gAppState == APPSTATE_DELAYCOOKING))
			{
				// Cancel a current delay countdown.
				MyOutputDebugStringA("[%s] Line %d: Escape key was pressed during delay countdown. Cancelling countdown.\n", __FUNCTION__, __LINE__);
				gAppState = APPSTATE_BEFORECAPTURE;
				KillTimer(gMainWindowHandle, DELAY_TIMER);
				gCurrentDelayCountdown = gStartingDelayCountdown;
			}

			InvalidateRect(Window, NULL, FALSE);
			break;
		}
		case WM_KEYUP:
		{
			for (UINT8 Counter = 0; Counter < ArrayCount(gButtons); Counter++)
			{
				if (WParam == gButtons[Counter]->Hotkey && gButtons[Counter]->Enabled == TRUE)
				{
					MyOutputDebugStringA("[%s] Line %d: Hotkey released, executing button '%s'\n", __FUNCTION__, __LINE__, gButtons[Counter]->Caption);
					SendMessage(gMainWindowHandle, WM_COMMAND, gButtons[Counter]->Id, 0);
					if (gButtons[Counter]->SelectedTool == FALSE)
					{
						gButtons[Counter]->State = BUTTONSTATE_NORMAL;
					}
				}
			}
			break;
		}
		case WM_SETCURSOR:
		{
			// To keep Windows from automatically trying to set my cursor for me.
			break;
		}
		case WM_LBUTTONDOWN:
		{
			MyOutputDebugStringA("[%s] Line %d: Left mouse button down.\n", __FUNCTION__, __LINE__);

			gLeftMouseButtonIsDown = TRUE;			
			
			if (CurrentlyDrawing == FALSE && gAppState == APPSTATE_AFTERCAPTURE)
			{
				BOOL DrawingToolSelected = FALSE;

				for (UINT8 Counter = 0; Counter < ArrayCount(gButtons); Counter++)
				{
					if (gButtons[Counter]->SelectedTool == TRUE)
					{
						DrawingToolSelected = TRUE;
					}
				}

				if (DrawingToolSelected == FALSE)
				{
					MyOutputDebugStringA("[%s] Line %d: No drawing tool is selected. Won't start drawing.\n", __FUNCTION__, __LINE__);
					break;
				}

				POINT Mouse = { 0 };

				GetCursorPos(&Mouse);

				ScreenToClient(gMainWindowHandle, &Mouse);

				MousePosWhenDrawingStarted = Mouse;

				if (gCurrentSnipState >= ArrayCount(gSnipStates) - 1)
				{
					MessageBox(gMainWindowHandle, "Maximum number of changes exceeded. Ctrl+Z to undo a change first.", "Error", MB_OK | MB_ICONERROR | MB_SYSTEMMODAL);
					break;
				}

				if (Mouse.x < 2 || Mouse.y < 52)
				{
					MyOutputDebugStringA("[%s] Line %d: Mouse was not over the screen capture area. Will not start drawing.\n", __FUNCTION__, __LINE__);
					break;
				}

				CurrentlyDrawing = TRUE;			

				if (gScratchBitmap != NULL)
				{
					MyOutputDebugStringA("[%s] Line %d: gScratchBitmap was not null, but it was expected to be!\n", __FUNCTION__, __LINE__);
				}

				gScratchBitmap = CopyImage(gSnipStates[gCurrentSnipState], IMAGE_BITMAP, 0, 0, 0);

				gCurrentSnipState++;
			}

			break;
		}
		case WM_LBUTTONUP:
		{
			MyOutputDebugStringA("[%s] Line %d: Left mouse button up.\n", __FUNCTION__, __LINE__);

			gLeftMouseButtonIsDown = FALSE;
			CurrentlyDrawing = FALSE;

			MousePosWhenDrawingStarted.x = 0;
			MousePosWhenDrawingStarted.y = 0;

			PreviousMousePos.x = 0;
			PreviousMousePos.y = 0;

			if (gScratchBitmap != NULL)
			{
				gSnipStates[gCurrentSnipState] = CopyImage(gScratchBitmap, IMAGE_BITMAP, 0, 0, 0);
				if (DeleteObject(gScratchBitmap) == 0)
				{
					MyOutputDebugStringA("[%s] Line %d: DeleteObject failed! Was the bitmap still selected into a DC?\n", __FUNCTION__, __LINE__);
				}
				gScratchBitmap = NULL;
			}

			break;
		}
		case WM_MOUSEMOVE:
		case WM_NCMOUSEMOVE:		
		{
			// Handle drawing first			

			if (CurrentlyDrawing)
			{
				if (gScratchBitmap == NULL)
				{
					MyOutputDebugStringA("[%s] Line %d: gScratchBitmap is NULL\n", __FUNCTION__, __LINE__);
					break;
				}				

				if (gHilighterButton.SelectedTool == TRUE)
				{
					POINT Mouse = { 0 };

					GetCursorPos(&Mouse);					

					ScreenToClient(gMainWindowHandle, &Mouse);

					// Adjust for snip area, maintain Y axis
					Mouse.x -= 7;
					Mouse.y = MousePosWhenDrawingStarted.y - 66;

					HDC ScratchDC = CreateCompatibleDC(NULL);
					HDC HilightDC = CreateCompatibleDC(NULL);

					SelectObject(ScratchDC, (HBITMAP)gScratchBitmap);					

					const BYTE HilightBits[] = { 0x00, 0xFF, 0xFF, 0xFF };
					HBITMAP HilightPixel = CreateBitmap(1, 1, 1, 32, HilightBits);

					SelectObject(HilightDC, HilightPixel);
					
					BLENDFUNCTION InverseBlendFunction = { AC_SRC_OVER, 0, 0, AC_SRC_ALPHA };					

					for (UINT8 XPixel = 0; XPixel < 10; XPixel++)
					{
						for (UINT8 YPixel = 0; YPixel < 20; YPixel++)
						{
							BOOL PixelAlreadyDrawn = FALSE;							

							for (UINT16 Counter = 0; Counter < ArrayCount(HilighterPixelsAlreadyDrawn); Counter++)
							{
								if ((HilighterPixelsAlreadyDrawn[Counter].x == Mouse.x + XPixel) && (HilighterPixelsAlreadyDrawn[Counter].y == Mouse.y + YPixel))
								{
									PixelAlreadyDrawn = TRUE;
								}
							}

							if (PixelAlreadyDrawn == FALSE)
							{								
								COLORREF CurrentPixelColor = GetPixel(ScratchDC, Mouse.x + XPixel, Mouse.y + YPixel);
								UINT16 CurrentPixelBrightness = GetRValue(CurrentPixelColor) + GetGValue(CurrentPixelColor) + GetBValue(CurrentPixelColor);

								// x/3, 765 = 255, 382.5 = 128, 0 = 0
								InverseBlendFunction.SourceConstantAlpha = (BYTE)(CurrentPixelBrightness / 3);

								GdiAlphaBlend(ScratchDC, Mouse.x + XPixel, Mouse.y + YPixel, 1, 1, HilightDC, 0, 0, 1, 1, InverseBlendFunction);

								HilighterPixelsAlreadyDrawn[HilighterPixelsAlreadyDrawnCounter].x = (UINT16)(Mouse.x + XPixel);
								HilighterPixelsAlreadyDrawn[HilighterPixelsAlreadyDrawnCounter].y = (UINT16)(Mouse.y + YPixel);
								HilighterPixelsAlreadyDrawnCounter++;
								if (HilighterPixelsAlreadyDrawnCounter >= ArrayCount(HilighterPixelsAlreadyDrawn) - 1)
								{
									HilighterPixelsAlreadyDrawnCounter = 0;
								}
							}
						}
					}
				
					if (DeleteDC(HilightDC) == 0)
					{
						MyOutputDebugStringA("[%s] Line %d: DeleteDC(HilightDC) failed!\n", __FUNCTION__, __LINE__);
					}
					if (DeleteDC(ScratchDC) == 0)
					{
						MyOutputDebugStringA("[%s] Line %d: DeleteDC(gScratchDC) failed!\n", __FUNCTION__, __LINE__);
					}
					if (DeleteObject(HilightPixel) == 0)
					{
						MyOutputDebugStringA("[%s] Line %d: DeleteObject(HilightPixel) failed!\n", __FUNCTION__, __LINE__);
					}

					PreviousMousePos.x = Mouse.x;

					RECT SnipRect = { 0 };
					GetClientRect(Window, &SnipRect);

					SnipRect.top = MousePosWhenDrawingStarted.y - 10;
					SnipRect.bottom = MousePosWhenDrawingStarted.y + 10;

					InvalidateRect(Window, &SnipRect, FALSE);
					UpdateWindow(gMainWindowHandle);
				}
				else if (gRectangleButton.SelectedTool == TRUE)
				{
					HBITMAP CleanCopy = CopyImage(gSnipStates[gCurrentSnipState - 1], IMAGE_BITMAP, 0, 0, 0);

					HDC CopyDC = CreateCompatibleDC(NULL);
					
					SelectObject(CopyDC, CleanCopy);

					HPEN RedPen = CreatePen(PS_SOLID, 2, RGB(255, 0, 0));

					SelectObject(CopyDC, RedPen);

					SelectObject(CopyDC, (HBRUSH)GetStockObject(NULL_BRUSH));

					POINT CurrentMousePos = { 0 };
					GetCursorPos(&CurrentMousePos);
					ScreenToClient(gMainWindowHandle, &CurrentMousePos);

					Rectangle(CopyDC, MousePosWhenDrawingStarted.x, MousePosWhenDrawingStarted.y - 56, CurrentMousePos.x, CurrentMousePos.y - 56);

					HDC ScratchDC = CreateCompatibleDC(NULL);

					SelectObject(ScratchDC, (HBITMAP)gScratchBitmap);
					
					BitBlt(ScratchDC, 0, 0, gCaptureWidth, gCaptureHeight, CopyDC, 0, 0, SRCCOPY);					
					
					// DeleteObject will fail if the object is still selected into a DC.
					if (DeleteDC(CopyDC) == 0)
					{
						MyOutputDebugStringA("[%s] Line %d: DeleteDC(CopyDC) failed!\n", __FUNCTION__, __LINE__);
					}
					if (DeleteDC(ScratchDC) == 0)
					{
						MyOutputDebugStringA("[%s] Line %d: DeleteDC(gScratchDC) failed!\n", __FUNCTION__, __LINE__);
					}
					if (DeleteObject(RedPen) == 0)
					{
						MyOutputDebugStringA("[%s] Line %d: DeleteObject(RedPen) failed!\n", __FUNCTION__, __LINE__);
					}
					if (DeleteObject(CleanCopy) == 0)
					{
						MyOutputDebugStringA("[%s] Line %d: DeleteObject(CleanCopy) failed!\n", __FUNCTION__, __LINE__);
					}

					RECT SnipRect = { 0 };
					GetClientRect(Window, &SnipRect);

					SnipRect.top = 54;					

					InvalidateRect(Window, &SnipRect, FALSE);
					UpdateWindow(gMainWindowHandle);
				}
				else if (gArrowButton.SelectedTool == TRUE)
				{
					HBITMAP CleanCopy = CopyImage(gSnipStates[gCurrentSnipState - 1], IMAGE_BITMAP, 0, 0, 0);

					HDC CopyDC = CreateCompatibleDC(NULL);

					SelectObject(CopyDC, CleanCopy);

					HPEN RedPen     = CreatePen(PS_SOLID, 2, RGB(255, 0, 0));
					HBRUSH RedBrush = CreateSolidBrush(RGB(255, 0, 0));

					SelectObject(CopyDC, RedPen);
					SelectObject(CopyDC, RedBrush);

					POINT CurrentMousePos = { 0 };
					
					GetCursorPos(&CurrentMousePos);
					ScreenToClient(gMainWindowHandle, &CurrentMousePos);

					POINT p0 = { 0 };
					POINT p1 = { 0 };

					p0.x = MousePosWhenDrawingStarted.x;
					p0.y = MousePosWhenDrawingStarted.y - 56;

					p1.x = CurrentMousePos.x;
					p1.y = CurrentMousePos.y - 56;

					MoveToEx(CopyDC, p0.x, p0.y, NULL);
					LineTo(CopyDC, p1.x, p1.y);					

					const float dx = (float)(p1.x - p0.x);
					const float dy = (float)(p1.y - p0.y);

					float ArrowLength = (float)sqrt(dx*dx + dy*dy);

					// unit vector parallel to the line.
					float ux = dx / ArrowLength;
					float uy = dy / ArrowLength;

					// unit vector perpendicular to ux,uy
					float vx = -uy;
					float vy = ux;

					int head_length = 10;
					int head_width  = 10;

					float half_width = 0.5f * head_width;
					
					POINT ArrowHead = p1;
					POINT ArrowCorner1 = { 0 };
					
					ArrowCorner1.x = (LONG)roundf(p1.x - head_length*ux + half_width*vx);
					ArrowCorner1.y = (LONG)roundf(p1.y - head_length*uy + half_width*vy);

					POINT ArrowCorner2 = { 0 };
					
					ArrowCorner2.x = (LONG)roundf(p1.x - head_length*ux - half_width*vx);
					ArrowCorner2.y = (LONG)roundf(p1.y - head_length*uy - half_width*vy);

					POINT Arrow[3];
					
					Arrow[0] = ArrowHead;
					Arrow[1] = ArrowCorner1;
					Arrow[2] = ArrowCorner2;
					
					Polygon(CopyDC, Arrow, 3);					

					HDC ScratchDC = CreateCompatibleDC(NULL);

					SelectObject(ScratchDC, (HBITMAP)gScratchBitmap);

					BitBlt(ScratchDC, 0, 0, gCaptureWidth, gCaptureHeight, CopyDC, 0, 0, SRCCOPY);

					// DeleteObject will fail if the object is still selected into a DC.
					if (DeleteDC(CopyDC) == 0)
					{
						MyOutputDebugStringA("[%s] Line %d: DeleteDC(CopyDC) failed!\n", __FUNCTION__, __LINE__);
					}
					if (DeleteDC(ScratchDC) == 0)
					{
						MyOutputDebugStringA("[%s] Line %d: DeleteDC(gScratchDC) failed!\n", __FUNCTION__, __LINE__);
					}
					if (DeleteObject(RedPen) == 0)
					{
						MyOutputDebugStringA("[%s] Line %d: DeleteObject(RedPen) failed!\n", __FUNCTION__, __LINE__);
					}
					if (DeleteObject(RedBrush) == 0)
					{
						MyOutputDebugStringA("[%s] Line %d: DeleteObject(RedBrush) failed!\n", __FUNCTION__, __LINE__);
					}
					if (DeleteObject(CleanCopy) == 0)
					{
						MyOutputDebugStringA("[%s] Line %d: DeleteObject(CleanCopy) failed!\n", __FUNCTION__, __LINE__);
					}

					RECT SnipRect = { 0 };
					GetClientRect(Window, &SnipRect);

					SnipRect.top = 54;

					InvalidateRect(Window, &SnipRect, FALSE);
					UpdateWindow(gMainWindowHandle);
				}
				else if (gRedactButton.SelectedTool == TRUE)
				{
					POINT Mouse = { 0 };

					GetCursorPos(&Mouse);

					ScreenToClient(gMainWindowHandle, &Mouse);

					// Adjust for snip area, maintain Y axis
					Mouse.x -= 7;
					Mouse.y = MousePosWhenDrawingStarted.y - 66;

					HDC ScratchDC = CreateCompatibleDC(NULL);					

					SelectObject(ScratchDC, (HBITMAP)gScratchBitmap);					

					for (UINT8 XPixel = 0; XPixel < 10; XPixel++)
					{
						for (UINT8 YPixel = 0; YPixel < 20; YPixel++)
						{
							PatBlt(ScratchDC, Mouse.x + XPixel, Mouse.y + YPixel, 1, 1, BLACKNESS);								
						}
					}

					if (DeleteDC(ScratchDC) == 0)
					{
						MyOutputDebugStringA("[%s] Line %d: DeleteDC(gScratchDC) failed!\n", __FUNCTION__, __LINE__);
					}

					PreviousMousePos.x = Mouse.x;

					RECT SnipRect = { 0 };
					GetClientRect(Window, &SnipRect);

					SnipRect.top = MousePosWhenDrawingStarted.y - 10;
					SnipRect.bottom = MousePosWhenDrawingStarted.y + 10;

					InvalidateRect(Window, &SnipRect, FALSE);
					UpdateWindow(gMainWindowHandle);
				}
			}

			// We only receive this message when the mouse is not over a button. So if we get this message, no button should be pressed. 
			// Unless it's a selected tool which should stay pressed.

			for (UINT8 Counter = 0; Counter < ArrayCount(gButtons); Counter++)
			{
				if (gButtons[Counter]->State != BUTTONSTATE_NORMAL && gButtons[Counter]->SelectedTool == FALSE)
				{
					gButtons[Counter]->State = BUTTONSTATE_NORMAL;
					InvalidateRect(Window, NULL, FALSE);
				}

				if (gButtons[Counter]->SelectedTool == TRUE && gAppState == APPSTATE_AFTERCAPTURE && gSnipStates[0] != NULL && gButtons[Counter]->Cursor != NULL)
				{
					POINT Mouse = { 0 };

					GetCursorPos(&Mouse);

					ScreenToClient(gMainWindowHandle, &Mouse);

					if (Mouse.x >= 2 && Mouse.y >= 54)
					{
						if (GetCursor() != gButtons[Counter]->Cursor)
						{
							SetCursor(gButtons[Counter]->Cursor);
						}
					}
					else
					{
						if (GetCursor() != LoadCursor(NULL, IDC_ARROW))
						{
							SetCursor(LoadCursor(NULL, IDC_ARROW));
						}
					}
				}
			}

			break;
		}
		case WM_PARENTNOTIFY:
		{
			switch (LOWORD(WParam))
			{
				case WM_LBUTTONDOWN:
				{
					POINT Mouse = { 0 };
					
					Mouse.x = GET_X_LPARAM(LParam);
					Mouse.y = GET_Y_LPARAM(LParam);					

					HWND Control = ChildWindowFromPoint(Window, Mouse);

					for (UINT8 Counter = 0; Counter < ArrayCount(gButtons); Counter++)
					{
						if (Control == gButtons[Counter]->Handle)
						{
							MyOutputDebugStringA("[%s] Line %d: Mouse button pressed over '%s' button.\n", __FUNCTION__, __LINE__, gButtons[Counter]->Caption);

							if (gButtons[Counter]->Enabled == TRUE)
							{
								gButtons[Counter]->State = BUTTONSTATE_PRESSED;
							}
							else
							{
								MyOutputDebugStringA("[%s] Line %d: ...but it was disabled.\n", __FUNCTION__, __LINE__);
							}
						}
						else
						{
							if (gButtons[Counter]->SelectedTool == FALSE)
							{
								gButtons[Counter]->State = BUTTONSTATE_NORMAL;
							}
						}
					}

					// Redraw the toolbar.
					RECT ToolbarRect = { 0 };
					GetClientRect(gMainWindowHandle, &ToolbarRect);
					ToolbarRect.bottom = gButtons[0]->Rectangle.bottom;					
					InvalidateRect(gMainWindowHandle, &ToolbarRect, FALSE);

					break;
				}
			}
			break;
		}
		case WM_DRAWITEM:
		{
			for (UINT8 Counter = 0; Counter < ArrayCount(gButtons); Counter++)
			{
				if (LOWORD(WParam) == gButtons[Counter]->Id)
				{
					DrawButton((DRAWITEMSTRUCT*)LParam, *gButtons[Counter]);
					return(TRUE);
				}
			}

			break;
		}
		case WM_COMMAND:
		{
			// User clicked a button - return all buttons to non-pressed state.
			for (UINT8 Counter = 0; Counter < ArrayCount(gButtons); Counter++)
			{
				if (gButtons[Counter]->SelectedTool == FALSE)
				{
					gButtons[Counter]->State = BUTTONSTATE_NORMAL;
				}
			}

			// Redraw the toolbar.
			RECT ToolbarRect = { 0 };
			GetClientRect(gMainWindowHandle, &ToolbarRect);
			ToolbarRect.bottom = gButtons[0]->Rectangle.bottom;
			InvalidateRect(gMainWindowHandle, &ToolbarRect, FALSE);

			// We just left-clicked a button, but we need to set focus back on the main window or else
			// the main window will stop receiving window messages such as WM_KEYDOWN.
			SetFocus(gMainWindowHandle);

			switch (LOWORD(WParam))
			{
				case BUTTON_NEW:
				{
					MyOutputDebugStringA("[%s] Line %d: Mouse button released over 'New' button.\n", __FUNCTION__, __LINE__);

					if (NewButton_Click() == FALSE)
					{
						// Fatal!
						gMainWindowIsRunning = FALSE;
					}				

					break;
				}
				case BUTTON_DELAY:
				{
					MyOutputDebugStringA("[%s] Line %d: Mouse button released over 'Delay' button.\n", __FUNCTION__, __LINE__);

					char TitleBuffer[32] = { 0 };
					(void)snprintf(TitleBuffer, sizeof(TitleBuffer), "SnipEx");
					SetWindowText(gMainWindowHandle, TitleBuffer);

					gAppState = APPSTATE_DELAYCOOKING;

					RECT CurrentWindowPos = { 0 };
					GetWindowRect(gMainWindowHandle, &CurrentWindowPos);

					SetWindowPos(
						gMainWindowHandle,
						HWND_TOP,
						CurrentWindowPos.left,
						CurrentWindowPos.top,
						gStartingMainWindowWidth,
						gStartingMainWindowHeight,
						0);

					KillTimer(gMainWindowHandle, DELAY_TIMER);

					SetTimer(gMainWindowHandle, DELAY_TIMER, 1000, NULL);

					for (UINT8 Counter = 0; Counter < ArrayCount(gButtons); Counter++)
					{
						if (gButtons[Counter]->Id != BUTTON_NEW && gButtons[Counter]->Id != BUTTON_DELAY)
						{
							gButtons[Counter]->Enabled = FALSE;
							gButtons[Counter]->SelectedTool = FALSE;
							gButtons[Counter]->State = BUTTONSTATE_NORMAL;
						}
					}					
					
					break;
				}
				case BUTTON_SAVE:
				{
					if (gSaveButton.Enabled == TRUE)
					{
						MyOutputDebugStringA("[%s] Line %d: Mouse button released over 'Save' button.\n", __FUNCTION__, __LINE__);

						if (SaveButton_Click() == FALSE)
						{
							MyOutputDebugStringA("[%s] Line %d: SaveButton_Click either failed or was cancelled by the user.\n", __FUNCTION__, __LINE__);
						}
					}
					break;
				}
				case BUTTON_COPY:
				{
					if (gCopyButton.Enabled == TRUE)
					{
						MyOutputDebugStringA("[%s] Line %d: Mouse button released over 'Copy' button.\n", __FUNCTION__, __LINE__);

						if (CopyButton_Click() == FALSE)
						{
							MyOutputDebugStringA("[%s] Line %d: CopyButton_Click failed!\n", __FUNCTION__, __LINE__);
						}
					}
					break;
				}
				case BUTTON_HILIGHT:
				{
					if (gHilighterButton.Enabled == TRUE)
					{
						MyOutputDebugStringA("[%s] Line %d: Mouse button released over 'Hilight' button.\n", __FUNCTION__, __LINE__);

						for (UINT8 Counter = 0; Counter < ArrayCount(gButtons); Counter++)
						{
							if (gButtons[Counter]->Id == BUTTON_HILIGHT)
							{
								gButtons[Counter]->SelectedTool = TRUE;
								gButtons[Counter]->State = BUTTONSTATE_PRESSED;
							}
							else
							{
								if (gButtons[Counter]->SelectedTool == TRUE)
								{
									gButtons[Counter]->SelectedTool = FALSE;
									gButtons[Counter]->State = BUTTONSTATE_NORMAL;
								}
							}
						}
					}
					break;
				}
				case BUTTON_BOX:
				{
					if (gRectangleButton.Enabled == TRUE)
					{
						MyOutputDebugStringA("[%s] Line %d: Mouse button released over 'Box' button.\n", __FUNCTION__, __LINE__);

						for (UINT8 Counter = 0; Counter < ArrayCount(gButtons); Counter++)
						{
							if (gButtons[Counter]->Id == BUTTON_BOX)
							{
								gButtons[Counter]->SelectedTool = TRUE;
								gButtons[Counter]->State = BUTTONSTATE_PRESSED;
							}
							else
							{
								if (gButtons[Counter]->SelectedTool == TRUE)
								{
									gButtons[Counter]->SelectedTool = FALSE;
									gButtons[Counter]->State = BUTTONSTATE_NORMAL;
								}
							}
						}
					}
					break;
				}
				case BUTTON_ARROW:
				{
					if (gArrowButton.Enabled == TRUE)
					{
						MyOutputDebugStringA("[%s] Line %d: Mouse button released over 'Arrow' button.\n", __FUNCTION__, __LINE__);

						for (UINT8 Counter = 0; Counter < ArrayCount(gButtons); Counter++)
						{
							if (gButtons[Counter]->Id == BUTTON_ARROW)
							{
								gButtons[Counter]->SelectedTool = TRUE;
								gButtons[Counter]->State = BUTTONSTATE_PRESSED;
							}
							else
							{
								if (gButtons[Counter]->SelectedTool == TRUE)
								{
									gButtons[Counter]->SelectedTool = FALSE;
									gButtons[Counter]->State = BUTTONSTATE_NORMAL;
								}
							}
						}
					}
					break;
				}
				case BUTTON_REDACT:
				{
					if (gRedactButton.Enabled == TRUE)
					{
						MyOutputDebugStringA("[%s] Line %d: Mouse button released over 'Redact' button.\n", __FUNCTION__, __LINE__);

						for (UINT8 Counter = 0; Counter < ArrayCount(gButtons); Counter++)
						{
							if (gButtons[Counter]->Id == BUTTON_REDACT)
							{
								gButtons[Counter]->SelectedTool = TRUE;
								gButtons[Counter]->State = BUTTONSTATE_PRESSED;
							}
							else
							{
								if (gButtons[Counter]->SelectedTool == TRUE)
								{
									gButtons[Counter]->SelectedTool = FALSE;
									gButtons[Counter]->State = BUTTONSTATE_NORMAL;
								}
							}
						}
					}
					break;
				}
			}
			break;
		}
		case WM_SYSCOMMAND:
		{
			// Default system messages are >= 0xf000
			if (WParam >= 0xF000) 
			{
				Result = DefWindowProc(Window, Message, WParam, LParam);
			}
			else if (WParam == SYSCMD_REPLACE)
			{
				MyOutputDebugStringA("[%s] Line %d: User clicked on 'Replace Snipping Tool'.\n", __FUNCTION__, __LINE__);
				// Need admin and UAC elevation to write to the image file execution options registry key
				if (IsAppRunningElevated())
				{
					MyOutputDebugStringA("[%s] Line %d: Already UAC elevated.\n", __FUNCTION__, __LINE__);

					HKEY IFEOKey = NULL;
					HKEY SnippingToolKey = NULL;

					if (RegOpenKeyEx(HKEY_LOCAL_MACHINE, "SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Image File Execution Options", 0, KEY_ALL_ACCESS, &IFEOKey) != ERROR_SUCCESS)
					{
						MyMessageBoxA(gMainWindowHandle, "Error", MB_OK | MB_ICONERROR | MB_SYSTEMMODAL, "Error while attempting to open the Image File Execution Options subkey! Error: 0x%lx", GetLastError());
					}
					else
					{
						
						DWORD SnippingToolKeyDisposition = 0;
						DWORD Error = RegCreateKeyEx(IFEOKey, "SnippingTool.exe", 0, NULL, 0, KEY_WRITE, NULL, &SnippingToolKey, &SnippingToolKeyDisposition);
						if (Error == ERROR_SUCCESS)
						{
							char ModulePath[MAX_PATH] = { 0 };

							GetModuleFileName(NULL, ModulePath, sizeof(ModulePath));
							Error = RegSetValueEx(SnippingToolKey, "Debugger", 0, REG_SZ, (const BYTE*)ModulePath, (DWORD)strlen(ModulePath));
							if (Error == ERROR_SUCCESS)
							{
								HMENU SystemMenu = GetSystemMenu(gMainWindowHandle, FALSE);

								RemoveMenu(SystemMenu, SYSCMD_REPLACE, MF_BYCOMMAND);
								AppendMenu(SystemMenu, MF_STRING, SYSCMD_RESTORE, "Restore Windows Snipping Tool\n");	

								if (gUACIcon != NULL)
								{
									MENUITEMINFO MenuItemInfo = { sizeof(MENUITEMINFO) };

									MenuItemInfo.fMask = MIIM_BITMAP;
									MenuItemInfo.hbmpItem = gUACIcon;

									SetMenuItemInfo(SystemMenu, SYSCMD_RESTORE, FALSE, &MenuItemInfo);
								}

								MyMessageBoxA(gMainWindowHandle, "Success", MB_OK | MB_ICONINFORMATION | MB_SYSTEMMODAL, "The built-in SnippingTool.exe has been replaced.");
							}
							else
							{
								MyMessageBoxA(gMainWindowHandle, "Error", MB_OK | MB_ICONERROR | MB_SYSTEMMODAL, "Error while creating Debugger registry value! Error: 0x%lx", Error);
							}
						}
						else
						{
							MyMessageBoxA(gMainWindowHandle, "Error", MB_OK | MB_ICONERROR | MB_SYSTEMMODAL, "Error while creating SnippingTool.exe registry key! Error: 0x%lx", Error);
						}
					}

					if (SnippingToolKey != NULL)
					{
						RegCloseKey(SnippingToolKey);
					}
					if (IFEOKey != NULL)
					{
						RegCloseKey(IFEOKey);
					}
				}
				else
				{
					MyOutputDebugStringA("[%s] Line %d: Need to UAC elevate first.\n", __FUNCTION__, __LINE__);

					MyMessageBoxA(gMainWindowHandle, "UAC Elevation Required", MB_OK | MB_ICONINFORMATION, "This action requires UAC elevation. After clicking OK, you will be prompted to elevate. Then, retry the operation.");

					char ModulePath[MAX_PATH] = { 0 };

					GetModuleFileName(NULL, ModulePath, sizeof(ModulePath));

					SHELLEXECUTEINFO ShellExecuteInfo = { sizeof(SHELLEXECUTEINFO) };
					ShellExecuteInfo.lpVerb = "runas";
					ShellExecuteInfo.lpFile = ModulePath;
					ShellExecuteInfo.hwnd = NULL;
					ShellExecuteInfo.nShow = SW_NORMAL;

					if (!ShellExecuteEx(&ShellExecuteInfo))
					{
						DWORD Error = GetLastError();
						if (Error != ERROR_CANCELLED)
						{
							MyMessageBoxA(gMainWindowHandle, "Error", MB_OK | MB_ICONERROR | MB_SYSTEMMODAL, "Error when attempting to re-launch the application with UAC elevation. Error: 0x%lx", Error);
						}
					}
					else
					{
						// We just re-launched the app with UAC elevation - need to exit this instance.
						PostQuitMessage(0);
						gMainWindowIsRunning = FALSE;
					}
				}
			}
			else if (WParam == SYSCMD_RESTORE)
			{
				MyOutputDebugStringA("[%s] Line %d: User clicked on 'Restore Snipping Tool'.\n", __FUNCTION__, __LINE__);
				// To "restore" the built-in snipping tool, all we have to do is delete the snippingtool.exe Image File Execution Options registry key.
				if (IsAppRunningElevated())
				{
					MyOutputDebugStringA("[%s] Line %d: Already UAC elevated.\n", __FUNCTION__, __LINE__);
					
					HKEY IFEOKey = NULL;
					if (RegOpenKeyEx(HKEY_LOCAL_MACHINE, "SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Image File Execution Options", 0, KEY_QUERY_VALUE, &IFEOKey) != ERROR_SUCCESS)
					{
						MyMessageBoxA(gMainWindowHandle, "Error", MB_OK | MB_ICONERROR | MB_SYSTEMMODAL, "Error while attempting to open the Image File Execution Options subkey! Error: 0x%lx", GetLastError());						
					}
					else
					{
						// This will fail if the subkey has subkeys
						DWORD Error = ERROR_SUCCESS;
						Error = RegDeleteKey(IFEOKey, "SnippingTool.exe");
						if (Error != ERROR_SUCCESS)
						{
							MyMessageBoxA(gMainWindowHandle, "Error", MB_OK | MB_ICONERROR | MB_SYSTEMMODAL, "Error while attempting to delete the SnippingTool.exe subkey! Error: 0x%lx", Error);
						}
						else
						{
							HMENU SystemMenu = GetSystemMenu(gMainWindowHandle, FALSE);

							RemoveMenu(SystemMenu, SYSCMD_RESTORE, MF_BYCOMMAND);
							AppendMenu(SystemMenu, MF_STRING, SYSCMD_REPLACE, "Replace Windows Snipping Tool with SnipEx\n");	

							if (gUACIcon != NULL)
							{
								MENUITEMINFO MenuItemInfo = { sizeof(MENUITEMINFO) };

								MenuItemInfo.fMask = MIIM_BITMAP;
								MenuItemInfo.hbmpItem = gUACIcon;

								SetMenuItemInfo(SystemMenu, SYSCMD_REPLACE, FALSE, &MenuItemInfo);
							}

							MyMessageBoxA(gMainWindowHandle, "Success", MB_OK | MB_ICONINFORMATION | MB_SYSTEMMODAL, "The built-in SnippingTool.exe has been restored.");
						}
					}

					if (IFEOKey != NULL)
					{
						RegCloseKey(IFEOKey);
					}
				}
				else
				{
					MyOutputDebugStringA("[%s] Line %d: Need to UAC elevate first.\n", __FUNCTION__, __LINE__);

					MyMessageBoxA(gMainWindowHandle, "UAC Elevation Required", MB_OK | MB_ICONINFORMATION, "This action requires UAC elevation. After clicking OK, you will be prompted to elevate. Then, retry the operation.");

					char ModulePath[MAX_PATH] = { 0 };

					GetModuleFileName(NULL, ModulePath, sizeof(ModulePath));

					SHELLEXECUTEINFO ShellExecuteInfo = { sizeof(SHELLEXECUTEINFO) };
					ShellExecuteInfo.lpVerb = "runas";
					ShellExecuteInfo.lpFile = ModulePath;
					ShellExecuteInfo.hwnd   = NULL;
					ShellExecuteInfo.nShow  = SW_NORMAL;

					if (!ShellExecuteEx(&ShellExecuteInfo))
					{
						DWORD Error = GetLastError();
						if (Error != ERROR_CANCELLED)
						{
							MyMessageBoxA(gMainWindowHandle, "Error", MB_OK | MB_ICONERROR | MB_SYSTEMMODAL, "Error when attempting to re-launch the application with UAC elevation. Error: 0x%lx", Error);						
						}							
					}
					else
					{
						// We just re-launched the app with UAC elevation - need to exit this instance.
						PostQuitMessage(0);
						gMainWindowIsRunning = FALSE;
					}
				}
			}
			else if (WParam == SYSCMD_SHADOW)
			{
				if (gShouldAddDropShadow)
				{
					CheckMenuItem(GetSystemMenu(gMainWindowHandle, FALSE), SYSCMD_SHADOW, MF_BYCOMMAND | MF_UNCHECKED);
					gShouldAddDropShadow = FALSE;
				}
				else
				{
					CheckMenuItem(GetSystemMenu(gMainWindowHandle, FALSE), SYSCMD_SHADOW, MF_BYCOMMAND | MF_CHECKED);
					gShouldAddDropShadow = TRUE;
				}

				HKEY SoftwareKey = NULL;
				HKEY SnipExKey   = NULL;

				DWORD SnipExKeyDisposition = 0;			

				// This key should always exist. Something is very wrong if we can't open it.
				if (RegOpenKeyEx(HKEY_CURRENT_USER, "SOFTWARE", 0, KEY_ALL_ACCESS, &SoftwareKey) != ERROR_SUCCESS)
				{
					MyMessageBoxA(NULL, "Error", MB_OK | MB_ICONERROR | MB_SYSTEMMODAL, "Unable to read HKCU\\SOFTWARE registry key! Error: 0x%lx", GetLastError());
					CheckMenuItem(GetSystemMenu(gMainWindowHandle, FALSE), SYSCMD_SHADOW, MF_BYCOMMAND | MF_UNCHECKED);
					gShouldAddDropShadow = FALSE;
				}
				else
				{
					// This subkey must exist because we created it during application startup in AddDropShadowToolMenuItem
					if (RegCreateKeyEx(SoftwareKey, "SnipEx", 0, NULL, 0, KEY_ALL_ACCESS, NULL, &SnipExKey, &SnipExKeyDisposition) == ERROR_SUCCESS)
					{
						if (RegSetValueEx(SnipExKey, "DropShadow", 0, REG_DWORD, (const BYTE*)&gShouldAddDropShadow, sizeof(DWORD)) != ERROR_SUCCESS)
						{
							MyMessageBoxA(gMainWindowHandle, "Error", MB_OK | MB_SYSTEMMODAL | MB_ICONERROR, "Could not set the DropShadow registry key value!");
							CheckMenuItem(GetSystemMenu(gMainWindowHandle, FALSE), SYSCMD_SHADOW, MF_BYCOMMAND | MF_UNCHECKED);
							gShouldAddDropShadow = FALSE;
						}
					}
					else
					{
						MyMessageBoxA(gMainWindowHandle, "Error", MB_OK | MB_SYSTEMMODAL | MB_ICONERROR, "Could not open the SnipEx registry key for writing!");
						CheckMenuItem(GetSystemMenu(gMainWindowHandle, FALSE), SYSCMD_SHADOW, MF_BYCOMMAND | MF_UNCHECKED);
						gShouldAddDropShadow = FALSE;
					}
				}

				if (SnipExKey != NULL)
				{
					RegCloseKey(SnipExKey);
				}

				if (SoftwareKey != NULL)
				{
					RegCloseKey(SoftwareKey);
				}
			}

			break;
		}
		case WM_TIMER:
		{
			if (WParam == DELAY_TIMER)
			{
				gCurrentDelayCountdown--;

				MyOutputDebugStringA("[%s] Line %d: WM_TIMER DELAY_TIMER received. %d seconds.\n", __FUNCTION__, __LINE__, gCurrentDelayCountdown);

				InvalidateRect(gMainWindowHandle, NULL, FALSE);				

				if (gCurrentDelayCountdown <= 0)
				{
					SendMessage(gMainWindowHandle, WM_COMMAND, BUTTON_NEW, 0);
				}
				else
				{
					SetTimer(gMainWindowHandle, DELAY_TIMER, 1000, NULL);
				}
			}
			break;
		}
		case WM_PAINT:
		{
			PAINTSTRUCT PaintStruct = { 0 };
			
			BeginPaint(Window, &PaintStruct);
			
			if (gAppState == APPSTATE_AFTERCAPTURE)
			{
				HDC MemDC = CreateCompatibleDC(PaintStruct.hdc);

				if (CurrentlyDrawing == TRUE)
				{
					if (gScratchBitmap != NULL)
					{						
						SelectObject(MemDC, gScratchBitmap);						
					}
				}
				else
				{
					if (gSnipStates[gCurrentSnipState] != NULL)
					{						
						SelectObject(MemDC, gSnipStates[gCurrentSnipState]);						
					}
				}

				BitBlt(PaintStruct.hdc, 2, 56, gCaptureWidth, gCaptureHeight, MemDC, 0, 0, SRCCOPY);
				DeleteDC(MemDC);
			}	

			EndPaint(Window, &PaintStruct);

			break;
		}
		default:
		{
			Result = DefWindowProc(Window, Message, WParam, LParam);
		}
	}

	return(Result);
}

void DrawButton(_In_ DRAWITEMSTRUCT* DrawItemStruct, _In_ BUTTON Button)
{
	HPEN DarkPen = CreatePen(PS_SOLID, 1, RGB(32, 32, 32));
	HPEN LighterPen = CreatePen(PS_SOLID, 1, RGB(224, 224, 224));
	SIZE StringSizeInPixels = { 0 };
	UINT16 ButtonWidth = (UINT16)(Button.Rectangle.right - Button.Rectangle.left);
	HFONT ButtonFont = NULL;

	ButtonFont = CreateFont(-12, 0, 0, 0, FW_NORMAL, 0, 0, 0, ANSI_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, DEFAULT_PITCH, "Segoe UI");

	SelectObject(DrawItemStruct->hDC, ButtonFont);
	SetDCBrushColor(DrawItemStruct->hDC, GetSysColor(COLOR_BTNFACE));
	SelectObject(DrawItemStruct->hDC, GetStockObject(DC_BRUSH));

	GetTextExtentPoint32(DrawItemStruct->hDC, Button.Caption, (int)strlen(Button.Caption), &StringSizeInPixels);

	UINT16 TextX = (UINT16)((ButtonWidth / 2) - (StringSizeInPixels.cx / 2));
	UINT16 TextY = 34;

	if (Button.State == BUTTONSTATE_PRESSED)
	{
		TextY += 1;
		SelectObject(DrawItemStruct->hDC, LighterPen);
	}
	else
	{
		SelectObject(DrawItemStruct->hDC, DarkPen);
	}	
	
	MoveToEx(DrawItemStruct->hDC, 0, Button.Rectangle.bottom - 1, NULL);
	LineTo(DrawItemStruct->hDC, ButtonWidth - 1, Button.Rectangle.bottom - 1);
	LineTo(DrawItemStruct->hDC, ButtonWidth - 1, 0);

	if (Button.State == BUTTONSTATE_PRESSED)
	{
		SelectObject(DrawItemStruct->hDC, DarkPen);
	}
	else
	{
		SelectObject(DrawItemStruct->hDC, LighterPen);
	}

	MoveToEx(DrawItemStruct->hDC, ButtonWidth - 2, Button.Rectangle.top, NULL);
	LineTo(DrawItemStruct->hDC, 0, Button.Rectangle.top);
	LineTo(DrawItemStruct->hDC, 0, Button.Rectangle.bottom);

	if (Button.Enabled == TRUE)
	{
		SetTextColor(DrawItemStruct->hDC, RGB(0, 0, 0));
		SelectObject(DrawItemStruct->hDC, DarkPen);
	}
	else
	{
		SetTextColor(DrawItemStruct->hDC, RGB(160, 160, 160));
		SelectObject(DrawItemStruct->hDC, LighterPen);
	}

	TextOut(DrawItemStruct->hDC, TextX, TextY, Button.Caption, (int)strlen(Button.Caption));

	MoveToEx(DrawItemStruct->hDC, TextX + 6, TextY + 13, NULL);
	LineTo(DrawItemStruct->hDC, TextX , TextY + 13);

	if (DeleteObject(ButtonFont) == 0)
	{
		MyOutputDebugStringA("[%s] Line %d: DeleteObject(ButtonFont) failed!\n", __FUNCTION__, __LINE__);
	}

	if (DeleteObject(LighterPen) == 0)
	{
		MyOutputDebugStringA("[%s] Line %d: DeleteObject(LighterPen) failed!\n", __FUNCTION__, __LINE__);
	}

	if (DeleteObject(DarkPen) == 0)
	{
		MyOutputDebugStringA("[%s] Line %d: DeleteObject(DarkPen) failed!\n", __FUNCTION__, __LINE__);
	}

	if (gAppState == APPSTATE_DELAYCOOKING && Button.Id == BUTTON_DELAY)
	{
		char CountdownBuffer[4] = { 0 };
		_itoa_s(gCurrentDelayCountdown, CountdownBuffer, sizeof(CountdownBuffer), 10);
		TextOut(DrawItemStruct->hDC, TextX + StringSizeInPixels.cx + 4, TextY, CountdownBuffer, (int)strlen(CountdownBuffer));
	}

	if (Button.Enabled == TRUE)
	{
		if (Button.EnabledIcon != NULL)
		{
			HDC IconDC = CreateCompatibleDC(DrawItemStruct->hDC);
			SelectObject(IconDC, Button.EnabledIcon);						
			TransparentBlt(DrawItemStruct->hDC, 19, TextY - 32, 32, 32, IconDC, 0, 0, 32, 32, RGB(255, 255, 255));
			DeleteDC(IconDC);
		}
	}
	else
	{
		if (Button.DisabledIcon != NULL)
		{
			HDC IconDC = CreateCompatibleDC(DrawItemStruct->hDC);
			SelectObject(IconDC, Button.DisabledIcon);
			TransparentBlt(DrawItemStruct->hDC, 19, TextY - 32, 32, 32, IconDC, 0, 0, 32, 32, RGB(255, 255, 255));
			DeleteDC(IconDC);
		}
	}
}

LRESULT CALLBACK CaptureWindowCallback(_In_ HWND Window, _In_ UINT Message, _In_ WPARAM WParam, _In_ LPARAM LParam)
{
	LRESULT Result = 0;
	
	static BOOL LMouseButtonDown = FALSE;
	static BOOL MouseHasMovedWhileLeftMouseButtonWasDown = FALSE;

	switch (Message)
	{
		case WM_KEYDOWN:
		{
			if (WParam == VK_ESCAPE)
			{
				MyOutputDebugStringA("Escape pressed during capture. Reverting to pre-capture state.\n");
				gAppState = APPSTATE_BEFORECAPTURE;	

				for (UINT8 Counter = 0; Counter < ArrayCount(gButtons); Counter++)
				{
					if (gButtons[Counter]->Id == BUTTON_NEW || gButtons[Counter]->Id == BUTTON_DELAY)
					{
						continue;
					}

					gButtons[Counter]->Enabled = FALSE;
				}

				ShowWindow(gCaptureWindowHandle, SW_HIDE);
				ShowWindow(gMainWindowHandle, SW_RESTORE);
				gCaptureSelectionRectangle.left   = 0;
				gCaptureSelectionRectangle.right  = 0;
				gCaptureSelectionRectangle.top    = 0;
				gCaptureSelectionRectangle.bottom = 0;
				
				RECT CurrentWindowPos = { 0 };
				GetWindowRect(gMainWindowHandle, &CurrentWindowPos);

				SetWindowPos(
					gMainWindowHandle,
					HWND_TOP,
					CurrentWindowPos.left,
					CurrentWindowPos.top,
					gStartingMainWindowWidth,
					gStartingMainWindowHeight,
					0);
			}
			break;
		}
		case WM_PAINT:
		{
			if (gCleanScreenShot != NULL)
			{
				PAINTSTRUCT PaintStruct = { 0 };			

				BeginPaint(Window, &PaintStruct);

				HDC BackBufferDC = CreateCompatibleDC(PaintStruct.hdc);

				HBITMAP ScreenShotCopy = CopyImage(gCleanScreenShot, IMAGE_BITMAP, 0, 0, 0);

				SelectObject(BackBufferDC, ScreenShotCopy);				

				if (MouseHasMovedWhileLeftMouseButtonWasDown)
				{					
					SelectObject(BackBufferDC, (HBRUSH)GetStockObject(NULL_BRUSH));

					Rectangle(BackBufferDC, gCaptureSelectionRectangle.left, gCaptureSelectionRectangle.top, gCaptureSelectionRectangle.right, gCaptureSelectionRectangle.bottom);
				}

				BLENDFUNCTION BlendFunction = { AC_SRC_OVER, 0, 128, AC_SRC_ALPHA };

				HDC AlphaDC = CreateCompatibleDC(PaintStruct.hdc);

									 // B G R A
				UCHAR BitmapBits[] = { 0xAA, 0xAA, 0xAA, 0xFF };

				// Create 1 32-bit pixel.
				HBITMAP AlphaBitmap = CreateBitmap(1, 1, 1, 32, BitmapBits);

				SelectObject(AlphaDC, AlphaBitmap);

				// This will darken the entire display area if nothing is selected.
				// Otherwise, it will darken everything to the right side of the selection rectangle.
				GdiAlphaBlend(
					BackBufferDC,
					max(gCaptureSelectionRectangle.left, gCaptureSelectionRectangle.right),
					0,
					gDisplayWidth - max(gCaptureSelectionRectangle.left, gCaptureSelectionRectangle.right),
					max(gDisplayHeight, gCaptureSelectionRectangle.top),
					AlphaDC,
					0,
					0,
					1,
					1,
					BlendFunction);

				// This darkens everything to the left side of the selection rectangle
				// Not needed if the user is not holding down the left mouse button
				if (LMouseButtonDown)
				{
					GdiAlphaBlend(
						BackBufferDC,
						0,
						0,
						min(gCaptureSelectionRectangle.left, gCaptureSelectionRectangle.right),
						max(gDisplayHeight, gCaptureSelectionRectangle.top),
						AlphaDC,
						0,
						0,
						1,
						1,
						BlendFunction);
				}

				// This darkens everything directly above the selection rectangle
				if (LMouseButtonDown)
				{
					GdiAlphaBlend(
						BackBufferDC,
						min(gCaptureSelectionRectangle.left, gCaptureSelectionRectangle.right),
						0,
						max(gCaptureSelectionRectangle.left, gCaptureSelectionRectangle.right) - min(gCaptureSelectionRectangle.left, gCaptureSelectionRectangle.right),
						min(gCaptureSelectionRectangle.top, gCaptureSelectionRectangle.bottom),
						AlphaDC,
						0,
						0,
						1,
						1,
						BlendFunction);
				}

				//// This darkens everything directly below the selection rectangle
				if (LMouseButtonDown)
				{
					GdiAlphaBlend(
						BackBufferDC,
						min(gCaptureSelectionRectangle.left, gCaptureSelectionRectangle.right),
						max(gCaptureSelectionRectangle.top, gCaptureSelectionRectangle.bottom),
						max(gCaptureSelectionRectangle.left, gCaptureSelectionRectangle.right) - min(gCaptureSelectionRectangle.left, gCaptureSelectionRectangle.right),
						max(gCaptureSelectionRectangle.bottom, gDisplayHeight),
						AlphaDC,
						0,
						0,
						1,
						1,
						BlendFunction);
				}

				// Finally, blit the fully-drawn backbuffer to the screen.
				BitBlt(PaintStruct.hdc, 0, 0, gDisplayWidth, gDisplayHeight, BackBufferDC, 0, 0, SRCCOPY);		
				
				if (ScreenShotCopy)
				{
					DeleteObject(ScreenShotCopy);
				}

				DeleteDC(BackBufferDC);
				
				if (AlphaBitmap)
				{
					DeleteObject(AlphaBitmap);
				}

				DeleteDC(AlphaDC);

				EndPaint(Window, &PaintStruct);
			}
			break;
		}
		case WM_LBUTTONDOWN:
		{
			MyOutputDebugStringA("[%s] Left mouse buttown down over capture window, capture selection beginning.\n", __FUNCTION__);
			
			MouseHasMovedWhileLeftMouseButtonWasDown = FALSE;

			gCaptureSelectionRectangle.left   = 0;
			gCaptureSelectionRectangle.right  = 0;
			gCaptureSelectionRectangle.bottom = 0;
			gCaptureSelectionRectangle.top    = 0;

			LMouseButtonDown = TRUE;
			
			InvalidateRect(Window, NULL, FALSE);
			UpdateWindow(gCaptureWindowHandle);

			POINT Mouse = { 0 };

			Mouse.x = GET_X_LPARAM(LParam);
			Mouse.y = GET_Y_LPARAM(LParam);

			gCaptureSelectionRectangle.left = Mouse.x;
			gCaptureSelectionRectangle.right = Mouse.x;
			gCaptureSelectionRectangle.top = Mouse.y;
			gCaptureSelectionRectangle.bottom = Mouse.y;

			break;
		}
		case WM_LBUTTONUP:
		{
			LMouseButtonDown = FALSE;
			CaptureWindow_OnLeftButtonUp();			
			break;
		}
		case WM_MOUSEMOVE:
		{
			if (LMouseButtonDown)
			{
				MouseHasMovedWhileLeftMouseButtonWasDown = TRUE;

				POINT Mouse = { 0 };

				Mouse.x = GET_X_LPARAM(LParam);
				Mouse.y = GET_Y_LPARAM(LParam);

				gCaptureSelectionRectangle.right = Mouse.x;
				gCaptureSelectionRectangle.bottom = Mouse.y;				

				InvalidateRect(Window, NULL, FALSE);
				UpdateWindow(gCaptureWindowHandle);
			}
			break;
		}
		default:
		{
			Result = DefWindowProc(Window, Message, WParam, LParam);
		}
	}

	return(Result);
}

void CaptureWindow_OnLeftButtonUp(void)
{
	MyOutputDebugStringA("[%s] Left mouse button up over capture window. Selection complete.\n", __FUNCTION__);

	if ((gCaptureSelectionRectangle.left != gCaptureSelectionRectangle.right) && (gCaptureSelectionRectangle.bottom != gCaptureSelectionRectangle.top)) 
	{
		MyOutputDebugStringA("[%s] Left mouse button was released with a valid capture region selected.\n", __FUNCTION__);

		gAppState = APPSTATE_AFTERCAPTURE;
		ShowWindow(gCaptureWindowHandle, SW_HIDE);
		ShowWindow(gMainWindowHandle, SW_RESTORE);

		RECT CurrentWindowPos = { 0 };

		// Includes both client and non-client area. In other words it returns the same values that I passed in to CreateWindowEx
		GetWindowRect(gMainWindowHandle, &CurrentWindowPos);

		// The reason behind all this is because depending on how the user dragged the selection rectangle, it might be inverted,
		// i.e. the right could actually be the left and the top could be the bottom.
		int PreviousWindowWidth  = CurrentWindowPos.right - CurrentWindowPos.left;
		int PreviousWindowHeight = CurrentWindowPos.bottom - CurrentWindowPos.top;
		
		gCaptureWidth  = (gCaptureSelectionRectangle.right - gCaptureSelectionRectangle.left) > 0 ? (gCaptureSelectionRectangle.right - gCaptureSelectionRectangle.left) + (gShouldAddDropShadow * 7) : (gCaptureSelectionRectangle.left - gCaptureSelectionRectangle.right) + (gShouldAddDropShadow * 7);
		gCaptureHeight = (gCaptureSelectionRectangle.bottom - gCaptureSelectionRectangle.top) > 0 ? (gCaptureSelectionRectangle.bottom - gCaptureSelectionRectangle.top) + (gShouldAddDropShadow * 7) : (gCaptureSelectionRectangle.top - gCaptureSelectionRectangle.bottom) + (gShouldAddDropShadow * 7);		

		int NewWindowWidth  = 0;
		int NewWindowHeight = 0;

		if (gCaptureWidth > PreviousWindowWidth - 20)
		{
			NewWindowWidth = gCaptureWidth + 20;
		}
		else
		{
			NewWindowWidth = PreviousWindowWidth;
		}

		NewWindowHeight = gCaptureHeight + PreviousWindowHeight + 7;

		SetWindowPos(
			gMainWindowHandle,
			HWND_TOP,
			CurrentWindowPos.left,
			CurrentWindowPos.top,
			NewWindowWidth,
			NewWindowHeight,
			0);

		gSnipStates[0] = CreateBitmap(gCaptureWidth, gCaptureHeight, 1, 32, NULL);

		HDC SnipDC = CreateCompatibleDC(NULL);

		SelectObject(SnipDC, gSnipStates[0]);

		HDC BigDC = CreateCompatibleDC(NULL);

		SelectObject(BigDC, gCleanScreenShot);
		
		BitBlt(SnipDC, 0, 0, gCaptureWidth, gCaptureHeight, BigDC, min(gCaptureSelectionRectangle.left, gCaptureSelectionRectangle.right), min(gCaptureSelectionRectangle.top, gCaptureSelectionRectangle.bottom), SRCCOPY);

		if (gShouldAddDropShadow)
		{
			HPEN Pen1 = CreatePen(PS_SOLID, 1, RGB(159, 159, 159));
			HPEN Pen2 = CreatePen(PS_SOLID, 1, RGB(172, 172, 172));
			HPEN Pen3 = CreatePen(PS_SOLID, 1, RGB(192, 192, 192));
			HPEN Pen4 = CreatePen(PS_SOLID, 1, RGB(215, 215, 215));
			HPEN Pen5 = CreatePen(PS_SOLID, 1, RGB(234, 234, 234));
			HPEN Pen6 = CreatePen(PS_SOLID, 1, RGB(245, 245, 245));
			HPEN Pen7 = CreatePen(PS_SOLID, 1, RGB(250, 250, 250));

			SelectObject(SnipDC, Pen1);
			MoveToEx(SnipDC, 0, gCaptureHeight - 7, NULL);
			LineTo(SnipDC, gCaptureWidth - 7, gCaptureHeight - 7);
			LineTo(SnipDC, gCaptureWidth - 7, -1);

			SelectObject(SnipDC, Pen2);
			MoveToEx(SnipDC, 0, gCaptureHeight - 6, NULL);
			LineTo(SnipDC, gCaptureWidth - 6, gCaptureHeight - 6);
			LineTo(SnipDC, gCaptureWidth - 6, -1);

			SelectObject(SnipDC, Pen3);
			MoveToEx(SnipDC, 0, gCaptureHeight - 5, NULL);
			LineTo(SnipDC, gCaptureWidth - 5, gCaptureHeight - 5);
			LineTo(SnipDC, gCaptureWidth - 5, -1);

			SelectObject(SnipDC, Pen4);
			MoveToEx(SnipDC, 0, gCaptureHeight - 4, NULL);
			LineTo(SnipDC, gCaptureWidth - 4, gCaptureHeight - 4);
			LineTo(SnipDC, gCaptureWidth - 4, -1);

			SelectObject(SnipDC, Pen5);
			MoveToEx(SnipDC, 0, gCaptureHeight - 3, NULL);
			LineTo(SnipDC, gCaptureWidth - 3, gCaptureHeight - 3);
			LineTo(SnipDC, gCaptureWidth - 3, -1);

			SelectObject(SnipDC, Pen6);
			MoveToEx(SnipDC, 0, gCaptureHeight - 2, NULL);
			LineTo(SnipDC, gCaptureWidth - 2, gCaptureHeight - 2);
			LineTo(SnipDC, gCaptureWidth - 2, -1);

			SelectObject(SnipDC, Pen7);
			MoveToEx(SnipDC, 0, gCaptureHeight - 1, NULL);
			LineTo(SnipDC, gCaptureWidth - 1, gCaptureHeight - 1);
			LineTo(SnipDC, gCaptureWidth - 1, -1);

			DeleteObject(Pen1);
			DeleteObject(Pen2);
			DeleteObject(Pen3);
			DeleteObject(Pen4);
			DeleteObject(Pen5);
			DeleteObject(Pen6);
			DeleteObject(Pen7);

			                                    //
			                                    //
			                                    //
			                                    //
			                                    //
			//////////////////////////////////////

			// Set the alpha to 0 on the shadowed bits

			//GetDIBits()
		}

		DeleteDC(BigDC);

		DeleteDC(SnipDC);

		for (UINT8 Counter = 0; Counter < ArrayCount(gButtons); Counter++)
		{
			gButtons[Counter]->Enabled = TRUE;
		}

		char TitleBuffer[64] = { 0 };
		(void)snprintf(TitleBuffer, sizeof(TitleBuffer), "SnipEx - Current Snip: %dx%d", gCaptureWidth, gCaptureHeight);

		SetWindowText(gMainWindowHandle, TitleBuffer);
		SetCursor(LoadCursor(NULL, IDC_ARROW));
	}

}

BOOL NewButton_Click(void)
{
	BOOL Result           = FALSE;
	RECT CurrentWindowPos = { 0 };
	char TitleBuffer[32]  = { 0 };
	HDC ScreenDC          = NULL;
	HDC MemoryDC          = NULL;

	for (UINT8 Counter = 0; Counter < ArrayCount(gButtons); Counter++)
	{
		gButtons[Counter]->SelectedTool = FALSE;
		gButtons[Counter]->State = BUTTONSTATE_NORMAL;
	}


	// Now the "capture window" comes to life. It is a duplicate of the entire display surface,
	// including multiple monitors. Take a screenshot of it, then overlay it on top of the real
	// desktop, and then allow the user to select a subsection of the screenshot with the mouse.

	KillTimer(gMainWindowHandle, DELAY_TIMER);
	gCurrentDelayCountdown = gStartingDelayCountdown;
	
	(void)snprintf(TitleBuffer, sizeof(TitleBuffer), "SnipEx");
	SetWindowText(gMainWindowHandle, TitleBuffer);
	
	GetWindowRect(gMainWindowHandle, &CurrentWindowPos);

	SetWindowPos(
		gMainWindowHandle,
		HWND_TOP,
		CurrentWindowPos.left,
		CurrentWindowPos.top,
		gStartingMainWindowWidth,
		gStartingMainWindowHeight,
		0);

	ShowWindow(gMainWindowHandle, SW_MINIMIZE);

	gAppState = APPSTATE_DURINGCAPTURE;

	ZeroMemory(&gCaptureSelectionRectangle, sizeof(RECT));

	for (INT8 SnipState = 0; SnipState < ArrayCount(gSnipStates) - 1; SnipState++)
	{
		if (gSnipStates[SnipState] != NULL)
		{
			DeleteObject(gSnipStates[SnipState]);
			gSnipStates[SnipState] = NULL;
		}
	}

	gCurrentSnipState = 0;

	// We just minimized the main window. Allow a brief moment for the minimize animation to finish before capturing the screen.
	Sleep(250);

	if (gCleanScreenShot != NULL)
	{
		DeleteObject(gCleanScreenShot);
	}

	ScreenDC = GetDC(NULL);

	if (ScreenDC == NULL)
	{
		MyMessageBoxA(NULL, "Error", MB_OK | MB_ICONERROR | MB_SYSTEMMODAL, "GetDC(NULL) failed with error 0x%lx!", GetLastError());		
		goto Cleanup;
	}

	MemoryDC = CreateCompatibleDC(ScreenDC);

	if (MemoryDC == NULL)
	{
		MyMessageBoxA(NULL, "Error", MB_OK | MB_ICONERROR | MB_SYSTEMMODAL, "CreateCompatibleDC(ScreenDC) failed with error 0x%lx!", GetLastError());		
		goto Cleanup;
	}

	gCleanScreenShot = CreateCompatibleBitmap(ScreenDC, gDisplayWidth, gDisplayHeight);

	if (gCleanScreenShot == NULL)
	{
		MyMessageBoxA(NULL, "Error", MB_OK | MB_ICONERROR | MB_SYSTEMMODAL, "CreateCompatibleBitmap failed with error 0x%lx!", GetLastError());		
		goto Cleanup;
	}

	SelectObject(MemoryDC, gCleanScreenShot);

	BitBlt(MemoryDC, 0, 0, gDisplayWidth, gDisplayHeight, ScreenDC, gDisplayLeft, gDisplayTop, SRCCOPY);

	ShowWindow(gCaptureWindowHandle, SW_SHOW);

	SetWindowPos(gCaptureWindowHandle, HWND_TOP, gDisplayLeft, gDisplayTop, gDisplayWidth, gDisplayHeight, SWP_NOOWNERZORDER | SWP_FRAMECHANGED);

	SetForegroundWindow(gCaptureWindowHandle);

	Result = TRUE;

Cleanup:

	if (MemoryDC != NULL)
	{
		DeleteDC(MemoryDC);
	}

	if (ScreenDC != NULL)
	{
		ReleaseDC(NULL, ScreenDC);
	}

	return(Result);
}

// Returns TRUE if the snip was saved. Returns FALSE if there was an error or if user cancelled.
BOOL SaveButton_Click(void)
{
	HRESULT COMError = 0;
	IFileSaveDialog* DialogInterface = NULL;
	UINT SelectedFileTypeIndex = 0;
	
	const COMDLG_FILTERSPEC FileTypeFilters[] = { 
		{ L"Portable Network Graphics (PNG)", L"*.png" }, 
		{ L"32-bpp Bitmap", L"*.bmp" } 
	};

	IShellItem* ResultItem = NULL;
	
	LPOLESTR FilePathFromDialogW = NULL;
	
	wchar_t FinalFilePathW[MAX_PATH] = { 0 };

	BOOL Result = FALSE;

	COMError = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);

	if (FAILED(COMError))
	{
		MyMessageBoxA(NULL, "Error", MB_OK | MB_ICONERROR | MB_SYSTEMMODAL, "Failed to initialize COM! Error 0x%lx", COMError);
		goto Cleanup;		
	}	

	COMError = CoCreateInstance(&CLSID_FileSaveDialog, NULL, CLSCTX_INPROC_SERVER, &IID_IFileSaveDialog, (void**)&DialogInterface);

	if (FAILED(COMError))
	{
		MyMessageBoxA(NULL, "Error", MB_OK | MB_ICONERROR | MB_SYSTEMMODAL, "Failed to create COM instance of IFileDialog! Error 0x%lx", COMError);
		goto Cleanup;
	}	

	DialogInterface->lpVtbl->SetFileTypes(DialogInterface, ArrayCount(FileTypeFilters), FileTypeFilters);
	
	DialogInterface->lpVtbl->SetFileTypeIndex(DialogInterface, 1); // 1-based array, does not start at 0	

	DialogInterface->lpVtbl->Show(DialogInterface, gMainWindowHandle);
	
	DialogInterface->lpVtbl->GetResult(DialogInterface, &ResultItem);

	if (ResultItem == NULL)
	{
		// User probably hit cancel.
		goto Cleanup;
	}

	ResultItem->lpVtbl->GetDisplayName(ResultItem, SIGDN_FILESYSPATH, &FilePathFromDialogW);

	if (wcslen(FilePathFromDialogW) <= 3 || wcslen(FilePathFromDialogW) >= MAX_PATH - 5)
	{
		MyMessageBoxA(NULL, "Error", MB_OK | MB_ICONERROR | MB_SYSTEMMODAL, "File path was too short or too long!");
		goto Cleanup;
	}

	wcscpy_s(FinalFilePathW, MAX_PATH, FilePathFromDialogW);

	DialogInterface->lpVtbl->GetFileTypeIndex(DialogInterface, &SelectedFileTypeIndex);

	switch (SelectedFileTypeIndex)
	{
		case 1:
		{
			if (wcslen(FinalFilePathW) < 5)
			{			
				wcscat_s(FinalFilePathW, MAX_PATH, L".png");
			}
			else
			{
				if (_wcsicmp(&FinalFilePathW[wcslen(FinalFilePathW) - 4], L".png") != 0)
				{
					wcscat_s(FinalFilePathW, MAX_PATH, L".png");
				}				
			}

			MyOutputDebugStringW(L"[%s] Line %d: Attempting to save file %s\n", __FUNCTIONW__, __LINE__, FinalFilePathW);

			if (SavePngToFile(FinalFilePathW) == FALSE)
			{
				goto Cleanup;
			}

			break;
		}
		case 2:
		{
			if (wcslen(FinalFilePathW) < 5)
			{
				wcscat_s(FinalFilePathW, MAX_PATH, L".bmp");
			}
			else
			{
				if (_wcsicmp(&FinalFilePathW[wcslen(FinalFilePathW) - 4], L".bmp") != 0)
				{
					wcscat_s(FinalFilePathW, MAX_PATH, L".bmp");
				}				
			}

			MyOutputDebugStringW(L"[%s] Line %d: Attempting to save file %s\n", __FUNCTIONW__, __LINE__, FinalFilePathW);

			if (SaveBitmapToFile(FinalFilePathW) == FALSE)
			{
				goto Cleanup;
			}

			break;
		}
		default:
		{
			MyMessageBoxA(NULL, "Error", MB_OK | MB_ICONERROR | MB_SYSTEMMODAL, "File type selection was not in the expected range of values!");
			goto Cleanup;
		}
	}

	// If nothing has gone wrong up to this point, set Success to TRUE.
	Result = TRUE;

	Cleanup:

	if (ResultItem != NULL)
	{
		ResultItem->lpVtbl->Release(ResultItem);
	}

	if (DialogInterface != NULL)
	{
		DialogInterface->lpVtbl->Release(DialogInterface); 
	}

	CoUninitialize();

	return(Result);
}

BOOL CopyButton_Click(void)
{
	BOOL Result           = FALSE;
	BITMAP Bitmap         = { 0 };
	HBITMAP ClipboardCopy = NULL;
	HDC SourceDC          = NULL;
	HDC DestinationDC     = NULL;

	if (gSnipStates[gCurrentSnipState] == NULL)
	{
		MyOutputDebugStringA("[%s] gSnipStates[gCurrentSnipState] is NULL. This is a bug.\n", __FUNCTION__);
		goto Cleanup;
	}

	GetObject(gSnipStates[gCurrentSnipState], sizeof(BITMAP), &Bitmap);

	ClipboardCopy = CreateBitmap(Bitmap.bmWidth, Bitmap.bmHeight, 1, 32, NULL);

	SourceDC = CreateCompatibleDC(NULL);
	DestinationDC = CreateCompatibleDC(NULL);

	SelectObject(SourceDC, gSnipStates[gCurrentSnipState]);
	SelectObject(DestinationDC, ClipboardCopy);

	BitBlt(DestinationDC, 0, 0, Bitmap.bmWidth, Bitmap.bmHeight, SourceDC, 0, 0, SRCCOPY);

	if (OpenClipboard(gMainWindowHandle) == 0)
	{		
		MyMessageBoxA(NULL, "Error", MB_OK | MB_ICONERROR | MB_SYSTEMMODAL, "OpenClipboard failed with error 0x%lx!", GetLastError());
		goto Cleanup;
	}

	if (EmptyClipboard() == 0)
	{		
		MyMessageBoxA(NULL, "Error", MB_OK | MB_ICONERROR | MB_SYSTEMMODAL, "EmptyClipboard failed with error 0x%lx!", GetLastError());
		goto Cleanup;
	}

	if (SetClipboardData(CF_BITMAP, ClipboardCopy) == NULL)
	{		
		MyMessageBoxA(NULL, "Error", MB_OK | MB_ICONERROR | MB_SYSTEMMODAL, "SetClipboardData failed with error 0x%lx!", GetLastError());
		goto Cleanup;
	}	

	Result = TRUE;

Cleanup:

	if (SourceDC != NULL)
	{
		DeleteDC(SourceDC);
	}

	if (DestinationDC != NULL)
	{
		DeleteDC(DestinationDC);
	}

	if (ClipboardCopy != NULL)
	{
		DeleteObject(ClipboardCopy);
	}

	CloseClipboard();

	return(Result);
}

BOOL SaveBitmapToFile(_In_ wchar_t* FilePath)
{
	BOOL Success = FALSE;
	HANDLE FileHandle = NULL;
	BITMAPFILEHEADER BitmapFileHeader = { 0 };
	PBITMAPINFOHEADER BitmapInfoHeaderPointer = { 0 };
	BITMAP Bitmap = { 0 };
	PBITMAPINFO BitmapInfoPointer = { 0 };
	WORD cClrBits = 0;
	LPBYTE Bits = { 0 };
	HDC DC = NULL;
	DWORD BytesWritten = 0;
	DWORD TotalBytes = 0;
	DWORD IncrementalBytes = 0;
	BYTE* BytePointer = 0;

	FileHandle = CreateFileW(FilePath, GENERIC_READ | GENERIC_WRITE, (DWORD)0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);

	if (FileHandle == INVALID_HANDLE_VALUE)
	{
		MyMessageBoxA(NULL, "Error", MB_OK | MB_ICONERROR | MB_SYSTEMMODAL, "Failed to create bitmap file! Error 0x%lx", GetLastError());
		goto Cleanup;
	}

	if (GetObject(gSnipStates[gCurrentSnipState], sizeof(BITMAP), &Bitmap) == 0)
	{
		MyOutputDebugStringA("[%s] Line %d: GetObject failed!\n", __FUNCTION__, __LINE__);
		goto Cleanup;
	}

	// Convert the color format to a count of bits.  
	cClrBits = (WORD)(Bitmap.bmPlanes * Bitmap.bmBitsPixel);

	if (cClrBits == 1)
	{
		cClrBits = 1;
	}
	else if (cClrBits <= 4)
	{
		cClrBits = 4;
	}
	else if (cClrBits <= 8)
	{
		cClrBits = 8;
	}
	else if (cClrBits <= 16)
	{
		cClrBits = 16;
	}
	else if (cClrBits <= 24)
	{
		cClrBits = 24;
	}
	else
	{
		cClrBits = 32;
	}

	// Allocate memory for the BITMAPINFO structure. (This structure  
	// contains a BITMAPINFOHEADER structure and an array of RGBQUAD  
	// data structures.)  

	if (cClrBits < 24)
	{
		BitmapInfoPointer = (PBITMAPINFO)LocalAlloc(LPTR, sizeof(BITMAPINFOHEADER) + sizeof(RGBQUAD) * (1i64 << cClrBits));
	}
	else
	{
		// There is no RGBQUAD array for these formats: 24-bit-per-pixel or 32-bit-per-pixel 
		BitmapInfoPointer = (PBITMAPINFO)LocalAlloc(LPTR, sizeof(BITMAPINFOHEADER));
	}

	if (BitmapInfoPointer == NULL)
	{
		MyMessageBoxA(NULL, "Error", MB_OK | MB_ICONERROR | MB_SYSTEMMODAL, "Failed to allocate memory! Error 0x%lx", GetLastError());
		MyOutputDebugStringA("[%s] LocalAlloc failed!\n", __FUNCTION__);
		goto Cleanup;
	}

	BitmapInfoPointer->bmiHeader.biSize     = sizeof(BITMAPINFOHEADER);
	BitmapInfoPointer->bmiHeader.biWidth    = Bitmap.bmWidth;
	BitmapInfoPointer->bmiHeader.biHeight   = Bitmap.bmHeight;
	BitmapInfoPointer->bmiHeader.biPlanes   = Bitmap.bmPlanes;
	BitmapInfoPointer->bmiHeader.biBitCount = Bitmap.bmBitsPixel;
	
	if (cClrBits < 24)
	{
		BitmapInfoPointer->bmiHeader.biClrUsed = (1 << cClrBits);
	}

	// If the bitmap is not compressed, set the BI_RGB flag.  
	BitmapInfoPointer->bmiHeader.biCompression = BI_RGB;

	// Compute the number of bytes in the array of color  
	// indices and store the result in biSizeImage.  
	// The width must be DWORD aligned unless the bitmap is RLE 
	// compressed. 
	BitmapInfoPointer->bmiHeader.biSizeImage = ((BitmapInfoPointer->bmiHeader.biWidth * cClrBits + 31) & ~31) / 8 * BitmapInfoPointer->bmiHeader.biHeight;
	// Set biClrImportant to 0, indicating that all of the  
	// device colors are important.  
	BitmapInfoPointer->bmiHeader.biClrImportant = 0;	

	BitmapInfoHeaderPointer = (PBITMAPINFOHEADER)BitmapInfoPointer;
	Bits = (LPBYTE)GlobalAlloc(GMEM_FIXED, BitmapInfoHeaderPointer->biSizeImage);

	if (Bits == 0)
	{
		MyMessageBoxA(NULL, "Error", MB_OK | MB_ICONERROR | MB_SYSTEMMODAL, "Failed to allocate memory! Error 0x%lx", GetLastError());
		MyOutputDebugStringA("[%s] GlobalAlloc failed!\n", __FUNCTION__);
		goto Cleanup;
	}
	
	DC = CreateCompatibleDC(NULL);
	
	SelectObject(DC, gSnipStates[gCurrentSnipState]);

	// Retrieve the color table (RGBQUAD array) and the bits (array of palette indices) from the DIB.  
	GetDIBits(DC, gSnipStates[gCurrentSnipState], 0, (WORD)BitmapInfoHeaderPointer->biHeight, Bits, BitmapInfoPointer, DIB_RGB_COLORS);

	BitmapFileHeader.bfType = 0x4d42;	// "BM"
	BitmapFileHeader.bfReserved1 = 0;
	BitmapFileHeader.bfReserved2 = 0;
	BitmapFileHeader.bfSize = (DWORD)(sizeof(BITMAPFILEHEADER) + BitmapInfoHeaderPointer->biSize + BitmapInfoHeaderPointer->biClrUsed * sizeof(RGBQUAD) + BitmapInfoHeaderPointer->biSizeImage);

	// Compute the offset to the array of color indices.  
	BitmapFileHeader.bfOffBits = (DWORD)sizeof(BITMAPFILEHEADER) + BitmapInfoHeaderPointer->biSize + BitmapInfoHeaderPointer->biClrUsed * sizeof(RGBQUAD);

	WriteFile(FileHandle, (LPVOID)&BitmapFileHeader, sizeof(BITMAPFILEHEADER), (LPDWORD)&BytesWritten, NULL);

	// Copy the BITMAPINFOHEADER and RGBQUAD array into the file.
	WriteFile(FileHandle, (LPVOID)BitmapInfoHeaderPointer, sizeof(BITMAPINFOHEADER) + BitmapInfoHeaderPointer->biClrUsed * sizeof(RGBQUAD), (LPDWORD)&BytesWritten, (NULL));

	// Copy the array of color indices into the .BMP file.  
	TotalBytes = IncrementalBytes = BitmapInfoHeaderPointer->biSizeImage;
	BytePointer = Bits;
	WriteFile(FileHandle, (LPSTR)BytePointer, (int)IncrementalBytes, (LPDWORD)&BytesWritten, NULL);

	// If nothing has failed up to this point, set Success to TRUE.
	Success = TRUE;

	Cleanup:

	if (Bits != NULL)
	{
		GlobalFree(Bits);
	}

	if (BitmapInfoPointer != NULL)
	{
		LocalFree(BitmapInfoPointer);
	}

	if (DC != NULL)
	{
		DeleteDC(DC);
	}

	if (FileHandle != INVALID_HANDLE_VALUE)
	{
		CloseHandle(FileHandle);
	}

	if (Success == TRUE)
	{
		MyOutputDebugStringA("[%s] Returning successfully.\n", __FUNCTION__);
		return(TRUE);
	}
	else
	{
		MyOutputDebugStringA("[%s] Returning failure!\n", __FUNCTION__);
		return(FALSE);
	}
}

BOOL SavePngToFile(_In_ wchar_t* FilePath)
{
	BOOL      Result       = FALSE;
	CLSID     PngCLSID     = { 0 };
	HRESULT   Error        = 0;
	HMODULE   GDIPlus      = LoadLibrary("GDIPlus.dll");
	ULONG_PTR GdiplusToken = 0;
	ULONG*    GdipBitmap   = 0;

	if (GDIPlus == NULL)
	{
		MyMessageBoxA(gMainWindowHandle, "Error", MB_OK | MB_ICONERROR | MB_SYSTEMMODAL, "LoadLibrary(\"GDIPlus.dll\") failed with error 0x%lx!", GetLastError());
		goto Cleanup;
	}
	
	Error = CLSIDFromString(L"{557CF406-1A04-11D3-9A73-0000F81EF32E}", &PngCLSID);

	if (Error != NOERROR)
	{
		MyMessageBoxA(gMainWindowHandle, "Error", MB_OK | MB_ICONERROR | MB_SYSTEMMODAL, "CLSIDFromString failed with error 0x%lx!", Error);
		goto Cleanup;
	}

	GdiplusStartup = (int (WINAPI*)(ULONG_PTR*, struct GdiplusStartupInput*, void*))GetProcAddress(GDIPlus, "GdiplusStartup");

	if (GdiplusStartup == NULL)
	{
		MyMessageBoxA(gMainWindowHandle, "Error", MB_OK | MB_ICONERROR | MB_SYSTEMMODAL, "Failed to load GdiplusStartup from gdiplus.dll!");
		goto Cleanup;
	}

	GdiplusShutdown = (int (WINAPI*)(ULONG_PTR))GetProcAddress(GDIPlus, "GdiplusShutdown");	

	if (GdiplusShutdown == NULL)
	{
		MyMessageBoxA(gMainWindowHandle, "Error", MB_OK | MB_ICONERROR | MB_SYSTEMMODAL, "Failed to load GdiplusShutdown from gdiplus.dll!");
		goto Cleanup;
	}

	GdipCreateBitmapFromHBITMAP = (int (WINAPI*)(HBITMAP, HPALETTE hPalette, ULONG** Bitmap))GetProcAddress(GDIPlus, "GdipCreateBitmapFromHBITMAP");

	if (GdipCreateBitmapFromHBITMAP == NULL)
	{
		MyMessageBoxA(gMainWindowHandle, "Error", MB_OK | MB_ICONERROR | MB_SYSTEMMODAL, "Failed to load GdipCreateBitmapFromHBITMAP from gdiplus.dll!");
		goto Cleanup;
	}

	GdipSaveImageToFile = (int (WINAPI*)(ULONG*, const WCHAR*, const CLSID*, const EncoderParameters*))GetProcAddress(GDIPlus, "GdipSaveImageToFile");

	if (GdipSaveImageToFile == NULL)
	{
		MyMessageBoxA(gMainWindowHandle, "Error", MB_OK | MB_ICONERROR | MB_SYSTEMMODAL, "Failed to load GdipSaveImageToFile from gdiplus.dll!");
		goto Cleanup;
	}

	GdipDisposeImage = (int (WINAPI*)(ULONG*))GetProcAddress(GDIPlus, "GdipDisposeImage");

	if (GdipDisposeImage == NULL)
	{
		MyMessageBoxA(gMainWindowHandle, "Error", MB_OK | MB_ICONERROR | MB_SYSTEMMODAL, "Failed to load GdipDisposeImage from gdiplus.dll!");
		goto Cleanup;
	}

	GdiplusStartupInput GdiplusStartupInput = { 0 };	

	GdiplusStartupInput.GdiplusVersion           = 1;
	GdiplusStartupInput.DebugEventCallback       = NULL;
	GdiplusStartupInput.SuppressBackgroundThread = FALSE;
	GdiplusStartupInput.SuppressExternalCodecs   = FALSE;
	
	if ((Error = GdiplusStartup(&GdiplusToken, &GdiplusStartupInput, NULL)) != S_OK)
	{
		MyMessageBoxA(gMainWindowHandle, "Error", MB_OK | MB_ICONERROR | MB_SYSTEMMODAL, "GdiplusStartup failed with GdiPlus Status code 0x%lx!", Error);
		goto Cleanup;
	}	

	if ((Error = GdipCreateBitmapFromHBITMAP(gSnipStates[gCurrentSnipState], NULL, &GdipBitmap)) != 0)
	{
		MyMessageBoxA(gMainWindowHandle, "Error", MB_OK | MB_ICONERROR | MB_SYSTEMMODAL, "GdipCreateBitmapFromHBITMAP failed with GdiPlus status code 0x%lx!", Error);
		goto Cleanup;
	}	

	EncoderParameters EncoderParams = { 0 };

	LONG Compression = 5;

	EncoderParams.Count = 1;
	EncoderParams.Parameter[0].NumberOfValues = 1;
	EncoderParams.Parameter[0].Guid = gEncoderCompressionGuid;
	EncoderParams.Parameter[0].Type = EncoderParameterValueTypeLong;
	EncoderParams.Parameter[0].Value = &Compression;

	if ((Error = GdipSaveImageToFile(GdipBitmap, FilePath, &PngCLSID, &EncoderParams)) != 0)
	{
		MyMessageBoxA(gMainWindowHandle, "Error", MB_OK | MB_ICONERROR | MB_SYSTEMMODAL, "GdipSaveImageToFile failed with GdiPlus status code 0x%lx!", Error);
		goto Cleanup;
	}


	Result = TRUE;
	
	Cleanup:

	if (GdipBitmap != NULL)
	{
		GdipDisposeImage(GdipBitmap);
	}

	if (GdiplusToken != 0)
	{
		GdiplusShutdown(GdiplusToken);
	}

	if (GDIPlus != NULL)
	{
		FreeLibrary(GDIPlus);
	}	

	return(Result);
}

// We will compare CommandLine to the value of the Debugger registry value
void AddReplaceSnippingToolMenuItem(_In_ HINSTANCE Instance)
{
	BOOL  AlreadyReplaced = FALSE;
	HKEY  IFEOKey = NULL;
	HKEY  SnipExKey = NULL;
	DWORD SnipExKeyDisposition = 0;
	char  DebuggerValue[256] = { 0 };
	DWORD ValueSize = sizeof(DebuggerValue);

	// This key should always exist. Something is very wrong if we can't open it.
	if (RegOpenKeyEx(HKEY_LOCAL_MACHINE, "SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Image File Execution Options", 0, KEY_QUERY_VALUE, &IFEOKey) != ERROR_SUCCESS)
	{		
		MyMessageBoxA(NULL, "Error", MB_OK | MB_ICONERROR | MB_SYSTEMMODAL, "Unable to read Image File Execution Options registry key! Error: 0x%lx", GetLastError());
		goto Cleanup;
	}

	// If this subkey is not present, then that is OK, it just means it has never been added before
	if (RegCreateKeyEx(IFEOKey, "SnippingTool.exe", 0, NULL, 0, KEY_QUERY_VALUE, NULL, &SnipExKey, &SnipExKeyDisposition) == ERROR_SUCCESS)
	{
		RegQueryValueEx(SnipExKey, "Debugger", NULL, NULL, (LPBYTE)&DebuggerValue, &ValueSize);

		if (strlen(DebuggerValue) > 0)
		{
			// Does the file specified by the "Debugger" value actually exist?

			DWORD FileAttributes = GetFileAttributes(DebuggerValue);

			if (FileAttributes != INVALID_FILE_ATTRIBUTES && !(FileAttributes & FILE_ATTRIBUTE_DIRECTORY))
			{
				AlreadyReplaced = TRUE;
			}
		}
	}

	HMENU SystemMenu = GetSystemMenu(gMainWindowHandle, FALSE);

	AppendMenu(SystemMenu, MF_SEPARATOR, 0, NULL);

	UINT16 Command = 0;
	char Text[64] = { 0 };

	if (AlreadyReplaced)
	{
		Command = SYSCMD_RESTORE;
		strcpy_s(Text, sizeof(Text), "Restore Windows Snipping Tool");
	}
	else
	{
		Command = SYSCMD_REPLACE;
		strcpy_s(Text, sizeof(Text), "Replace Windows Snipping Tool with SnipEx");
	}

	AppendMenu(SystemMenu, MF_STRING, Command, Text);
	
	// Only works with 8bpp bitmaps
	gUACIcon = (HBITMAP)LoadImage(Instance, MAKEINTRESOURCE(IDB_UAC), IMAGE_BITMAP, 0, 0, LR_LOADTRANSPARENT | LR_LOADMAP3DCOLORS | LR_SHARED);
	
	if (gUACIcon != NULL)
	{
		MENUITEMINFO MenuItemInfo = { sizeof(MENUITEMINFO) };

		MenuItemInfo.fMask = MIIM_BITMAP;
		MenuItemInfo.hbmpItem = gUACIcon;

		SetMenuItemInfo(SystemMenu, Command, FALSE, &MenuItemInfo);
	}
	else
	{
		MyOutputDebugStringA("[%s] Line %d: Attempting to load the UAC icon failed! Error 0x%lx\n", __FUNCTION__, __LINE__, GetLastError());
	}

Cleanup:

	if (SnipExKey != NULL)
	{
		RegCloseKey(SnipExKey);
	}

	if (IFEOKey != NULL)
	{
		RegCloseKey(IFEOKey);
	}
}

void AddDropShadowToolMenuItem(void)
{
	HMENU SystemMenu  = GetSystemMenu(gMainWindowHandle, FALSE);
	HKEY  SoftwareKey = NULL;
	HKEY  SnipExKey   = NULL;

	DWORD SnipExKeyDisposition = 0;
	DWORD DropShadowValue = 0;
	DWORD ValueSize = sizeof(DWORD);	

	// This key should always exist. Something is very wrong if we can't open it.
	if (RegOpenKeyEx(HKEY_CURRENT_USER, "SOFTWARE", 0, KEY_QUERY_VALUE, &SoftwareKey) != ERROR_SUCCESS)
	{
		MyMessageBoxA(NULL, "Error", MB_OK | MB_ICONERROR | MB_SYSTEMMODAL, "Unable to read HKCU\\SOFTWARE registry key! Error: 0x%lx", GetLastError());
		goto Cleanup;
	}

	// If this subkey is not present, then that is OK, it just means it has never been added before
	if (RegCreateKeyEx(SoftwareKey, "SnipEx", 0, NULL, 0, KEY_QUERY_VALUE, NULL, &SnipExKey, &SnipExKeyDisposition) == ERROR_SUCCESS)
	{
		RegQueryValueEx(SnipExKey, "DropShadow", NULL, NULL, (LPBYTE)&DropShadowValue, &ValueSize);

		if (DropShadowValue > 0)
		{;
			AppendMenu(SystemMenu, MF_STRING | MF_CHECKED, SYSCMD_SHADOW, "Drop Shadow Effect");
			gShouldAddDropShadow = TRUE;
		}
		else
		{
			AppendMenu(SystemMenu, MF_STRING | MF_UNCHECKED, SYSCMD_SHADOW, "Drop Shadow Effect");
			gShouldAddDropShadow = FALSE;
		}
	}
	else
	{
		MyMessageBoxA(gMainWindowHandle, "Error", MB_OK | MB_ICONERROR | MB_SYSTEMMODAL, "Unable to open or create the HKCU\\SnipEx registry key. The drop shadow effect will be unavailable. :(");
		gShouldAddDropShadow = FALSE;
	}

Cleanup:

	if (SnipExKey != NULL)
	{
		RegCloseKey(SnipExKey);
	}

	if (SoftwareKey != NULL)
	{
		RegCloseKey(SoftwareKey);
	}
}

BOOL IsAppRunningElevated(void)
{
	MyOutputDebugStringA("[%s] Line %d: Entered.\n", __FUNCTION__, __LINE__);

	BOOL  IsElevated = FALSE;
	DWORD Error = ERROR_SUCCESS;
	PSID  AdministratorsGroup = NULL;

	// Allocate and initialize a SID of the administrators group.

	SID_IDENTIFIER_AUTHORITY NtAuthority = SECURITY_NT_AUTHORITY;
	if (!AllocateAndInitializeSid(
		&NtAuthority,
		2,
		SECURITY_BUILTIN_DOMAIN_RID,
		DOMAIN_ALIAS_RID_ADMINS,
		0, 0, 0, 0, 0, 0,
		&AdministratorsGroup))
	{
		Error = GetLastError();
		goto Cleanup;
	}

	// Determine whether the SID of administrators group is enabled in 

	// the primary access token of the process.

	if (!CheckTokenMembership(NULL, AdministratorsGroup, &IsElevated))
	{
		Error = GetLastError();
		goto Cleanup;
	}

Cleanup:
	// Centralized cleanup for all allocated resources.

	if (AdministratorsGroup)
	{
		FreeSid(AdministratorsGroup);
		AdministratorsGroup = NULL;
	}

	// Throw the error if something failed in the function.

	if (ERROR_SUCCESS != Error)
	{
		MyMessageBoxA(gMainWindowHandle, "Error", MB_OK | MB_ICONERROR | MB_SYSTEMMODAL, "Error checking for UAC Elevation! Error: 0x%lx", Error);
	}

	MyOutputDebugStringA("[%s] Line %d: Process is elevated: %d.\n", __FUNCTION__, __LINE__, IsElevated);
	return IsElevated;
}

// MessageBox enhanced with varargs.
int MyMessageBoxA(_In_opt_ HWND ParentWindow, _In_opt_ LPCTSTR Caption, _In_ UINT Type, _In_ LPCTSTR Text, _In_ ...)
{
	char Buffer[512] = { 0 };

	va_list VarArgPointer = NULL;

	va_start(VarArgPointer, Text);
	_vsnprintf_s(Buffer, _countof(Buffer), _TRUNCATE, Text, VarArgPointer);
	va_end(VarArgPointer);

	Buffer[strlen(Buffer)] = '\0';

	int MessageBoxError = MessageBoxA(ParentWindow, Buffer, Caption, Type);

	return(MessageBoxError);
}

// OutputDebugString enhanced with varargs. Only works in debug builds.
void MyOutputDebugStringA(_In_ char* Message, _In_ ...)
{
	#if _DEBUG

	char Buffer[512] = { 0 };
	
	va_list VarArgPointer = NULL;

	va_start(VarArgPointer, Message);
	_vsnprintf_s(Buffer, _countof(Buffer), _TRUNCATE, Message, VarArgPointer);
	va_end(VarArgPointer);

	OutputDebugStringA(Buffer);
	#else
	UNREFERENCED_PARAMETER(Message);
	#endif
}

// OutputDebugString enhanced with varargs. Only works in debug builds.
void MyOutputDebugStringW(_In_ wchar_t* Message, _In_ ...)
{
	#if _DEBUG
	wchar_t Buffer[512] = { 0 };

	va_list VarArgPointer = NULL;

	va_start(VarArgPointer, Message);
	_vsnwprintf_s(Buffer, _countof(Buffer), _TRUNCATE, Message, VarArgPointer);
	va_end(VarArgPointer);

	OutputDebugStringW(Buffer);
	#else
	UNREFERENCED_PARAMETER(Message);
	#endif
}

BOOL AdjustForCustomScaling(void)
{
	// We need to do some adjusting for non-default DPI monitors/scaling factors.

	UINT DPIx = 0;
	UINT DPIy = 0;

	if (GetDpiForMonitor(MonitorFromWindow(GetDesktopWindow(), MONITOR_DEFAULTTOPRIMARY), MDT_EFFECTIVE_DPI, &DPIx, &DPIy) != S_OK)
	{
		MyMessageBoxA(NULL, "Error", MB_OK | MB_ICONERROR | MB_SYSTEMMODAL, "Unable to determine the monitor DPI of your primary monitor!");
		return(FALSE);
	}

	MyOutputDebugStringA("[%s] Line %d: Detected a monitor DPI of %d.\n", __FUNCTION__, __LINE__, DPIx);

	// Pct    DPI
	// 100% =  96 X
	// 105% = 101
	// 110% = 106
	// 115% = 110
	// 120% = 115
	// 125% = 120
	// 130% = 125
	// 135% = 130
	// 140% = 134
	// 145% = 139
	// 150% = 144
	// 155% = 149
	// 160% = 154
	// 165% = 158
	// 170% = 163
	// 175% = 168
	// 180% = 173
	// 185% = 178
	// 190% = 182
	// 195% = 187
	// 200% = 192
	

	if (DPIx == 96)
	{
		return(TRUE);
	}

	if (DPIx > 96 && DPIx <= 106)
	{
		gStartingMainWindowHeight += 2;
		gStartingMainWindowWidth  += 2;
	}
	else if (DPIx > 106 && DPIx <= 110)
	{
		gStartingMainWindowHeight += 5;
		gStartingMainWindowWidth  += 3;
	}
	else if (DPIx > 110 && DPIx <= 115)
	{
		gStartingMainWindowHeight += 6;
		gStartingMainWindowWidth  += 3;
	}
	else if (DPIx > 115 && DPIx <= 120)
	{
		gStartingMainWindowHeight += 8;
		gStartingMainWindowWidth  += 3;
	}
	else if (DPIx > 120 && DPIx <= 125)
	{
		gStartingMainWindowHeight += 10;
		gStartingMainWindowWidth  += 3;
	}
	else if (DPIx > 125 && DPIx <= 130)
	{
		gStartingMainWindowHeight += 10;
		gStartingMainWindowWidth  += 3;
	}
	else if (DPIx > 130 && DPIx <= 134)
	{
		gStartingMainWindowHeight += 15;
		gStartingMainWindowWidth  += 5;
	}
	else if (DPIx > 134 && DPIx <= 139)
	{
		gStartingMainWindowHeight += 15;
		gStartingMainWindowWidth  += 5;
	}
	else if (DPIx > 139 && DPIx <= 144)
	{
		gStartingMainWindowHeight += 17;
		gStartingMainWindowWidth  += 7;
	}
	else if (DPIx > 144 && DPIx <= 149)
	{
		gStartingMainWindowHeight += 18;
		gStartingMainWindowWidth  += 7;
	}
	else if (DPIx > 149 && DPIx <= 154)
	{
		gStartingMainWindowHeight += 19;
		gStartingMainWindowWidth  += 7;
	}
	else if (DPIx > 154 && DPIx <= 158)
	{
		gStartingMainWindowHeight += 22;
		gStartingMainWindowWidth  += 9;
	}
	else if (DPIx > 158 && DPIx <= 163)
	{
		gStartingMainWindowHeight += 23;
		gStartingMainWindowWidth  += 9;
	}
	else if (DPIx > 163 && DPIx <= 168)
	{
		gStartingMainWindowHeight += 25;
		gStartingMainWindowWidth  += 9;
	}
	else if (DPIx > 168 && DPIx <= 173)
	{
		gStartingMainWindowHeight += 26;
		gStartingMainWindowWidth  += 9;
	}
	else if (DPIx > 173 && DPIx <= 178)
	{
		gStartingMainWindowHeight += 27;
		gStartingMainWindowWidth  += 9;
	}
	else if (DPIx > 178 && DPIx <= 182)
	{
		gStartingMainWindowHeight += 30;
		gStartingMainWindowWidth  += 10;
	}
	else if (DPIx > 182 && DPIx <= 187)
	{
		gStartingMainWindowHeight += 31;
		gStartingMainWindowWidth  += 10;
	}
	else if (DPIx > 187 && DPIx <= 192)
	{
		gStartingMainWindowHeight += 32;
		gStartingMainWindowWidth  += 10;
	}
	else
	{
		MessageBox(NULL, "Unable to deal with your custom scaling level. I can only handle up to 200% scaling. Contact me at ryanries09@gmail.com if you want me to add support for your scaling level.", "Error", MB_OK | MB_ICONERROR | MB_SYSTEMMODAL);
		MyOutputDebugStringA("[%s] Line %d: ERROR! Unsupported DPI!\n", __FUNCTION__, __LINE__);
		return(FALSE);
	}

	return(TRUE);
}