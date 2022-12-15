#include "NotePadDlg.h"
#include <gdiplus.h>
using namespace Gdiplus;

int WINAPI _tWinMain(HINSTANCE hInst, HINSTANCE hPrev, LPTSTR lpszCmdLine, int nCmdShow) {
	// Init Common Controls
	INITCOMMONCONTROLSEX icc = { sizeof(INITCOMMONCONTROLSEX) };
	icc.dwICC = ICC_COOL_CLASSES;
	InitCommonControlsEx(&icc);
	
	// Init RichEdit Control
	LoadLibrary(TEXT("Riched20.dll")); // Load RichEdit support

	// Init GDI++ (for loading images)
	GdiplusStartupInput gdiplusStartupInput;
	ULONG_PTR gdiplusToken;
	GdiplusStartup(&gdiplusToken, &gdiplusStartupInput, NULL);	
	
	// Create and show dialog
	NotePadDlg dlg;
	DialogBox(hInst, MAKEINTRESOURCE(IDD_DIALOG1), NULL, NotePadDlg::DlgProc);

	GdiplusShutdown(gdiplusToken);

	return 0;
}