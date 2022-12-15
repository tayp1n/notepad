#include "NotePadDlg.h"

#include <richedit.h>
#include <gdiplus.h>
using namespace Gdiplus;

#include "ImageDataObject.h"

// ��������� �� ��������� NotePadDlg
NotePadDlg* NotePadDlg::ptr = NULL;

// �������� ������������, ����� �� �� ��������� ���������
BOOL askSave(HWND hwnd)
{
	return MessageBox(hwnd, L"������ ��������� ���������?", L"����������", MB_YESNO) == IDYES;
}

// �����������. �������������� ��� ����������.
NotePadDlg::NotePadDlg(void)
{
	ptr = this;
	bShowStatusBar = TRUE;
	bIsOpenF = FALSE;
	
	ZeroMemory(&currFont, sizeof(currFont));
	currColor = RGB(0, 0, 0);
	currBgColor = RGB(255, 255, 255);	

	currPageWidth = 400;
}

// ��� �������� ����
void NotePadDlg::Cls_OnClose(HWND hwnd)
{
	// ���� �������� ��� �������
	if (isModified())
	{
		// �� �������� ������������, ����� �� �� ��������� ���������
		if (askSave(hwnd))
		{
			// ��������� ��������
			saveFileCmd(hwnd, 0);			
		}		
	}
	
	// ������� ������
	EndDialog(hwnd, 0);
}

// ��� ������������� �������
BOOL NotePadDlg::Cls_OnInitDialog(HWND hwnd, HWND hwndFocus, LPARAM lParam)
{
	hDialog = hwnd;
	// ������� ���������� ���������� ����
	hEdit = GetDlgItem(hDialog, IDC_RICHEDIT21);	
	// �������� ������ ���������
	hStatus = CreateStatusWindow(WS_CHILD | WS_VISIBLE | CCS_BOTTOM | SBARS_TOOLTIPS | SBARS_SIZEGRIP, 0, hDialog, WM_USER);
	// �������� ���� �� �������� ����������
	hMenuRU = LoadMenu(GetModuleHandle(NULL), MAKEINTRESOURCE(IDR_MENU1));
	hMenuEN = LoadMenu(GetModuleHandle(NULL), MAKEINTRESOURCE(IDR_MENU2));
	// ����������� ���� � �������� ���� ����������
	SetMenu(hDialog, hMenuRU);	

	// ���������� OLE callback ��� �������� ���������� RichEdit (��. RichCallback.h)
	RichCallback* pCallback = new RichCallback;
	SendMessage(hEdit, EM_SETOLECALLBACK, 0, reinterpret_cast<LPARAM>(pCallback));

	// ���������� ������ ���������
	PARAFORMAT2 pf;
	ZeroMemory(&pf, sizeof(pf));
	pf.cbSize = sizeof(pf);
	pf.dwMask = PFM_OFFSET | PFM_OFFSETINDENT | PFM_RIGHTINDENT;
	pf.dxOffset = -10 * 20;
	pf.dxStartIndent = 10 * 20;
	pf.dxRightIndent = 10 * 20;
	SendMessage(hEdit, EM_SETPARAFORMAT, 0, (LPARAM)&pf);
	// �������� ���� ����������� RichEdit
	SendMessage(hEdit, EM_SETMODIFY, FALSE, 0);	

	// ���������� ������������ ���� �� ��������� ����������� RichEdit
	SendMessage(hEdit, EM_SETEVENTMASK, 0, ENM_CHANGE);

	// �������� ���������� ��������� � StatusBar
	Cls_OnEdit(hEdit);

	return TRUE;
}

// ���������� ��������� WM_COMMAND ����� ������ ��� ������ ������ ����
void NotePadDlg::Cls_OnCommand(HWND hwnd, int id, HWND hwndCtl, UINT codeNotify)
{
	switch (id)
	{
	case ID_NEW: // ���� ����/�������
		newFileCmd(hwnd);
		break;
	case ID_OPEN:  // ���� ����/�������
		openFileCmd(hwnd);		
		break;
	case ID_SAVE: // ���� ����/���������
		saveFileCmd(hwnd, bIsOpenF);
		break;
	case ID_SAVEAS: // ���� ����/��������� ���
		saveFileCmd(hwnd, FALSE);
		break;
	case ID_RU: // ���� ���/����/�������
		SetMenu(hDialog, hMenuRU);
		break;
	case ID_EN: // ���� ���/����/����������
		SetMenu(hDialog, hMenuEN);
		break;
	case ID_UNDO:  // ���� ������/��������
		// ������� ��������� ��������
		SendMessage(hEdit, WM_UNDO, 0, 0);
		break;
	case ID_CUT:  // ���� ������/��������
		// ������ ���������� �������� ������ � ����� ������
		SendMessage(hEdit, WM_CUT, 0, 0);
		break;
	case ID_COPY:  // ���� ������/����������
		// ��������� ���������� �������� ������ � ����� ������
		SendMessage(hEdit, WM_COPY, 0, 0);
		break;
	case ID_PASTE:  // ���� ������/��������
		// ������� ����� � Edit Control �� ������ ������
		SendMessage(hEdit, WM_PASTE, 0, 0);		
		break;
	case ID_DEL:  // ���� ������/�������
		// ������ ���������� �������� ������
		SendMessage(hEdit, WM_CLEAR, 0, 0);
		break;
	case ID_SELECTALL:  // ���� ������/�������� ���
		// ������� ���� ����� � Edit Control
		SendMessage(hEdit, EM_SETSEL, 0, -1);
		break;
	case ID_STATUSBAR:  // ���� ���/������ ���������
		// ���� ���� ����� TRUE, �� ������ ��������� ����������
		if (bShowStatusBar)
		{
			// ������� ���������� �������� ����
			HMENU hMenu = GetMenu(hDialog);
			// ������ ������� � ������ ���� "������ ���������"
			CheckMenuItem(hMenu, ID_STATUSBAR, MF_BYCOMMAND | MF_UNCHECKED);
			// ������ ������ ���������
			ShowWindow(hStatus, SW_HIDE);
		}
		else
		{
			// ������� ���������� �������� ����
			HMENU hMenu = GetMenu(hDialog);
			// ��������� ������� �� ������ ���� "������ ���������"
			CheckMenuItem(hMenu, ID_STATUSBAR, MF_BYCOMMAND | MF_CHECKED);
			// ��������� ������ ���������
			ShowWindow(hStatus, SW_SHOW);
		}
		bShowStatusBar = !bShowStatusBar;
		break;
	case ID_FILE_EXIT:// ���� ����/�����
		Cls_OnClose(hwnd);
		break;
	case ID_FORMAT_FONT: // ���� ������/�����
		// �������� ����� ���������� ���������
		chooseFontCmd(hwnd);
		break;
	case ID_FORMAT_BACKGROUND: // ���� ������/���
		// �������� ��� ���������� ���������
		chooseBgColorCmd(hwnd);
		break;
	case ID_EDIT_INSERT_IMAGE: // ���� ������/�������� ��������
		// �������� ��������
		insertImageCmd(hwnd);
		break;
	case ID_FILE_FILEINFO: // ���� ����/���������� � �����
		// �������� ���������� ���������
		showFileInfoCmd(hwnd);
		break;
	case ID_ALIGNMENT_LEFT: // ���� ������/������������/�� ������ ����
		// �������� ������������ ���������� ���������
		changeParaAlignmentCmd(hwnd, PFA_LEFT);
		break;
	case ID_ALIGNMENT_CENTER: // ���� ������/������������/�� ������
		// �������� ������������ ���������� ���������
		changeParaAlignmentCmd(hwnd, PFA_CENTER);
		break;
	case ID_ALIGNMENT_RIGHT: // ���� ������/������������/�� ������� ����
		// �������� ������������ ���������� ���������
		changeParaAlignmentCmd(hwnd, PFA_RIGHT);
		break;
	case ID_STARTINDENT_INCREASE: // ���� ������/������ �����/���������
		// ��������� ������ ������ ������ ��������� ���������� ���������
		changeParaStartIndentCmd(hwnd, 4 * 20);
		break;
	case ID_STARTINDENT_DECREASE: // ���� ������/������ �����/���������
		// ��������� ������ ������ ������ ��������� ���������� ���������
		changeParaStartIndentCmd(hwnd, -4 * 20);
		break;
	case ID_RIGHTINDENT_INCREASE: // ���� ������/������ ������/���������
		// ��������� ������ ������ ���������� ���������
		changeParaRightIndentCmd(hwnd, 4 * 20);
		break;
	case ID_RIGHTINDENT_DECREASE: // ���� ������/������ ������/���������
		// ��������� ������ ������ ���������� ���������
		changeParaRightIndentCmd(hwnd, -4 * 20);
		break;
	case ID_OFFSET_INCREASE: // ���� ������/�����/���������
		// ��������� ������ ��������� ����� ��������� ������������ ������ ������
		changeParaOffsetCmd(hwnd, 4 * 20);
		break;
	case ID_OFFSET_DECREASE: // ���� ������/�����/���������
		// ��������� ������ ��������� ����� ��������� ������������ ������ ������
		changeParaOffsetCmd(hwnd, -4 * 20);
		break;
	case ID_PAGEWIDTH_INCREASE: // ���� ������/������ ��������/���������
		// ��������� ������ ������� ��������������
		changePageWidth(hwnd, 20);
		break;
	case ID_PAGEWIDTH_DECREASE: // ���� ������/������ ��������/��������
		// ��������� ������ ������� ��������������
		changePageWidth(hwnd, -20);
		break;	 
	}

	if (hwndCtl == hEdit && codeNotify == EN_UPDATE)
	{
		// �������� ���������� ��������� � StatusBar
		Cls_OnEdit(hEdit);
	}
}

// ���������� ��������� WM_SIZE ����� ������ ��� ��������� �������� �������� ����
// ���� ��� ������������/�������������� �������� ����
void NotePadDlg::Cls_OnSize(HWND hwnd, UINT state, int cx, int cy)
{
	RECT rect1, rect2;
	// ������� ���������� ���������� ������� �������� ����
	GetClientRect(hDialog, &rect1);
	// ������� ���������� ������������ ������ ������ ���������
	SendMessage(hStatus, SB_GETRECT, 0, (LPARAM)&rect2);
	// ��������� ����� ������� ���������� ����
	MoveWindow(hEdit, rect1.left, rect1.top, rect1.right, rect1.bottom - (rect2.bottom - rect2.top), 1);
	// ��������� ������ ������ ���������, 
	// ������ ������ ���������� ������� �������� ����
	SendMessage(hStatus, WM_SIZE, 0, 0);

	// �������� ������� �������������� � ������������ � ����� ��������
	changePageWidth(hwnd, 0);	
}

// ���������� WM_INITMENUPOPUP ����� ������ ��������������� 
// ����� ������������ ������������ ����
void NotePadDlg::Cls_OnInitMenuPopup(HWND hwnd, HMENU hMenu, UINT item, BOOL fSystemMenu)
{
	if (item == 0) // �������������� ����� ���� "������"
	{
		// ������� ������� ��������� ������
		DWORD dwPosition = SendMessage(hEdit, EM_GETSEL, 0, 0);
		WORD wBeginPosition = LOWORD(dwPosition);
		WORD wEndPosition = HIWORD(dwPosition);

		if (wEndPosition != wBeginPosition) // ������� �� �����?
		{
			// ���� ������� ���������� �����, 
			// �� ������� ������������ ������ ���� "����������", "��������" � "�������"
			EnableMenuItem(hMenu, ID_COPY, MF_BYCOMMAND | MF_ENABLED);
			EnableMenuItem(hMenu, ID_CUT, MF_BYCOMMAND | MF_ENABLED);
			EnableMenuItem(hMenu, ID_DEL, MF_BYCOMMAND | MF_ENABLED);
		}
		else
		{
			// ���� ����������� ���������� �����, 
			// �� ������� ������������ ������ ���� "����������", "��������" � "�������"
			EnableMenuItem(hMenu, ID_COPY, MF_BYCOMMAND | MF_GRAYED);
			EnableMenuItem(hMenu, ID_CUT, MF_BYCOMMAND | MF_GRAYED);
			EnableMenuItem(hMenu, ID_DEL, MF_BYCOMMAND | MF_GRAYED);
		}

		if (IsClipboardFormatAvailable(CF_TEXT)) // ������� �� ����� � ������ ������?
			// ���� ������� ����� � ������ ������, 
			// �� ������� ����������� ����� ���� "��������"
			EnableMenuItem(hMenu, ID_PASTE, MF_BYCOMMAND | MF_ENABLED);
		else
			// ���� ����������� ����� � ������ ������, 
			// �� ������� ����������� ����� ���� "��������"
			EnableMenuItem(hMenu, ID_PASTE, MF_BYCOMMAND | MF_GRAYED);

		// ���������� �� ����������� ������ ���������� ��������?
		if (SendMessage(hEdit, EM_CANUNDO, 0, 0))
			// ���� ���������� ����������� ������ ���������� ��������,
			// �� ������� ����������� ����� ���� "��������"
			EnableMenuItem(hMenu, ID_UNDO, MF_BYCOMMAND | MF_ENABLED);
		else
			// ���� ����������� ����������� ������ ���������� ��������,
			// �� ������� ����������� ����� ���� "��������"
			EnableMenuItem(hMenu, ID_UNDO, MF_BYCOMMAND | MF_GRAYED);

		// ��������� ����� ������ � Edit Control
		int length = SendMessage(hEdit, WM_GETTEXTLENGTH, 0, 0);
		// ������� �� ���� ����� � Edit Control?
		if (length != wEndPosition - wBeginPosition)
			//���� �� ���� ����� ������� � Edit Control,
			// �� ������� ����������� ����� ���� "�������� ��"
			EnableMenuItem(hMenu, ID_SELECTALL, MF_BYCOMMAND | MF_ENABLED);
		else
			// ���� ������� ���� ����� � Edit Control,
			// �� ������� ����������� ����� ���� "�������� ��"
			EnableMenuItem(hMenu, ID_SELECTALL, MF_BYCOMMAND | MF_GRAYED);
	}
}

void NotePadDlg::Cls_OnMenuSelect(HWND hwnd, HMENU hmenu, int item, HMENU hmenuPopup, UINT flags)
{
	if (flags & MF_POPUP) // ��������, �������� �� ���������� ����� ���� ���������� ����������� �������?
	{
		// ���������� ����� ���� �������� ���������� ����������� �������
		SendMessage(hStatus, SB_SETTEXT, 0, 0); // ������� ����� �� ������ ���������
	}
	else
	{
		// ���������� ����� ���� �������� �������� ������� (����� ���� "�������")
		TCHAR buf[200];
		// ������� ���������� �������� ���������� ����������
		HINSTANCE hInstance = GetModuleHandle(NULL);
		// ������� ������ �� ������� �����, ������������� � �������� ����������
		// ��� ���� ������������� ����������� ������ ������ ������������� �������������� ����������� ������ ����
		LoadString(hInstance, item, buf, 200);
		// ������� � ������ ��������� ����������� �������, ��������������� ����������� ������ ����
		SendMessage(hStatus, SB_SETTEXT, 0, LPARAM(buf));
	}
}

// �������� ���������� ��������� � StatusBar
BOOL NotePadDlg::Cls_OnEdit(HWND hwnd)
{
	// ��������� ����������
	Stat stat = calcStat();

	// ���������� ���������� � string stream
	std::wostringstream woss;
	woss << L"��������: 1 �� 1  |  ";
	woss << L"����� ��������: " << stat.numChars << L"  |  ";
	woss << L"����� ����: " << stat.numWords << L"  |  ";
	woss << L"����� �����: " << stat.numLines;
	
	// ���������� ����� � StatusBar
	SendMessage(hStatus, SB_SETTEXT, 0, LPARAM(woss.str().c_str()));

	return 0;
}

// ������� ��������� ��������� ����������� ����.
BOOL CALLBACK NotePadDlg::DlgProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam)
{	
	switch (message)
	{
		HANDLE_MSG(hwnd, WM_CLOSE, ptr->Cls_OnClose);
		HANDLE_MSG(hwnd, WM_INITDIALOG, ptr->Cls_OnInitDialog);
		HANDLE_MSG(hwnd, WM_COMMAND, ptr->Cls_OnCommand);
		HANDLE_MSG(hwnd, WM_SIZE, ptr->Cls_OnSize);
		HANDLE_MSG(hwnd, WM_INITMENUPOPUP, ptr->Cls_OnInitMenuPopup);
		HANDLE_MSG(hwnd, WM_MENUSELECT, ptr->Cls_OnMenuSelect);		
	}
	return FALSE;
}

// ���� callback ������������ � ���������� EM_STREAMIN ��� ������ �����
DWORD NotePadDlg::readStreamCallback(DWORD_PTR dwCookie, LPBYTE lpBuff, LONG cb, PLONG pcb)
{
	HANDLE hFile = (HANDLE)dwCookie;
	// ������ �������� ����� � �����
	if (ReadFile(hFile, lpBuff, cb, (DWORD*)pcb, NULL))
	{
		return 0;
	}
	return -1;
}

// ���� callback ������������ � ���������� EM_STREAMOUT ��� ������ �����
DWORD NotePadDlg::writeStreamCallback(DWORD_PTR dwCookie, LPBYTE lpBuff, LONG cb, PLONG pcb)
{
	HANDLE hFile = (HANDLE)dwCookie;
	// �������� �������� �� ������ � ����
	if (WriteFile(hFile, lpBuff, cb, (DWORD*)pcb, NULL))
	{
		return 0;
	}
	return -1;
}

// ������� ����� ���� (������ �������� Edit Control).
// ���� ������� �������� ���� ������� ����������, �� �������� �����������.
void NotePadDlg::newFileCmd(HWND hwnd)
{
	// ���� �������� ��� �������
	if (isModified())
	{
		// �� �������� ������������, ����� �� �� ��������� ���������
		if (askSave(hwnd))
		{
			// ��������� ��������
			saveFileCmd(hwnd, 0);
		}		
	}

	// �������� �����
	SendMessage(hEdit, WM_SETTEXT, 0, 0);
	// �������� ���� �����������
	SendMessage(hEdit, EM_SETMODIFY, FALSE, 0);
	// ���������� ����
	bIsOpenF = FALSE;
}

// ������� ����.
void NotePadDlg::openFileCmd(HWND hwnd)
{
	// �������� Edit Control
	newFileCmd(hwnd);

	// ����� ��� ������� ���� � �����
	TCHAR fullPath[MAX_PATH] = { 0 };

	// �������� ��� �����
	if (showOpenFileDlg(hwnd, fullPath))
	{
		// ����� ������������ ��������� EM_STREAMIN ��� ���������� �������� ������ �� ����� 

		// ������� ����
		HANDLE hFile = CreateFile(fullPath, GENERIC_READ, FILE_SHARE_READ, 0, OPEN_EXISTING, FILE_FLAG_SEQUENTIAL_SCAN, NULL);

		if (hFile != INVALID_HANDLE_VALUE)
		{
			// ��������� EDITSTREAM
			EDITSTREAM es = { 0 };
			es.pfnCallback = readStreamCallback;
			es.dwCookie = (DWORD_PTR)hFile;

			// ���� ���������� ����� RTF, ����� ������� ��� RTF, ����� ��� plain text
			DWORD wParam;
			if (isRTF(fullPath))
			{
				wParam = SF_RTF;				
			}
			else
			{
				wParam = SF_TEXT;
			}			

			// ��������� ����� � ��
			if (SendMessage(hEdit, EM_STREAMIN, wParam, (LPARAM)&es) && es.dwError == 0)
			{
				// �������� ���� �����������
				SendMessage(hEdit, EM_SETMODIFY, FALSE, 0);
				// ��������� ��� ����� ��� ������������ ���������� � ������� ������� "Save"
				bIsOpenF = TRUE;
				lstrcpy(FullPath, fullPath);				
			}
			
			// ������� ����
			CloseHandle(hFile);
		}				
	}
}
// ��������� ����. ���� ������ �������� FALSE, �� 
// ������� ��� �����, ����� ������������ ��� �� ���������� FullPath
void NotePadDlg::saveFileCmd(HWND hwnd, BOOL hasName)
{
	if (!hasName)
	{
		TCHAR fullPath[MAX_PATH] = { 0 };
		// �������� ��� �����
		if (!showSaveFileDlg(hwnd, fullPath))
		{
			return;
		}
		// ��������� ��� ����� ��� ������������ ���������� � ������� ������� "Save"
		lstrcpy(FullPath, fullPath);
	}

	// ��������� ������ ���� ����� ������� � ���������� ���������� ��� ��� �� ����������
	if (isModified() || !hasName)
	{
		// ����� ������������ ��������� EM_STREAMOUT ��� ���������� �������� ������ � ����

		// ������� ����
		HANDLE hFile = CreateFile(FullPath, GENERIC_WRITE, FILE_SHARE_READ, 0, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
		
		if (hFile != INVALID_HANDLE_VALUE)
		{
			// ��������� EDITSTREAM
			EDITSTREAM es = { 0 };
			es.pfnCallback = writeStreamCallback;
			es.dwCookie = (DWORD_PTR)hFile;

			// ���� ���������� ����� RTF, ����� ��������� ��� RTF, ����� ��� plain text
			DWORD wParam;
			if (isRTF(FullPath))
			{
				wParam = SF_RTF;
			}
			else
			{
				wParam = SF_TEXT;
			}

			// ��������� �����
			if (SendMessage(hEdit, EM_STREAMOUT, wParam, (LPARAM)&es) && es.dwError == 0)
			{
				// �������� ���� �����������
				SendMessage(hEdit, EM_SETMODIFY, FALSE, 0);
				// ��� ����� ��������� � ���������� FullPath
				bIsOpenF = TRUE;				
			}
			// ������� ����
			CloseHandle(hFile);
		}		
	}	
}

// ������� �������� ������ � ���������� ����� ��� �������� ���������
void NotePadDlg::chooseFontCmd(HWND hwnd)
{
	CHOOSEFONT cf;            // common dialog box structure	

	// ���������������� ��������� CHOOSEFONT
	ZeroMemory(&cf, sizeof(cf));
	cf.lStructSize = sizeof(cf);
	cf.hwndOwner = hwnd;
	cf.lpLogFont = &currFont;
	cf.rgbColors = currColor;
	cf.Flags = CF_SCREENFONTS | CF_EFFECTS | CF_INITTOLOGFONTSTRUCT;

	// ������� �������� ������
	if (ChooseFont(&cf) == TRUE)
	{
		// ���������� ����� ��� �������� ���������

		// ���������������� ��������� CHARFORMAT2
		CHARFORMAT2 chf;
		ZeroMemory(&chf, sizeof(chf));
		chf.cbSize = sizeof(chf);

		// �������� ������� ������ ����������� ���������
		SendMessage(hEdit, EM_GETCHARFORMAT, SCF_SELECTION, (LPARAM)&chf);

		// �������������� ��������� CHARFORMAT2 � ������������ � �������, ����������� �� ��������� CHOOSEFONT

		chf.dwMask = CFM_ALL | CFM_WEIGHT;
		chf.bPitchAndFamily = currFont.lfPitchAndFamily;
		lstrcpy(chf.szFaceName, currFont.lfFaceName);
		chf.bCharSet = currFont.lfCharSet;
		chf.wWeight = (WORD)currFont.lfWeight;

		chf.dwEffects = 0;		
		if (currFont.lfWeight == FW_BOLD)
		{
			chf.dwEffects |= CFE_BOLD;
		}
		if (currFont.lfItalic)
		{
			chf.dwEffects |= CFE_ITALIC;
		}
		if (currFont.lfStrikeOut)
		{
			chf.dwEffects |= CFE_STRIKEOUT;
		}
		if (currFont.lfUnderline)
		{
			chf.dwEffects |= CFE_UNDERLINE;
		}

		chf.yHeight = cf.iPointSize * 2;
		chf.crTextColor = cf.rgbColors;

		// ���������� ������ ����������� ���������
		SendMessage(hEdit, EM_SETCHARFORMAT, SCF_SELECTION, (LPARAM)&chf);

		// ��������� ������� ����
		currColor = cf.rgbColors;
	}
}

// ������� ���� ����
void NotePadDlg::chooseBgColorCmd(HWND hwnd)
{
	CHOOSECOLOR cc;                 // common dialog box structure	
	
	// ���������������� ��������� CHOOSECOLOR 
	ZeroMemory(&cc, sizeof(cc));
	cc.lStructSize = sizeof(cc);
	cc.hwndOwner = hwnd;	
	cc.lpCustColors = (LPDWORD)acrCustClr;
	cc.rgbResult = currBgColor;
	cc.Flags = CC_FULLOPEN | CC_RGBINIT;

	// �������� ������ ������ �����
	if (ChooseColor(&cc) == TRUE)
	{
		// ���������������� ��������� CHARFORMAT2 
		CHARFORMAT2 chf;
		ZeroMemory(&chf, sizeof(chf));
		chf.cbSize = sizeof(chf);

		chf.dwMask = CFM_BACKCOLOR; // ������ ������ ���� ����
		chf.crBackColor = cc.rgbResult;

		// ���������� ������ ����������� ���������
		SendMessage(hEdit, EM_SETCHARFORMAT, SCF_SELECTION, (LPARAM)&chf);

		// ��������� ������� ���� ����
		currBgColor = cc.rgbResult;
	}
}

// �������� �������� � ��������
void NotePadDlg::insertImageCmd(HWND hwnd)
{
	// ������� ���� � ��������� � �������� HBITMAP
	HBITMAP hBitmap = chooseImage(hwnd);

	if (hBitmap)
	{
		// ������������ OLE ��� ������� �������� � ��������
		
		
		// �������� ��������� �� OLE-���������
		LPRICHEDITOLE pRichEditOle;
		SendMessage(hEdit, EM_GETOLEINTERFACE, 0, (LPARAM)&pRichEditOle);

		// �������� �������� � �������� (��. ImageDataObject.h)
		ImageDataObject::InsertBitmap(pRichEditOle, hBitmap);

		DeleteObject(hBitmap);

		pRichEditOle->Release();
	}
}

// �������� ���������� ��������� � MessageBox
void NotePadDlg::showFileInfoCmd(HWND hwnd)
{
	// ��������� ����������
	Stat stat = calcStat();

	// ���������� ���������� � string stream
	std::wostringstream woss;
	woss << "Length:      " << stat.totalChars << "\r";
	woss << "Characters:  " << stat.numChars  << "\r";
	woss << "Words:       " << stat.numWords<< "\r";
	woss << "Lines:       " << stat.numLines << "\r";
	woss << "Pages:       " << stat.numPages << "\r";

	MessageBox(hwnd, woss.str().c_str(), L"���������� � �����", MB_OK);	
}

// �������� ������������ ���������� ���������
void NotePadDlg::changeParaAlignmentCmd(HWND hwnd, WORD wAlignment)
{
	// ���������������� ��������� PARAFORMAT2
	PARAFORMAT2 pf;
	ZeroMemory(&pf, sizeof(pf));
	pf.cbSize = sizeof(pf);
	pf.dwMask = PFM_ALIGNMENT;
	pf.wAlignment = wAlignment;

	// ������ ������������ ���������� ���������
	SendMessage(hEdit, EM_SETPARAFORMAT, 0, (LPARAM)&pf);
}

// �������� ������ ������ ������ ��������� ���������� ���������
void NotePadDlg::changeParaStartIndentCmd(HWND hwnd, LONG twips)
{
	// ���������������� ��������� PARAFORMAT2
	PARAFORMAT2 pf;
	ZeroMemory(&pf, sizeof(pf));
	pf.cbSize = sizeof(pf);

	// �������� ������� ������ ���������
	SendMessage(hEdit, EM_GETPARAFORMAT, 0, (LPARAM)&pf);

	// �������� ������ ������ ������ �� ��������� �������� � ���������� twips
	pf.dwMask = PFM_STARTINDENT;
	pf.dxStartIndent += twips;
	SendMessage(hEdit, EM_SETPARAFORMAT, 0, (LPARAM)&pf);
}

// �������� ������ ������ ���������� ���������
void NotePadDlg::changeParaRightIndentCmd(HWND hwnd, LONG twips)
{
	// ���������������� ��������� PARAFORMAT2
	PARAFORMAT2 pf;
	ZeroMemory(&pf, sizeof(pf));
	pf.cbSize = sizeof(pf);

	// �������� ������� ������ ���������
	SendMessage(hEdit, EM_GETPARAFORMAT, 0, (LPARAM)&pf);

	// �������� ������ ������  �� ��������� �������� � ���������� twips
	pf.dwMask = PFM_RIGHTINDENT;
	pf.dxRightIndent += twips;
	SendMessage(hEdit, EM_SETPARAFORMAT, 0, (LPARAM)&pf);
}

// �������� ������ ��������� ����� ��������� ������������ ������ ������
void NotePadDlg::changeParaOffsetCmd(HWND hwnd, LONG twips)
{
	// ���������������� ��������� PARAFORMAT2
	PARAFORMAT2 pf;
	ZeroMemory(&pf, sizeof(pf));
	pf.cbSize = sizeof(pf);

	// �������� ������� ������ ���������
	SendMessage(hEdit, EM_GETPARAFORMAT, 0, (LPARAM)&pf);

	// �������� ������ �� ��������� �������� � ���������� twips
	pf.dwMask = PFM_OFFSET;
	pf.dxOffset += twips;
	SendMessage(hEdit, EM_SETPARAFORMAT, 0, (LPARAM)&pf);
}

// �������� ������ ������� �������������� �� ��������� �������� � ���������� pixels
void NotePadDlg::changePageWidth(HWND hwnd, LONG pixels)
{
	currPageWidth += pixels;

	// �������� ������ ���������� ����� RichEdit
	RECT rc;
	GetClientRect(hEdit, &rc);

	// ������ ������� ������� ��������������
	rc.left += 20;
	rc.right = rc.left + currPageWidth;
	rc.top += 10;
	rc.bottom += 0;

	SendMessage(hEdit, EM_SETRECT, 0, (LPARAM)&rc);
}

// �������� ������ �������� �����. ��� ����� ������������ � ���������� buffer
BOOL NotePadDlg::showOpenFileDlg(HWND hwnd, TCHAR buffer[MAX_PATH])
{
	// �������� �����
	buffer[0] = L'\0';
	
	// ���������������� ��������� OPENFILENAME
	OPENFILENAME open = { sizeof(OPENFILENAME) };
	open.hwndOwner = hwnd;
	open.lpstrFilter = TEXT("Text Files(*.txt)\0*.txt\0RTF Files(*.rtf)\0*.rtf\0All Files(*.*)\0*.*\0");
	open.lpstrFile = buffer;
	open.nMaxFile = MAX_PATH;
	open.lpstrInitialDir = TEXT("C:\\");
	open.Flags = OFN_FILEMUSTEXIST | OFN_HIDEREADONLY;
	// �������� ������ ������� ��� ������ �����
	return GetOpenFileName(&open);
}

// �������� ������ ���������� �����. ��� ����� ������������ � ���������� buffer
BOOL NotePadDlg::showSaveFileDlg(HWND hwnd, TCHAR buffer[MAX_PATH])
{
	// �������� �����
	buffer[0] = L'\0';
	// ���������������� ��������� OPENFILENAME
	OPENFILENAME open = { sizeof(OPENFILENAME) };
	open.hwndOwner = hwnd;
	open.lpstrFilter = TEXT("Text Files(*.txt)\0*.txt\0RTF Files(*.rtf)\0*.rtf\0All Files(*.*)\0*.*\0");
	open.lpstrFile = buffer;
	open.nMaxFile = MAX_PATH;
	open.Flags = OFN_OVERWRITEPROMPT | OFN_PATHMUSTEXIST;	
	
	// �������� ������ ������� ��� ������ �����
	if (GetSaveFileName(&open))
	{
		// ���������, ���� ������������ �� ���� ����������
		if (open.nFileExtension == 0)
		{
			if (open.nFilterIndex == 1)
			{
				// �������� .txt
				lstrcat(buffer, L".txt");
			}
			else if (open.nFilterIndex == 2)
			{
				// �������� .rtf
				lstrcat(buffer, L".rtf");
			}
			else
			{
				// do nothing
			}
		}
		return TRUE;
	}
	return FALSE;
}

// ������� ����������� ��� �������, ��������� � ������� HBITMAP
HBITMAP NotePadDlg::chooseImage(HWND hwnd)
{
	// �� ����� �� ���������� �������� ����� GetOpenFileName() ��������
	// � ��������� ����� ������ � RichEdit. ������� � �������� ������� ���������
	// � ����� �������������� ����� GetOpenFileName().
	CHARRANGE cr;
	SendMessage(hEdit, EM_EXGETSEL, 0, (LPARAM)&cr);

	
	// ����� ��� ����� �����
	TCHAR lpstrFile[MAX_PATH] = {0};
	// ���������������� ��������� OPENFILENAME
	OPENFILENAME open = { sizeof(OPENFILENAME) };
	open.hwndOwner = hwnd;
	open.lpstrFilter = TEXT("Image Files (*.jpg, *.png, *.bmp)\0*.jpg;*.png;*.bmp\0");
	open.lpstrFile = lpstrFile;
	open.nMaxFile = MAX_PATH;
	open.lpstrInitialDir = TEXT("C:\\");
	open.Flags = OFN_FILEMUSTEXIST | OFN_HIDEREADONLY;
	// �������� ��� ����� ��������
	BOOL result = GetOpenFileName(&open);
	
	// ������������ ��������� � RichEdit
	SendMessage(hEdit, EM_EXSETSEL, 0, (LPARAM)&cr);

	if (!result)
	{		
		return NULL;
	}	
	
	HBITMAP hBitmap = NULL;

	// ��������� �������� ��������� GDI+
	Bitmap * bitmap = Bitmap::FromFile(lpstrFile);
	// �������� ���������� HBITMAP
	bitmap->GetHBITMAP(Color(255, 255, 255), &hBitmap);
	delete bitmap;

	return hBitmap;
}

// ��������� ��������� �������������� ���������� � ���������.
Stat NotePadDlg::calcStat()
{
	// �������� ���� ����� � �����
	int len = SendMessage(hEdit, WM_GETTEXTLENGTH, 0, 0);
	TCHAR* buffer = new TCHAR[len + 1];
	SendMessage(hEdit, WM_GETTEXT, len + 1, (LPARAM)buffer);

	Stat stat;
	stat.totalChars = len;

	std::wstring word; // ������� �����
	TCHAR ch; // ������� ������

	// ��� ������� �������:
	for (int i = 0; i < len; ++i)
	{
		ch = buffer[i];

		if (ch != L' ' && ch != L'\r' && ch != L'\n')
		{
			// ���� ��� �� ���������� ������, �� ��������� ������� ��������
			stat.numChars += 1;
		}

		if (IsCharAlphaNumeric(ch) || ch == L'\'' || ch == L'-')
		{
			// ���� ��� ���������-�������� ������, ��� ��������, ��� �����,
			// �� �������� ������ � �������� �����
			word += ch;
		}
		else
		{
			// ���� ��� ������ ����� ������, �� ��������� ������� �����
			if (ch == L'\r')
			{
				stat.numLines += 1;
			}

			if (!word.empty())
			{
				// ���� ������� ����� �� ������, �� �������� ���
				word.clear();
				// � ��������� ������� ����
				stat.numWords += 1;
			}
		}
	}

	// ����� �����:

	if (!word.empty())
	{
		// ���� ������� ����� �� ������, �� ��������� ������� ����
		stat.numWords += 1;
	}

	// ���� ��������� ������ ��� �� ������ ����� ������, �� ��������� ������� �����
	if (len > 0 && ch != L'\r' && ch != L'\n')
	{
		stat.numLines += 1;
	}

	delete[] buffer;
	
	return stat;
}

// ���������� TRUE ���� �������� ��� �������������
BOOL NotePadDlg::isModified()
{
	return SendMessage(hEdit, EM_GETMODIFY, 0, 0);
}

// ���������� TRUE ���� ���������� ����� RTF
BOOL NotePadDlg::isRTF(LPCTSTR lpszFileName)
{
	int len = lstrlen(lpszFileName);
	return len >= 4 && lstrcmpi(L".rtf", lpszFileName + len - 4) == 0;
}
