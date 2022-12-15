#pragma once
#include "Header.h"
#include "RichCallback.h"

// Стркутура для хранения статистики документа
struct Stat
{
	int totalChars = 0; // общее кол-во символов
	int numChars = 0; // кол-во печатных символов
	int numWords = 0; // кол-во слов
	int numLines = 0; // кол-во строк
	int numPages = 0; // кол-во страниц (всегда 1 в текущей реализации)
};

// Класс, который реализует главное окно на основе Dialog
class NotePadDlg
{
public:
	// Конструктор
	NotePadDlg(void);
public:
	// Callbacks
	static BOOL CALLBACK DlgProc(HWND hWnd, UINT mes, WPARAM wp, LPARAM lp);
	static DWORD CALLBACK readStreamCallback(DWORD_PTR dwCookie, LPBYTE lpBuff, LONG cb, PLONG pcb);
	static DWORD CALLBACK writeStreamCallback(DWORD_PTR dwCookie, LPBYTE lpBuff, LONG cb, PLONG pcb);
private:
	// Комманды
	BOOL Cls_OnInitDialog(HWND hwnd, HWND hwndFocus, LPARAM lParam);
	void Cls_OnCommand(HWND hwnd, int id, HWND hwndCtl, UINT codeNotify);
	void Cls_OnClose(HWND hwnd);
	void Cls_OnSize(HWND hwnd, UINT state, int cx, int cy);
	void Cls_OnInitMenuPopup(HWND hwnd, HMENU hMenu, UINT item, BOOL fSystemMenu);
	void Cls_OnMenuSelect(HWND hwnd, HMENU hmenu, int item, HMENU hmenuPopup, UINT flags);
	BOOL Cls_OnEdit(HWND hwnd);

	void newFileCmd(HWND hwnd);
	void openFileCmd(HWND hwnd);
	void saveFileCmd(HWND hwnd, BOOL hasName);
	void chooseFontCmd(HWND hwnd);
	void chooseBgColorCmd(HWND hwnd);
	void insertImageCmd(HWND hwnd);
	void showFileInfoCmd(HWND hwnd);
	void changeParaAlignmentCmd(HWND hwnd, WORD wAlignment);
	void changeParaStartIndentCmd(HWND hwnd, LONG twips);
	void changeParaRightIndentCmd(HWND hwnd, LONG twips);
	void changeParaOffsetCmd(HWND hwnd, LONG twips);
	void changePageWidth(HWND hwnd, LONG pixels);

	// Диалоги выбора файла и шрифта
	BOOL showOpenFileDlg(HWND hwnd, TCHAR buffer[MAX_PATH]);
	BOOL showSaveFileDlg(HWND hwnd, TCHAR buffer[MAX_PATH]);
	HBITMAP chooseImage(HWND hwnd);

	// Вычисление статистики документа
	Stat calcStat();

	// Вспомогательные функции
	BOOL isModified();
	BOOL isRTF(LPCTSTR lpszFileName);
private:	
	static NotePadDlg* ptr;

	LOGFONT currFont;        // current font
	DWORD currColor;  // current text color
	DWORD currBgColor;  // current background color
	COLORREF acrCustClr[16]; // array of custom colors	
	LONG currPageWidth; // current page width

	TCHAR FullPath[MAX_PATH]; // Полный путь к текущему документу
	BOOL bIsOpenF; // TRUE, если путь к текущему документу задан
	HWND hDialog, hStatus, hEdit; // Дескрипторы окон
	HMENU hMenuRU, hMenuEN; // Дескрипторы меню
	BOOL bShowStatusBar;	
};
