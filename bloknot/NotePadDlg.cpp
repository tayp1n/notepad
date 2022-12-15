#include "NotePadDlg.h"

#include <richedit.h>
#include <gdiplus.h>
using namespace Gdiplus;

#include "ImageDataObject.h"

// Указатель на экземпляр NotePadDlg
NotePadDlg* NotePadDlg::ptr = NULL;

// Спросить пользователя, хочет ли он сохранить изменения
BOOL askSave(HWND hwnd)
{
	return MessageBox(hwnd, L"Хотите сохранить изменения?", L"Секундочку", MB_YESNO) == IDYES;
}

// Конструктор. Инициализирует все переменные.
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

// При закрытии окна
void NotePadDlg::Cls_OnClose(HWND hwnd)
{
	// Если документ был изменен
	if (isModified())
	{
		// то спросить пользователя, хочет ли он сохранить изменения
		if (askSave(hwnd))
		{
			// Сохранить документ
			saveFileCmd(hwnd, 0);			
		}		
	}
	
	// Закрыть диалог
	EndDialog(hwnd, 0);
}

// При инициализации диалога
BOOL NotePadDlg::Cls_OnInitDialog(HWND hwnd, HWND hwndFocus, LPARAM lParam)
{
	hDialog = hwnd;
	// Получим дескриптор текстового поля
	hEdit = GetDlgItem(hDialog, IDC_RICHEDIT21);	
	// Создадим строку состояния
	hStatus = CreateStatusWindow(WS_CHILD | WS_VISIBLE | CCS_BOTTOM | SBARS_TOOLTIPS | SBARS_SIZEGRIP, 0, hDialog, WM_USER);
	// Загрузим меню из ресурсов приложения
	hMenuRU = LoadMenu(GetModuleHandle(NULL), MAKEINTRESOURCE(IDR_MENU1));
	hMenuEN = LoadMenu(GetModuleHandle(NULL), MAKEINTRESOURCE(IDR_MENU2));
	// Присоединим меню к главному окну приложения
	SetMenu(hDialog, hMenuRU);	

	// Установить OLE callback для элемента управления RichEdit (см. RichCallback.h)
	RichCallback* pCallback = new RichCallback;
	SendMessage(hEdit, EM_SETOLECALLBACK, 0, reinterpret_cast<LPARAM>(pCallback));

	// Установить формат параграфа
	PARAFORMAT2 pf;
	ZeroMemory(&pf, sizeof(pf));
	pf.cbSize = sizeof(pf);
	pf.dwMask = PFM_OFFSET | PFM_OFFSETINDENT | PFM_RIGHTINDENT;
	pf.dxOffset = -10 * 20;
	pf.dxStartIndent = 10 * 20;
	pf.dxRightIndent = 10 * 20;
	SendMessage(hEdit, EM_SETPARAFORMAT, 0, (LPARAM)&pf);
	// Очистить флаг модификации RichEdit
	SendMessage(hEdit, EM_SETMODIFY, FALSE, 0);	

	// Уведомлять родительское окно об изменении содержимого RichEdit
	SendMessage(hEdit, EM_SETEVENTMASK, 0, ENM_CHANGE);

	// Обновить статистику документа в StatusBar
	Cls_OnEdit(hEdit);

	return TRUE;
}

// Обработчик сообщения WM_COMMAND будет вызван при выборе пункта меню
void NotePadDlg::Cls_OnCommand(HWND hwnd, int id, HWND hwndCtl, UINT codeNotify)
{
	switch (id)
	{
	case ID_NEW: // Меню Файл/Создать
		newFileCmd(hwnd);
		break;
	case ID_OPEN:  // Меню Файл/Открыть
		openFileCmd(hwnd);		
		break;
	case ID_SAVE: // Меню Файл/Сохранить
		saveFileCmd(hwnd, bIsOpenF);
		break;
	case ID_SAVEAS: // Меню Файл/Сохранить как
		saveFileCmd(hwnd, FALSE);
		break;
	case ID_RU: // Меню Вид/Язык/Русский
		SetMenu(hDialog, hMenuRU);
		break;
	case ID_EN: // Меню Вид/Язык/Английский
		SetMenu(hDialog, hMenuEN);
		break;
	case ID_UNDO:  // Меню Правка/Отменить
		// Отменим последнее действие
		SendMessage(hEdit, WM_UNDO, 0, 0);
		break;
	case ID_CUT:  // Меню Правка/Вырезать
		// Удалим выделенный фрагмент текста в буфер обмена
		SendMessage(hEdit, WM_CUT, 0, 0);
		break;
	case ID_COPY:  // Меню Правка/Копировать
		// Скопируем выделенный фрагмент текста в буфер обмена
		SendMessage(hEdit, WM_COPY, 0, 0);
		break;
	case ID_PASTE:  // Меню Правка/Вставить
		// Вставим текст в Edit Control из буфера обмена
		SendMessage(hEdit, WM_PASTE, 0, 0);		
		break;
	case ID_DEL:  // Меню Правка/Удалить
		// Удалим выделенный фрагмент текста
		SendMessage(hEdit, WM_CLEAR, 0, 0);
		break;
	case ID_SELECTALL:  // Меню Правка/Выделить все
		// Выделим весь текст в Edit Control
		SendMessage(hEdit, EM_SETSEL, 0, -1);
		break;
	case ID_STATUSBAR:  // Меню Вид/Строка состояния
		// Если флаг равен TRUE, то строка состояния отображена
		if (bShowStatusBar)
		{
			// Получим дескриптор главного меню
			HMENU hMenu = GetMenu(hDialog);
			// Снимем отметку с пункта меню "Строка состояния"
			CheckMenuItem(hMenu, ID_STATUSBAR, MF_BYCOMMAND | MF_UNCHECKED);
			// Скроем строку состояния
			ShowWindow(hStatus, SW_HIDE);
		}
		else
		{
			// Получим дескриптор главного меню
			HMENU hMenu = GetMenu(hDialog);
			// Установим отметку на пункте меню "Строка состояния"
			CheckMenuItem(hMenu, ID_STATUSBAR, MF_BYCOMMAND | MF_CHECKED);
			// Отобразим строку состояния
			ShowWindow(hStatus, SW_SHOW);
		}
		bShowStatusBar = !bShowStatusBar;
		break;
	case ID_FILE_EXIT:// Меню Файл/Выход
		Cls_OnClose(hwnd);
		break;
	case ID_FORMAT_FONT: // Меню Формат/Шрифт
		// Изменить шрифт выбранного фрагмента
		chooseFontCmd(hwnd);
		break;
	case ID_FORMAT_BACKGROUND: // Меню Формат/Фон
		// Изменить фон выбранного фрагмента
		chooseBgColorCmd(hwnd);
		break;
	case ID_EDIT_INSERT_IMAGE: // Меню Правка/Вставить картинку
		// Вставить картинку
		insertImageCmd(hwnd);
		break;
	case ID_FILE_FILEINFO: // Меню Файл/Информация о файле
		// Показать статистику документа
		showFileInfoCmd(hwnd);
		break;
	case ID_ALIGNMENT_LEFT: // Меню Формат/Выравнивание/По левому краю
		// Изменить выравнивание выбранного фрагмента
		changeParaAlignmentCmd(hwnd, PFA_LEFT);
		break;
	case ID_ALIGNMENT_CENTER: // Меню Формат/Выравнивание/По центру
		// Изменить выравнивание выбранного фрагмента
		changeParaAlignmentCmd(hwnd, PFA_CENTER);
		break;
	case ID_ALIGNMENT_RIGHT: // Меню Формат/Выравнивание/По правому краю
		// Изменить выравнивание выбранного фрагмента
		changeParaAlignmentCmd(hwnd, PFA_RIGHT);
		break;
	case ID_STARTINDENT_INCREASE: // Меню Формат/Отступ слева/Увеличить
		// Увеличить отступ первой строки параграфа выбранного фрагмента
		changeParaStartIndentCmd(hwnd, 4 * 20);
		break;
	case ID_STARTINDENT_DECREASE: // Меню Формат/Отступ слева/Уменьшить
		// Уменьшить отступ первой строки параграфа выбранного фрагмента
		changeParaStartIndentCmd(hwnd, -4 * 20);
		break;
	case ID_RIGHTINDENT_INCREASE: // Меню Формат/Отступ справа/Увеличить
		// Увеличить отступ справа выбранного фрагмента
		changeParaRightIndentCmd(hwnd, 4 * 20);
		break;
	case ID_RIGHTINDENT_DECREASE: // Меню Формат/Отступ справа/Уменьшить
		// Уменьшить отступ справа выбранного фрагмента
		changeParaRightIndentCmd(hwnd, -4 * 20);
		break;
	case ID_OFFSET_INCREASE: // Меню Формат/Абзац/Увеличить
		// Увеличить отступ остальных строк параграфа относительно первой строки
		changeParaOffsetCmd(hwnd, 4 * 20);
		break;
	case ID_OFFSET_DECREASE: // Меню Формат/Абзац/Уменьшить
		// Уменьшить отступ остальных строк параграфа относительно первой строки
		changeParaOffsetCmd(hwnd, -4 * 20);
		break;
	case ID_PAGEWIDTH_INCREASE: // Меню Формат/Ширина страницы/Увеличить
		// Увеличить ширину области редактирования
		changePageWidth(hwnd, 20);
		break;
	case ID_PAGEWIDTH_DECREASE: // Меню Формат/Ширина страницы/Умеьшить
		// Уменьшить ширину области редактирования
		changePageWidth(hwnd, -20);
		break;	 
	}

	if (hwndCtl == hEdit && codeNotify == EN_UPDATE)
	{
		// Обновить статистику документа в StatusBar
		Cls_OnEdit(hEdit);
	}
}

// Обработчик сообщения WM_SIZE будет вызван при изменении размеров главного окна
// либо при сворачивании/восстановлении главного окна
void NotePadDlg::Cls_OnSize(HWND hwnd, UINT state, int cx, int cy)
{
	RECT rect1, rect2;
	// Получим координаты клиентской области главного окна
	GetClientRect(hDialog, &rect1);
	// Получим координаты единственной секции строки состояния
	SendMessage(hStatus, SB_GETRECT, 0, (LPARAM)&rect2);
	// Установим новые размеры текстового поля
	MoveWindow(hEdit, rect1.left, rect1.top, rect1.right, rect1.bottom - (rect2.bottom - rect2.top), 1);
	// Установим размер строки состояния, 
	// равный ширине клиентской области главного окна
	SendMessage(hStatus, WM_SIZE, 0, 0);

	// Изменить область редактирования в соответствии с новым размером
	changePageWidth(hwnd, 0);	
}

// Обработчик WM_INITMENUPOPUP будет вызван непосредственно 
// перед активизацией всплывающего меню
void NotePadDlg::Cls_OnInitMenuPopup(HWND hwnd, HMENU hMenu, UINT item, BOOL fSystemMenu)
{
	if (item == 0) // Активизируется пункт меню "Правка"
	{
		// Получим границы выделения текста
		DWORD dwPosition = SendMessage(hEdit, EM_GETSEL, 0, 0);
		WORD wBeginPosition = LOWORD(dwPosition);
		WORD wEndPosition = HIWORD(dwPosition);

		if (wEndPosition != wBeginPosition) // Выделен ли текст?
		{
			// Если имеется выделенный текст, 
			// то сделаем разрешёнными пункты меню "Копировать", "Вырезать" и "Удалить"
			EnableMenuItem(hMenu, ID_COPY, MF_BYCOMMAND | MF_ENABLED);
			EnableMenuItem(hMenu, ID_CUT, MF_BYCOMMAND | MF_ENABLED);
			EnableMenuItem(hMenu, ID_DEL, MF_BYCOMMAND | MF_ENABLED);
		}
		else
		{
			// Если отсутствует выделенный текст, 
			// то сделаем недоступными пункты меню "Копировать", "Вырезать" и "Удалить"
			EnableMenuItem(hMenu, ID_COPY, MF_BYCOMMAND | MF_GRAYED);
			EnableMenuItem(hMenu, ID_CUT, MF_BYCOMMAND | MF_GRAYED);
			EnableMenuItem(hMenu, ID_DEL, MF_BYCOMMAND | MF_GRAYED);
		}

		if (IsClipboardFormatAvailable(CF_TEXT)) // Имеется ли текст в буфере обмена?
			// Если имеется текст в буфере обмена, 
			// то сделаем разрешённым пункт меню "Вставить"
			EnableMenuItem(hMenu, ID_PASTE, MF_BYCOMMAND | MF_ENABLED);
		else
			// Если отсутствует текст в буфере обмена, 
			// то сделаем недоступным пункт меню "Вставить"
			EnableMenuItem(hMenu, ID_PASTE, MF_BYCOMMAND | MF_GRAYED);

		// Существует ли возможность отмены последнего действия?
		if (SendMessage(hEdit, EM_CANUNDO, 0, 0))
			// Если существует возможность отмены последнего действия,
			// то сделаем разрешённым пункт меню "Отменить"
			EnableMenuItem(hMenu, ID_UNDO, MF_BYCOMMAND | MF_ENABLED);
		else
			// Если отсутствует возможность отмены последнего действия,
			// то сделаем недоступным пункт меню "Отменить"
			EnableMenuItem(hMenu, ID_UNDO, MF_BYCOMMAND | MF_GRAYED);

		// Определим длину текста в Edit Control
		int length = SendMessage(hEdit, WM_GETTEXTLENGTH, 0, 0);
		// Выделен ли весь текст в Edit Control?
		if (length != wEndPosition - wBeginPosition)
			//Если не весь текст выделен в Edit Control,
			// то сделаем разрешённым пункт меню "Выделить всё"
			EnableMenuItem(hMenu, ID_SELECTALL, MF_BYCOMMAND | MF_ENABLED);
		else
			// Если выделен весь текст в Edit Control,
			// то сделаем недоступным пункт меню "Выделить всё"
			EnableMenuItem(hMenu, ID_SELECTALL, MF_BYCOMMAND | MF_GRAYED);
	}
}

void NotePadDlg::Cls_OnMenuSelect(HWND hwnd, HMENU hmenu, int item, HMENU hmenuPopup, UINT flags)
{
	if (flags & MF_POPUP) // Проверим, является ли выделенный пункт меню заголовком выпадающего подменю?
	{
		// Выделенный пункт меню является заголовком выпадающего подменю
		SendMessage(hStatus, SB_SETTEXT, 0, 0); // Убираем текст со строки состояния
	}
	else
	{
		// Выделенный пункт меню является конечным пунктом (пункт меню "команда")
		TCHAR buf[200];
		// Получим дескриптор текущего экземпляра приложения
		HINSTANCE hInstance = GetModuleHandle(NULL);
		// Зарузим строку из таблицы строк, расположенной в ресурсах приложения
		// При этом идентификатор загружаемой строки строго соответствует идентификатору выделенного пункта меню
		LoadString(hInstance, item, buf, 200);
		// Выводим в строку состояния контекстную справку, соответствующую выделенному пункту меню
		SendMessage(hStatus, SB_SETTEXT, 0, LPARAM(buf));
	}
}

// Обновить статистику документа в StatusBar
BOOL NotePadDlg::Cls_OnEdit(HWND hwnd)
{
	// Расчитать статистику
	Stat stat = calcStat();

	// Напечатать информацию в string stream
	std::wostringstream woss;
	woss << L"Страница: 1 из 1  |  ";
	woss << L"Число символов: " << stat.numChars << L"  |  ";
	woss << L"Число слов: " << stat.numWords << L"  |  ";
	woss << L"Число строк: " << stat.numLines;
	
	// Отобразить текст в StatusBar
	SendMessage(hStatus, SB_SETTEXT, 0, LPARAM(woss.str().c_str()));

	return 0;
}

// Функция обработки сообщений диалогового окна.
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

// Этот callback используется с сообщением EM_STREAMIN для чтения файла
DWORD NotePadDlg::readStreamCallback(DWORD_PTR dwCookie, LPBYTE lpBuff, LONG cb, PLONG pcb)
{
	HANDLE hFile = (HANDLE)dwCookie;
	// Читать фрагмент файла в буфер
	if (ReadFile(hFile, lpBuff, cb, (DWORD*)pcb, NULL))
	{
		return 0;
	}
	return -1;
}

// Этот callback используется с сообщением EM_STREAMOUT для записи файла
DWORD NotePadDlg::writeStreamCallback(DWORD_PTR dwCookie, LPBYTE lpBuff, LONG cb, PLONG pcb)
{
	HANDLE hFile = (HANDLE)dwCookie;
	// Записать фрагмент из буфера в файл
	if (WriteFile(hFile, lpBuff, cb, (DWORD*)pcb, NULL))
	{
		return 0;
	}
	return -1;
}

// Создать новый файл (просто очистить Edit Control).
// Если текущий открытый файл требует сохранения, то спросить пользоателя.
void NotePadDlg::newFileCmd(HWND hwnd)
{
	// Если документ был изменен
	if (isModified())
	{
		// то спросить пользователя, хочет ли он сохранить изменения
		if (askSave(hwnd))
		{
			// Сохранить документ
			saveFileCmd(hwnd, 0);
		}		
	}

	// Очистить текст
	SendMessage(hEdit, WM_SETTEXT, 0, 0);
	// Сбросить флаг модификации
	SendMessage(hEdit, EM_SETMODIFY, FALSE, 0);
	// Безымянный файл
	bIsOpenF = FALSE;
}

// Открыть файл.
void NotePadDlg::openFileCmd(HWND hwnd)
{
	// очистить Edit Control
	newFileCmd(hwnd);

	// буфер для полного пути к файлу
	TCHAR fullPath[MAX_PATH] = { 0 };

	// Получить имя файла
	if (showOpenFileDlg(hwnd, fullPath))
	{
		// Будем использовать сообщение EM_STREAMIN для правильной загрузки текста из файла 

		// Открыть файл
		HANDLE hFile = CreateFile(fullPath, GENERIC_READ, FILE_SHARE_READ, 0, OPEN_EXISTING, FILE_FLAG_SEQUENTIAL_SCAN, NULL);

		if (hFile != INVALID_HANDLE_VALUE)
		{
			// Настроить EDITSTREAM
			EDITSTREAM es = { 0 };
			es.pfnCallback = readStreamCallback;
			es.dwCookie = (DWORD_PTR)hFile;

			// Если расширкние файла RTF, тогда открыть как RTF, иначе как plain text
			DWORD wParam;
			if (isRTF(fullPath))
			{
				wParam = SF_RTF;				
			}
			else
			{
				wParam = SF_TEXT;
			}			

			// Загрузить текст в ЭУ
			if (SendMessage(hEdit, EM_STREAMIN, wParam, (LPARAM)&es) && es.dwError == 0)
			{
				// Сбросить флаг модификации
				SendMessage(hEdit, EM_SETMODIFY, FALSE, 0);
				// Запомнить имя файла для последующего сохранения с помощью команды "Save"
				bIsOpenF = TRUE;
				lstrcpy(FullPath, fullPath);				
			}
			
			// Закрыть файл
			CloseHandle(hFile);
		}				
	}
}
// Сохранить файл. Если второй параметр FALSE, то 
// спросиь имя файла, иначе использовать имя из переменной FullPath
void NotePadDlg::saveFileCmd(HWND hwnd, BOOL hasName)
{
	if (!hasName)
	{
		TCHAR fullPath[MAX_PATH] = { 0 };
		// Получить имя файла
		if (!showSaveFileDlg(hwnd, fullPath))
		{
			return;
		}
		// Запомнить имя файла для последующего сохранения с помощью команды "Save"
		lstrcpy(FullPath, fullPath);
	}

	// Сохранять только если текст изменен с последнего сохранения или еще не сохранялся
	if (isModified() || !hasName)
	{
		// Будем использовать сообщение EM_STREAMOUT для правильной выгрузки текста в файл

		// Создать файл
		HANDLE hFile = CreateFile(FullPath, GENERIC_WRITE, FILE_SHARE_READ, 0, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
		
		if (hFile != INVALID_HANDLE_VALUE)
		{
			// Настроить EDITSTREAM
			EDITSTREAM es = { 0 };
			es.pfnCallback = writeStreamCallback;
			es.dwCookie = (DWORD_PTR)hFile;

			// Если расширкние файла RTF, тогда сохранить как RTF, иначе как plain text
			DWORD wParam;
			if (isRTF(FullPath))
			{
				wParam = SF_RTF;
			}
			else
			{
				wParam = SF_TEXT;
			}

			// Сохранить текст
			if (SendMessage(hEdit, EM_STREAMOUT, wParam, (LPARAM)&es) && es.dwError == 0)
			{
				// Сбросить флаг модификации
				SendMessage(hEdit, EM_SETMODIFY, FALSE, 0);
				// Имя файла сохранено в переменной FullPath
				bIsOpenF = TRUE;				
			}
			// Закрыть файл
			CloseHandle(hFile);
		}		
	}	
}

// Выбрать свойства шрифта и установить шрифт для текущего выделения
void NotePadDlg::chooseFontCmd(HWND hwnd)
{
	CHOOSEFONT cf;            // common dialog box structure	

	// Инициализировать структуру CHOOSEFONT
	ZeroMemory(&cf, sizeof(cf));
	cf.lStructSize = sizeof(cf);
	cf.hwndOwner = hwnd;
	cf.lpLogFont = &currFont;
	cf.rgbColors = currColor;
	cf.Flags = CF_SCREENFONTS | CF_EFFECTS | CF_INITTOLOGFONTSTRUCT;

	// Выбрать свойства шрифта
	if (ChooseFont(&cf) == TRUE)
	{
		// Установить шрифт для текущего выделения

		// Инициализировать структуру CHARFORMAT2
		CHARFORMAT2 chf;
		ZeroMemory(&chf, sizeof(chf));
		chf.cbSize = sizeof(chf);

		// Получить текущий формат выделенного фрагмента
		SendMessage(hEdit, EM_GETCHARFORMAT, SCF_SELECTION, (LPARAM)&chf);

		// Модифицировать структуру CHARFORMAT2 в соответствии с данными, полученными из структуры CHOOSEFONT

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

		// Установить формат выделенного фрагмента
		SendMessage(hEdit, EM_SETCHARFORMAT, SCF_SELECTION, (LPARAM)&chf);

		// Запомнить текущий цвет
		currColor = cf.rgbColors;
	}
}

// Выбрать цвет фона
void NotePadDlg::chooseBgColorCmd(HWND hwnd)
{
	CHOOSECOLOR cc;                 // common dialog box structure	
	
	// Инициализировать структуру CHOOSECOLOR 
	ZeroMemory(&cc, sizeof(cc));
	cc.lStructSize = sizeof(cc);
	cc.hwndOwner = hwnd;	
	cc.lpCustColors = (LPDWORD)acrCustClr;
	cc.rgbResult = currBgColor;
	cc.Flags = CC_FULLOPEN | CC_RGBINIT;

	// Показать диалог выбора цвета
	if (ChooseColor(&cc) == TRUE)
	{
		// Инициализировать структуру CHARFORMAT2 
		CHARFORMAT2 chf;
		ZeroMemory(&chf, sizeof(chf));
		chf.cbSize = sizeof(chf);

		chf.dwMask = CFM_BACKCOLOR; // Меняем только цвет фона
		chf.crBackColor = cc.rgbResult;

		// Установить формат выделенного фрагмента
		SendMessage(hEdit, EM_SETCHARFORMAT, SCF_SELECTION, (LPARAM)&chf);

		// Запомнить текущий цвет фона
		currBgColor = cc.rgbResult;
	}
}

// Вставить картинку в документ
void NotePadDlg::insertImageCmd(HWND hwnd)
{
	// Открыть файл с кортинкой и получить HBITMAP
	HBITMAP hBitmap = chooseImage(hwnd);

	if (hBitmap)
	{
		// Использовать OLE для вставки картинки в документ
		
		
		// Получить указатель на OLE-интерфейс
		LPRICHEDITOLE pRichEditOle;
		SendMessage(hEdit, EM_GETOLEINTERFACE, 0, (LPARAM)&pRichEditOle);

		// Вставить картинку в документ (см. ImageDataObject.h)
		ImageDataObject::InsertBitmap(pRichEditOle, hBitmap);

		DeleteObject(hBitmap);

		pRichEditOle->Release();
	}
}

// Показать статистику документа в MessageBox
void NotePadDlg::showFileInfoCmd(HWND hwnd)
{
	// Расчитать статистику
	Stat stat = calcStat();

	// Напечатать информацию в string stream
	std::wostringstream woss;
	woss << "Length:      " << stat.totalChars << "\r";
	woss << "Characters:  " << stat.numChars  << "\r";
	woss << "Words:       " << stat.numWords<< "\r";
	woss << "Lines:       " << stat.numLines << "\r";
	woss << "Pages:       " << stat.numPages << "\r";

	MessageBox(hwnd, woss.str().c_str(), L"Информация о файле", MB_OK);	
}

// Изменить выравнивание выбранного фрагмента
void NotePadDlg::changeParaAlignmentCmd(HWND hwnd, WORD wAlignment)
{
	// Инициализировать структуру PARAFORMAT2
	PARAFORMAT2 pf;
	ZeroMemory(&pf, sizeof(pf));
	pf.cbSize = sizeof(pf);
	pf.dwMask = PFM_ALIGNMENT;
	pf.wAlignment = wAlignment;

	// Задать выравнивание выбранного фрагмента
	SendMessage(hEdit, EM_SETPARAFORMAT, 0, (LPARAM)&pf);
}

// Изменить отступ первой строки параграфа выбранного фрагмента
void NotePadDlg::changeParaStartIndentCmd(HWND hwnd, LONG twips)
{
	// Инициализировать структуру PARAFORMAT2
	PARAFORMAT2 pf;
	ZeroMemory(&pf, sizeof(pf));
	pf.cbSize = sizeof(pf);

	// Получить текущий формат параграфа
	SendMessage(hEdit, EM_GETPARAFORMAT, 0, (LPARAM)&pf);

	// Изменить отступ первой строки на значенние заданное в переменной twips
	pf.dwMask = PFM_STARTINDENT;
	pf.dxStartIndent += twips;
	SendMessage(hEdit, EM_SETPARAFORMAT, 0, (LPARAM)&pf);
}

// Изменить отступ справа выбранного фрагмента
void NotePadDlg::changeParaRightIndentCmd(HWND hwnd, LONG twips)
{
	// Инициализировать структуру PARAFORMAT2
	PARAFORMAT2 pf;
	ZeroMemory(&pf, sizeof(pf));
	pf.cbSize = sizeof(pf);

	// Получить текущий формат параграфа
	SendMessage(hEdit, EM_GETPARAFORMAT, 0, (LPARAM)&pf);

	// Изменить отступ справа  на значенние заданное в переменной twips
	pf.dwMask = PFM_RIGHTINDENT;
	pf.dxRightIndent += twips;
	SendMessage(hEdit, EM_SETPARAFORMAT, 0, (LPARAM)&pf);
}

// Изменить отступ остальных строк параграфа относительно первой строки
void NotePadDlg::changeParaOffsetCmd(HWND hwnd, LONG twips)
{
	// Инициализировать структуру PARAFORMAT2
	PARAFORMAT2 pf;
	ZeroMemory(&pf, sizeof(pf));
	pf.cbSize = sizeof(pf);

	// Получить текущий формат параграфа
	SendMessage(hEdit, EM_GETPARAFORMAT, 0, (LPARAM)&pf);

	// Изменить отступ на значенние заданное в переменной twips
	pf.dwMask = PFM_OFFSET;
	pf.dxOffset += twips;
	SendMessage(hEdit, EM_SETPARAFORMAT, 0, (LPARAM)&pf);
}

// Изменить ширину области редактирования на значенние заданное в переменной pixels
void NotePadDlg::changePageWidth(HWND hwnd, LONG pixels)
{
	currPageWidth += pixels;

	// Получить размер клиентской части RichEdit
	RECT rc;
	GetClientRect(hEdit, &rc);

	// Задать размеры области редактирования
	rc.left += 20;
	rc.right = rc.left + currPageWidth;
	rc.top += 10;
	rc.bottom += 0;

	SendMessage(hEdit, EM_SETRECT, 0, (LPARAM)&rc);
}

// Показать диалог открытия файла. Имя файла возвращается в переменной buffer
BOOL NotePadDlg::showOpenFileDlg(HWND hwnd, TCHAR buffer[MAX_PATH])
{
	// Очистить буфер
	buffer[0] = L'\0';
	
	// Инициализировать структуру OPENFILENAME
	OPENFILENAME open = { sizeof(OPENFILENAME) };
	open.hwndOwner = hwnd;
	open.lpstrFilter = TEXT("Text Files(*.txt)\0*.txt\0RTF Files(*.rtf)\0*.rtf\0All Files(*.*)\0*.*\0");
	open.lpstrFile = buffer;
	open.nMaxFile = MAX_PATH;
	open.lpstrInitialDir = TEXT("C:\\");
	open.Flags = OFN_FILEMUSTEXIST | OFN_HIDEREADONLY;
	// Показать диалог системы для выбора файла
	return GetOpenFileName(&open);
}

// Показать диалог сохранения файла. Имя файла возвращается в переменной buffer
BOOL NotePadDlg::showSaveFileDlg(HWND hwnd, TCHAR buffer[MAX_PATH])
{
	// Очистить буфер
	buffer[0] = L'\0';
	// Инициализировать структуру OPENFILENAME
	OPENFILENAME open = { sizeof(OPENFILENAME) };
	open.hwndOwner = hwnd;
	open.lpstrFilter = TEXT("Text Files(*.txt)\0*.txt\0RTF Files(*.rtf)\0*.rtf\0All Files(*.*)\0*.*\0");
	open.lpstrFile = buffer;
	open.nMaxFile = MAX_PATH;
	open.Flags = OFN_OVERWRITEPROMPT | OFN_PATHMUSTEXIST;	
	
	// Показать диалог системы для выбора файла
	if (GetSaveFileName(&open))
	{
		// Проверить, если пользователь не ввел расширение
		if (open.nFileExtension == 0)
		{
			if (open.nFilterIndex == 1)
			{
				// Добавить .txt
				lstrcat(buffer, L".txt");
			}
			else if (open.nFilterIndex == 2)
			{
				// Добавить .rtf
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

// Выбрать изображение для вставки, загрузить и вернуть HBITMAP
HBITMAP NotePadDlg::chooseImage(HWND hwnd)
{
	// По каким то непонятным причинам вызов GetOpenFileName() приводит
	// к выделению всего текста в RichEdit. Поэтому я сохраняю текущее выделение
	// и потом восстанавливаю после GetOpenFileName().
	CHARRANGE cr;
	SendMessage(hEdit, EM_EXGETSEL, 0, (LPARAM)&cr);

	
	// Буфер для имени файла
	TCHAR lpstrFile[MAX_PATH] = {0};
	// Инициализировать структуру OPENFILENAME
	OPENFILENAME open = { sizeof(OPENFILENAME) };
	open.hwndOwner = hwnd;
	open.lpstrFilter = TEXT("Image Files (*.jpg, *.png, *.bmp)\0*.jpg;*.png;*.bmp\0");
	open.lpstrFile = lpstrFile;
	open.nMaxFile = MAX_PATH;
	open.lpstrInitialDir = TEXT("C:\\");
	open.Flags = OFN_FILEMUSTEXIST | OFN_HIDEREADONLY;
	// Получить имя файла картинки
	BOOL result = GetOpenFileName(&open);
	
	// Восстановить выделение в RichEdit
	SendMessage(hEdit, EM_EXSETSEL, 0, (LPARAM)&cr);

	if (!result)
	{		
		return NULL;
	}	
	
	HBITMAP hBitmap = NULL;

	// Загрузить картинку используя GDI+
	Bitmap * bitmap = Bitmap::FromFile(lpstrFile);
	// Получить дескриптор HBITMAP
	bitmap->GetHBITMAP(Color(255, 255, 255), &hBitmap);
	delete bitmap;

	return hBitmap;
}

// Вычислить некоторую статистическую информацию о документе.
Stat NotePadDlg::calcStat()
{
	// Получить весь текст в буфер
	int len = SendMessage(hEdit, WM_GETTEXTLENGTH, 0, 0);
	TCHAR* buffer = new TCHAR[len + 1];
	SendMessage(hEdit, WM_GETTEXT, len + 1, (LPARAM)buffer);

	Stat stat;
	stat.totalChars = len;

	std::wstring word; // текущее слово
	TCHAR ch; // текущий символ

	// Для каждого символа:
	for (int i = 0; i < len; ++i)
	{
		ch = buffer[i];

		if (ch != L' ' && ch != L'\r' && ch != L'\n')
		{
			// Если это не пробельный символ, то увеличить счетчик символов
			stat.numChars += 1;
		}

		if (IsCharAlphaNumeric(ch) || ch == L'\'' || ch == L'-')
		{
			// Если это алфавитно-цифровой символ, или апостроф, или дефис,
			// то добавить символ к текущему слову
			word += ch;
		}
		else
		{
			// если это символ новой строки, то увеличить счетчик строк
			if (ch == L'\r')
			{
				stat.numLines += 1;
			}

			if (!word.empty())
			{
				// если текущее слово не пустое, то очистить его
				word.clear();
				// и увеличить счетчик слов
				stat.numWords += 1;
			}
		}
	}

	// После всего:

	if (!word.empty())
	{
		// если текущее слово не пустое, то увеличить счетчик слов
		stat.numWords += 1;
	}

	// если последний символ был не символ новой строки, то увелмчить счетчик строк
	if (len > 0 && ch != L'\r' && ch != L'\n')
	{
		stat.numLines += 1;
	}

	delete[] buffer;
	
	return stat;
}

// Возвращает TRUE если документ был модифицирован
BOOL NotePadDlg::isModified()
{
	return SendMessage(hEdit, EM_GETMODIFY, 0, 0);
}

// Возвращает TRUE если расширение файла RTF
BOOL NotePadDlg::isRTF(LPCTSTR lpszFileName)
{
	int len = lstrlen(lpszFileName);
	return len >= 4 && lstrcmpi(L".rtf", lpszFileName + len - 4) == 0;
}
