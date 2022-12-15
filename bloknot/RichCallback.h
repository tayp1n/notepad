#pragma once
#include <ObjIdl.h>
#include <richedit.h>
#include <richole.h>

// Это реализация интерфейса IRichEditOleCallback
// используется для корректной работы с объектами, такими как картинки.
// В частности, это позволяет использовать Copy/Paste для OLE-объектов,
// а также обеспечивает корректную загрузку и сохранение файлов, 
// содержащих такие объекты.
// Основная работа этого класса - вызов функций StgCreateStorageEx и CreateStorage.
class RichCallback :  public IRichEditOleCallback
{
private:
	ULONG	m_ref, m_stgNum;
	LPSTORAGE pStorage;
public:
	// Constructor
	RichCallback()
	{
		HRESULT hr = E_FAIL;

		hr = StgCreateStorageEx(NULL, STGM_READWRITE | STGM_SHARE_EXCLUSIVE | STGM_CREATE | STGM_TRANSACTED | STGM_DELETEONRELEASE,
			STGFMT_STORAGE, 0, NULL, NULL, IID_IStorage, (LPVOID*)&pStorage);
	}

	// IUnknown methods
	STDMETHOD(QueryInterface) (REFIID riid, LPVOID FAR* ppvObj)
	{
		HRESULT	hr = E_NOINTERFACE;

		if (ppvObj == NULL)
			hr = E_POINTER;
		else
		{
			*ppvObj = NULL;

			if (riid == IID_IRichEditOleCallback || riid == IID_IUnknown)
			{
				*ppvObj = (LPVOID)this;
				AddRef();
				hr = S_OK;
			}
		}

		return hr;
	}

	STDMETHOD_(ULONG, AddRef) ()
	{
		m_ref++;
		return m_ref;
	}

	STDMETHOD_(ULONG, Release) ()
	{
		if (--m_ref == 0)
		{
			delete this;
		}

		return m_ref;
	}

	// IRichEditOleCallback methods 
	STDMETHOD(GetNewStorage) (LPSTORAGE FAR* ppstg)
	{
		HRESULT hr = E_FAIL;
		WCHAR stgName[31], stgNum[8];

		wcscpy_s(stgName, _countof(stgName), L"My_Notepad_Storage");
		_ultow_s(++m_stgNum, stgNum, _countof(stgNum), 10);
		wcscat_s(stgName, _countof(stgName), stgNum);

		hr = pStorage->CreateStorage(stgName, STGM_READWRITE | STGM_SHARE_EXCLUSIVE | STGM_CREATE | STGM_TRANSACTED, 0, 0, ppstg);

		return hr;
	}

	STDMETHOD(GetInPlaceContext) (LPOLEINPLACEFRAME FAR* lplpFrame,	LPOLEINPLACEUIWINDOW FAR* lplpDoc, LPOLEINPLACEFRAMEINFO lpFrameInfo)
	{
		return E_NOTIMPL;
	}

	STDMETHOD(ShowContainerUI) (BOOL fShow)
	{
		return E_NOTIMPL;
	}

	STDMETHOD(QueryInsertObject) (LPCLSID lpclsid, LPSTORAGE lpstg,	LONG cp)
	{
		return S_OK;
	}

	STDMETHOD(DeleteObject) (LPOLEOBJECT lpoleobj)
	{
		return S_OK;
	}

	STDMETHOD(QueryAcceptData) (LPDATAOBJECT lpdataobj,	CLIPFORMAT FAR* lpcfFormat, DWORD reco,	BOOL fReally, HGLOBAL hMetaPict)
	{
		return E_NOTIMPL;
	}

	STDMETHOD(ContextSensitiveHelp) (BOOL fEnterMode)
	{
		return E_NOTIMPL;
	}

	STDMETHOD(GetClipboardData) (CHARRANGE FAR* lpchrg, DWORD reco,	LPDATAOBJECT FAR* lplpdataobj)
	{
		return E_NOTIMPL;
	}

	STDMETHOD(GetDragDropEffect) (BOOL fDrag, DWORD grfKeyState, LPDWORD pdwEffect)
	{
		return E_NOTIMPL;
	}

	STDMETHOD(GetContextMenu) (WORD seltype, LPOLEOBJECT lpoleobj, CHARRANGE FAR* lpchrg, HMENU FAR* lphmenu)
	{
		return E_NOTIMPL;
	}
};

