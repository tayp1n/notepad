#pragma once
#include <ObjIdl.h>
#include <richedit.h>
#include <richole.h>

// Ёто реализаци€ интерфейса IDataObject
// используетс€ дл€ вставки картинки в документ.
//  артинка должна быть предварительно загружена в HBITMAP.
class ImageDataObject : public IDataObject
{
private:
	ULONG	m_ref;
	BOOL	m_release;

	// The data being bassed to the richedit
	//
	STGMEDIUM m_stgmed;
	FORMATETC m_fromat;
public:	
	static void InsertBitmap(IRichEditOle* pRichEditOle, HBITMAP hBitmap)
	{
		HRESULT hr;

		// Get the image data object

		ImageDataObject* ido = new ImageDataObject;
		LPDATAOBJECT lpDataObject;
		ido->QueryInterface(IID_IDataObject, (void**)&lpDataObject);

		STGMEDIUM stgm;
		stgm.tymed = TYMED_GDI;					// Storage medium = HBITMAP handle		
		stgm.hBitmap = hBitmap;
		stgm.pUnkForRelease = NULL;				// Use ReleaseStgMedium

		FORMATETC fm;
		fm.cfFormat = CF_BITMAP;				// Clipboard format = CF_BITMAP
		fm.ptd = NULL;							// Target Device = Screen
		fm.dwAspect = DVASPECT_CONTENT;			// Level of detail = Full content
		fm.lindex = -1;							// Index = Not applicaple
		fm.tymed = TYMED_GDI;					// Storage medium = HBITMAP handle

		ido->SetData(&fm, &stgm, TRUE);

		// Get the RichEdit container site

		IOleClientSite* pOleClientSite;
		pRichEditOle->GetClientSite(&pOleClientSite);

		// Initialize a Storage Object

		IStorage* pStorage;

		LPLOCKBYTES lpLockBytes = NULL;
		hr = ::CreateILockBytesOnHGlobal(NULL, TRUE, &lpLockBytes);	
		
		hr = ::StgCreateDocfileOnILockBytes(lpLockBytes, STGM_SHARE_EXCLUSIVE | STGM_CREATE | STGM_READWRITE, 0, &pStorage);
		if (hr != S_OK)
		{			
			lpLockBytes = NULL;			
		}		

		// The final ole object which will be inserted in the richedit control

		IOleObject* pOleObject;
		hr = ::OleCreateStaticFromData(ido, IID_IOleObject, OLERENDER_FORMAT, &ido->m_fromat, pOleClientSite, pStorage, (void**)&pOleObject);
		
		// all items are "contained" -- this makes our reference to this object
		//  weak -- which is needed for links to embedding silent update.
		::OleSetContainedObject(pOleObject, TRUE);

		// Now Add the object to the RichEdit 

		REOBJECT reobject;
		::ZeroMemory(&reobject, sizeof(REOBJECT));
		reobject.cbStruct = sizeof(REOBJECT);

		CLSID clsid;
		hr = pOleObject->GetUserClassID(&clsid);
		
		reobject.clsid = clsid;
		reobject.cp = REO_CP_SELECTION;
		reobject.dvaspect = DVASPECT_CONTENT;
		reobject.poleobj = pOleObject;
		reobject.polesite = pOleClientSite;
		reobject.pstg = pStorage;

		// Insert the bitmap at the current location in the richedit control
		pRichEditOle->InsertObject(&reobject);

		// Release all unnecessary interfaces
		pOleObject->Release();
		pOleClientSite->Release();
		pStorage->Release();
		lpDataObject->Release();
	}

public:
	ImageDataObject() : m_ref(0)
	{
		m_release = FALSE;
	}

	~ImageDataObject()
	{
		if (m_release)
			::ReleaseStgMedium(&m_stgmed);
	}

	// Methods of the IUnknown interface
	
	STDMETHOD(QueryInterface)(REFIID iid, void** ppvObject)
	{
		if (iid == IID_IUnknown || iid == IID_IDataObject)
		{
			*ppvObject = this;
			AddRef();
			return S_OK;
		}
		else
			return E_NOINTERFACE;
	}

	STDMETHOD_(ULONG, AddRef)(void)
	{
		m_ref++;
		return m_ref;
	}

	STDMETHOD_(ULONG, Release)(void)
	{
		if (--m_ref == 0)
		{
			delete this;
		}

		return m_ref;
	}

	// Methods of the IDataObject Interface

	STDMETHOD(GetData)(FORMATETC* pformatetcIn, STGMEDIUM* pmedium) 
	{
		HANDLE hDst;
		hDst = ::OleDuplicateData(m_stgmed.hBitmap, CF_BITMAP, NULL);
		if (hDst == NULL)
		{
			return E_HANDLE;
		}

		pmedium->tymed = TYMED_GDI;
		pmedium->hBitmap = (HBITMAP)hDst;
		pmedium->pUnkForRelease = NULL;

		return S_OK;
	}

	STDMETHOD(GetDataHere)(FORMATETC* pformatetc, STGMEDIUM* pmedium) 
	{
		return E_NOTIMPL;
	}

	STDMETHOD(QueryGetData)(FORMATETC* pformatetc) 
	{
		return E_NOTIMPL;
	}

	STDMETHOD(GetCanonicalFormatEtc)(FORMATETC* pformatectIn, FORMATETC* pformatetcOut) 
	{
		return E_NOTIMPL;
	}

	STDMETHOD(SetData)(FORMATETC* pformatetc, STGMEDIUM* pmedium, BOOL  fRelease) 
	{
		m_fromat = *pformatetc;
		m_stgmed = *pmedium;

		return S_OK;
	}

	STDMETHOD(EnumFormatEtc)(DWORD  dwDirection, IEnumFORMATETC** ppenumFormatEtc) 
	{
		return E_NOTIMPL;
	}

	STDMETHOD(DAdvise)(FORMATETC* pformatetc, DWORD advf, IAdviseSink* pAdvSink, DWORD* pdwConnection) 
	{
		return E_NOTIMPL;
	}

	STDMETHOD(DUnadvise)(DWORD dwConnection) 
	{
		return E_NOTIMPL;
	}

	STDMETHOD(EnumDAdvise)(IEnumSTATDATA** ppenumAdvise) 
	{
		return E_NOTIMPL;
	}
};

