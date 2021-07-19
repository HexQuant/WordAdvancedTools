// Connect.h : Declaration of the CConnect

#pragma once
#include "resource.h"       // main symbols

#include <atlstr.h>
#include <atlimage.h>

#include "WordAdvancedTools_i.h"

#include "RevisionMacros.h"
#include "RevisionSettingsDialog.h"


#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "Single-threaded COM objects are not properly supported on Windows CE platform, such as the Windows Mobile platforms that do not include full DCOM support. Define _CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA to force ATL to support creating single-thread COM object's and allow use of it's single-threaded COM object implementations. The threading model in your rgs file was set to 'Free' as that is the only threading model supported in non DCOM Windows CE platforms."
#endif

using namespace ATL;

typedef public IDispatchImpl<_IDTExtensibility2, &__uuidof(_IDTExtensibility2), &LIBID_AddInDesignerObjects, /* wMajor = */ 1, /* wMinor = */ 0>
IDTImpl;

typedef IDispatchImpl<IRibbonExtensibility, &__uuidof(IRibbonExtensibility), &__uuidof(__Office), /* wMajor = */ 2, /* wMinor = */ 5>
RibbonImpl;

typedef IDispatchImpl<IRibbonCallback, &__uuidof(IRibbonCallback)>
CallbackImpl;


// CConnect

class ATL_NO_VTABLE CConnect :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CConnect, &CLSID_Connect>,
	public IDispatchImpl<IConnect, &IID_IConnect, &LIBID_WordAdvancedToolsLib, /*wMajor =*/ 1, /*wMinor =*/ 0>,
	public IDTImpl,
	public RibbonImpl,
	public CallbackImpl
	//public IDispatchImpl<_IDTExtensibility2, &__uuidof(_IDTExtensibility2), &LIBID_AddInDesignerObjects, /* wMajor = */ 1, /* wMinor = */ 0>
{
public:
	CConnect()
	{
	}

DECLARE_REGISTRY_RESOURCEID(106)


BEGIN_COM_MAP(CConnect)
	COM_INTERFACE_ENTRY2(IDispatch, IRibbonCallback)
	COM_INTERFACE_ENTRY(IConnect)
	COM_INTERFACE_ENTRY(_IDTExtensibility2)
	COM_INTERFACE_ENTRY(IRibbonExtensibility)
	COM_INTERFACE_ENTRY(IRibbonCallback)
END_COM_MAP()



	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		return S_OK;
	}

	void FinalRelease()
	{
	}

public:

	//Ribbon Callbacks

	STDMETHOD(OnRevisionSettingsButton)(IDispatch* ribbon)
	{
		CRevisionSettingsDialog Dialog;
		Dialog.DoModal();

		return S_OK;
	}

	STDMETHOD(OnRemoveAllRevisionButton)(IDispatch* ribbon)
	{
		return S_OK;
	}

	STDMETHOD(OnAddRevisionButton)(IDispatch* ribbon)
	{
		auto ar = RevisionMacros(spApp,  SysAllocString(L"1"));
		ar.Insert();
			//MessageBoxW(NULL, MyBstr, L"Native Addin", MB_OK);
		//Office::MsoLanguageID lang;

		//spApp->get_Language(&lang);
		//WORD lang2 = MAKELANGID(PRIMARYLANGID(lang), SUBLANG_DEFAULT);

		return S_OK;

	}

	IStream* CreateStreamOnResource(LPCTSTR lpName, LPCTSTR lpType)
	{
		IStream* ipStream = NULL;

		HRSRC hrsrc = FindResource(NULL, lpName, lpType);
		if (hrsrc == NULL)
			goto Return;

		DWORD dwResourceSize = SizeofResource(NULL, hrsrc);
		HGLOBAL hglbImage = LoadResource(NULL, hrsrc);
		if (hglbImage == NULL)
			goto Return;

		LPVOID pvSourceResourceData = LockResource(hglbImage);
		if (pvSourceResourceData == NULL)
			goto Return;

		HGLOBAL hgblResourceData = GlobalAlloc(GMEM_MOVEABLE, dwResourceSize);
		if (hgblResourceData == NULL)
			goto Return;

		LPVOID pvResourceData = GlobalLock(hgblResourceData);

		if (pvResourceData == NULL)
			goto FreeData;

		CopyMemory(pvResourceData, pvSourceResourceData, dwResourceSize);

		GlobalUnlock(hgblResourceData);

		if (SUCCEEDED(CreateStreamOnHGlobal(hgblResourceData, TRUE, &ipStream)))
			goto Return;

	FreeData:
		GlobalFree(hgblResourceData);

	Return:
		return ipStream;
	}


	STDMETHOD(GetImage3)(IDispatch* ribbonControl, IPictureDisp** pPicDisp)
	//STDMETHOD(GetImage3) (BSTR* pBSTRImgName, IDispatch** pPicDisp)
	{
		PICTDESC	pt;
		CComPtr<IPictureDisp> pPic;
		HRESULT	hr = E_FAIL;
		HRSRC	hRsrc = NULL;
		HGLOBAL	hResource = NULL, hRes = NULL;
		LPBYTE  lpBuffer = NULL;
		UINT    resID = 0, uiSize = 0;
		LPVOID lpResBuffer = NULL;
		CComPtr<IStream> pStream;
		CComBSTR bstrImg;
		CImage	png;

		//HMODULE hModule = _AtlBaseModule.GGetModuleInstance();

		//ATLTRACE("%p - In loadImages, Image is %S\n", this, *pBSTRImgName);

		//bstrImg.Attach(*pBSTRImgName);

		//if (bstrImg == L"MyImage")

		IRibbonControl* RibbonCtl = static_cast<IRibbonControl*>(ribbonControl);
		CComBSTR bstr;
		RibbonCtl->get_Id(&bstr);
		if (bstr=="AddRevisionButton")
		{	
			resID = IDB_PNG1;
		}
		if (bstr=="RemoveAllRevisionButton")
		{	
			resID = IDB_PNG2;
		}
		if (bstr=="LinkRevisionButton")
		{	
			resID = IDB_PNG3;
		}
		if (bstr=="AboutButton")
		{	
			resID = IDB_PNG4;
		}



		if (resID != 0)
		{
			hRsrc = FindResource(_AtlBaseModule.GetResourceInstance(), MAKEINTRESOURCE(resID), _T("PNG"));

			if (hRsrc == NULL)
				return hr;

			hResource = LoadResource(_AtlBaseModule.GetResourceInstance(), hRsrc);

			if (hResource == NULL)
				return hr;

			lpBuffer = (LPBYTE)LockResource(hResource);

			if (lpBuffer == NULL)
				return hr;

			uiSize = SizeofResource(_AtlBaseModule.GetResourceInstance(), hRsrc);

			if (uiSize == 0)
				return hr;

			hRes = GlobalAlloc(GMEM_MOVEABLE, uiSize);

			if (hRes != NULL)
			{
				lpResBuffer = GlobalLock(hRes);

				if (lpResBuffer != NULL)
				{
					memcpy(lpResBuffer, lpBuffer, uiSize);

					hr = CreateStreamOnHGlobal(hRes, TRUE, &pStream);

					if (SUCCEEDED(hr))
						hr = png.Load(pStream);

					GlobalUnlock(hRes);
				}

				GlobalFree(hRes);
			}

			if (SUCCEEDED(hr))
			{
				SecureZeroMemory(&pt, sizeof(pt));

				pt.cbSizeofstruct = sizeof(pt);

				pt.picType = PICTYPE_BITMAP;

				pt.bmp.hbitmap = png.Detach();

				hr = OleCreatePictureIndirect(&pt, IID_IPictureDisp, TRUE, (LPVOID*)&pPic);

				if (pPic)
				{
					*pPicDisp = pPic.Detach();
				}
			}
		}

		return hr;
	}




	STDMETHOD(GetImage2)(IDispatch* ribbonControl, IPictureDisp** ppdispImage)
	{

		CImage image;
		auto hr=image.Load(CreateStreamOnResource(MAKEINTRESOURCE(IDB_PNG1), _T("PNG")));
		if (FAILED(hr)) return hr;

		HBITMAP hbm = (HBITMAP)image.Detach();
		//if (hbm)
		//{
			// Use the factory implemented by the framework to produce an IUIImage.
			//hr = m_pifbFactory->CreateImage(hbm, UI_OWNERSHIP_TRANSFER, ppimg);
			//if (FAILED(hr))
			//{
				//DeleteObject(hbm);
			//}
		//}

		//CComPtr<IPictureDisp> pPic(image.);
	
		return S_OK;
	}

	STDMETHOD(GetImage)(IDispatch* ribbonControl, IPictureDisp** ppdispImage)
	{
		HMODULE hModule = _AtlBaseModule.GetModuleInstance();
		IID IID_Picture;
		HRESULT hRes = E_FAIL;
		IIDFromString(L"{7BF80980-BF32-101A-8BBB-00AA00300CAB}", &IID_Picture);
		HANDLE hBtnImg = LoadImage(hModule, MAKEINTRESOURCE(IDB_BITMAP2), IMAGE_BITMAP, 64, 64, LR_DEFAULTSIZE);
		if (!hBtnImg) return E_POINTER;

		PICTDESC picDesc = { 0 };
		picDesc.bmp.hbitmap = (HBITMAP)hBtnImg;
		picDesc.picType = PICTYPE_BITMAP;
		picDesc.cbSizeofstruct = sizeof(picDesc);
		CComPtr<IPictureDisp> pPic;

		//CImage image;



		OleCreatePictureIndirect(&picDesc, IID_Picture, true, reinterpret_cast<LPVOID*>(&pPic));
		if (pPic)
		{
			*ppdispImage = pPic.Detach();
		}
		return S_OK;
	}

	STDMETHOD(GetLabel)(IDispatch* control, BSTR* Label)
	{
		HMODULE hModule = _AtlBaseModule.GetModuleInstance();
		
		IRibbonControl* RibbonCtl = static_cast<IRibbonControl*>(control);
		CComBSTR bstr;
		RibbonCtl->get_Id(&bstr);
		if (bstr=="RevisionGroup")
		{	
			*Label = bstr;
		}
		//MessageBoxW(NULL, bstr, L"Native Addin", MB_OK);
		CString str;
		WORD LangID = MAKELANGID(LANG_RUSSIAN, SUBLANG_DEFAULT);
		str.LoadString(hModule, IDS_REV, LangID);
		*Label = str.AllocSysString();
		return S_OK;
	}

	CComQIPtr<_Application> spApp;
	//Word::Application *spApp;

	HRESULT HrGetResource(int nId,
		LPCTSTR lpType,
		LPVOID* ppvResourceData,
		DWORD* pdwSizeInBytes)
	{
		HMODULE hModule = _AtlBaseModule.GetModuleInstance();
		if (!hModule)
			return E_UNEXPECTED;
		HRSRC hRsrc = FindResource(hModule, MAKEINTRESOURCE(nId), lpType);
		if (!hRsrc)
			return HRESULT_FROM_WIN32(GetLastError());
		HGLOBAL hGlobal = LoadResource(hModule, hRsrc);
		if (!hGlobal)
			return HRESULT_FROM_WIN32(GetLastError());
		*pdwSizeInBytes = SizeofResource(hModule, hRsrc);
		*ppvResourceData = LockResource(hGlobal);
		return S_OK;
	}

	BSTR GetXMLResource(int nId)
	{
		LPVOID pResourceData = NULL;
		DWORD dwSizeInBytes = 0;
		HRESULT hr = HrGetResource(nId, TEXT("XML"),
			&pResourceData, &dwSizeInBytes);
		if (FAILED(hr))
			return NULL;
		// Assumes that the data is not stored in Unicode.
		CComBSTR cbstr(dwSizeInBytes, reinterpret_cast<LPCSTR>(pResourceData));
		return cbstr.Detach();
	}

	SAFEARRAY* GetOFSResource(int nId)
	{
		LPVOID pResourceData = NULL;
		DWORD dwSizeInBytes = 0;
		if (FAILED(HrGetResource(nId, TEXT("OFS"),
			&pResourceData, &dwSizeInBytes)))
			return NULL;
		SAFEARRAY* psa;
		SAFEARRAYBOUND dim = { dwSizeInBytes, 0 };
		psa = SafeArrayCreate(VT_UI1, 1, &dim);
		if (psa == NULL)
			return NULL;
		BYTE* pSafeArrayData;
		SafeArrayAccessData(psa, (void**)&pSafeArrayData);
		memcpy((void*)pSafeArrayData, pResourceData, dwSizeInBytes);
		SafeArrayUnaccessData(psa);
		return psa;
	}


// Implement IRibbonExtensibility Methods
	STDMETHOD(GetCustomUI)(BSTR RibbonID, BSTR* RibbonXml)
	{
		if (!RibbonXml) [[unlikely]]
		{
			return E_POINTER;
		}

		*RibbonXml = GetXMLResource(IDR_XML1);
		return S_OK;
	}

// _IDTExtensibility2 Methods
public:
	STDMETHOD(OnConnection)(LPDISPATCH Application, ext_ConnectMode ConnectMode, LPDISPATCH AddInInst, SAFEARRAY * * custom)
	{
		if (!Application) [[unlikely]]
		{
			return E_POINTER;
		}

		spApp = Application;
		return S_OK;
	}

	STDMETHOD(OnDisconnection)(ext_DisconnectMode RemoveMode, SAFEARRAY * * custom)
	{
		 return S_OK;
	}

	STDMETHOD(OnAddInsUpdate)(SAFEARRAY * * custom)
	{
		 return S_OK;
	}

	STDMETHOD(OnStartupComplete)(SAFEARRAY * * custom)
	{
		 return S_OK;
	}

	STDMETHOD(OnBeginShutdown)(SAFEARRAY * * custom)
	{
		 return S_OK;
	}

};

OBJECT_ENTRY_AUTO(__uuidof(Connect), CConnect)
