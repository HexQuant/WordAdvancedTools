

/* this ALWAYS GENERATED file contains the definitions for the interfaces */


 /* File created by MIDL compiler version 8.01.0622 */
/* at Tue Jan 19 06:14:07 2038
 */
/* Compiler settings for WordAdvancedTools.idl:
    Oicf, W1, Zp8, env=Win64 (32b run), target_arch=AMD64 8.01.0622 
    protocol : all , ms_ext, c_ext, robust
    error checks: allocation ref bounds_check enum stub_data 
    VC __declspec() decoration level: 
         __declspec(uuid()), __declspec(selectany), __declspec(novtable)
         DECLSPEC_UUID(), MIDL_INTERFACE()
*/
/* @@MIDL_FILE_HEADING(  ) */



/* verify that the <rpcndr.h> version is high enough to compile this file*/
#ifndef __REQUIRED_RPCNDR_H_VERSION__
#define __REQUIRED_RPCNDR_H_VERSION__ 500
#endif

#include "rpc.h"
#include "rpcndr.h"

#ifndef __RPCNDR_H_VERSION__
#error this stub requires an updated version of <rpcndr.h>
#endif /* __RPCNDR_H_VERSION__ */

#ifndef COM_NO_WINDOWS_H
#include "windows.h"
#include "ole2.h"
#endif /*COM_NO_WINDOWS_H*/

#ifndef __WordAdvancedTools_i_h__
#define __WordAdvancedTools_i_h__

#if defined(_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

/* Forward Declarations */ 

#ifndef __IConnect_FWD_DEFINED__
#define __IConnect_FWD_DEFINED__
typedef interface IConnect IConnect;

#endif 	/* __IConnect_FWD_DEFINED__ */


#ifndef __Connect_FWD_DEFINED__
#define __Connect_FWD_DEFINED__

#ifdef __cplusplus
typedef class Connect Connect;
#else
typedef struct Connect Connect;
#endif /* __cplusplus */

#endif 	/* __Connect_FWD_DEFINED__ */


#ifndef __IRibbonCallback_FWD_DEFINED__
#define __IRibbonCallback_FWD_DEFINED__
typedef interface IRibbonCallback IRibbonCallback;

#endif 	/* __IRibbonCallback_FWD_DEFINED__ */


/* header files for imported files */
#include "oaidl.h"
#include "ocidl.h"
#include "shobjidl.h"

#ifdef __cplusplus
extern "C"{
#endif 


#ifndef __IConnect_INTERFACE_DEFINED__
#define __IConnect_INTERFACE_DEFINED__

/* interface IConnect */
/* [unique][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_IConnect;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("fd2efaae-869d-457a-8fb9-b33833e9553c")
    IConnect : public IDispatch
    {
    public:
    };
    
    
#else 	/* C style interface */

    typedef struct IConnectVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IConnect * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IConnect * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IConnect * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IConnect * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IConnect * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IConnect * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IConnect * This,
            /* [annotation][in] */ 
            _In_  DISPID dispIdMember,
            /* [annotation][in] */ 
            _In_  REFIID riid,
            /* [annotation][in] */ 
            _In_  LCID lcid,
            /* [annotation][in] */ 
            _In_  WORD wFlags,
            /* [annotation][out][in] */ 
            _In_  DISPPARAMS *pDispParams,
            /* [annotation][out] */ 
            _Out_opt_  VARIANT *pVarResult,
            /* [annotation][out] */ 
            _Out_opt_  EXCEPINFO *pExcepInfo,
            /* [annotation][out] */ 
            _Out_opt_  UINT *puArgErr);
        
        END_INTERFACE
    } IConnectVtbl;

    interface IConnect
    {
        CONST_VTBL struct IConnectVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IConnect_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IConnect_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IConnect_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IConnect_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IConnect_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IConnect_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IConnect_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __IConnect_INTERFACE_DEFINED__ */



#ifndef __WordAdvancedToolsLib_LIBRARY_DEFINED__
#define __WordAdvancedToolsLib_LIBRARY_DEFINED__

/* library WordAdvancedToolsLib */
/* [version][uuid] */ 


EXTERN_C const IID LIBID_WordAdvancedToolsLib;

EXTERN_C const CLSID CLSID_Connect;

#ifdef __cplusplus

class DECLSPEC_UUID("f6fb665a-c3ff-441b-94ab-027ab20c4f34")
Connect;
#endif
#endif /* __WordAdvancedToolsLib_LIBRARY_DEFINED__ */

#ifndef __IRibbonCallback_INTERFACE_DEFINED__
#define __IRibbonCallback_INTERFACE_DEFINED__

/* interface IRibbonCallback */
/* [unique][helpstring][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_IRibbonCallback;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("CE895442-9981-4315-AA85-4B9A5C7739D8")
    IRibbonCallback : public IDispatch
    {
    public:
        virtual HRESULT STDMETHODCALLTYPE OnAddRevisionButton( 
            /* [in] */ IDispatch *ribbonControl) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE OnPinBookmarkButton( 
            /* [in] */ IDispatch *ribbonControl) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE OnRevisionSettingsButton( 
            /* [in] */ IDispatch *ribbonControl) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE OnRemoveAllRevisionButton( 
            /* [in] */ IDispatch *ribbonControl) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE GetLabel( 
            /* [in] */ IDispatch *ribbonControl,
            /* [retval][out] */ BSTR *Label) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE GetImage( 
            /* [in] */ IDispatch *ribbonControl,
            /* [retval][out] */ IPictureDisp **ppdispImage) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE GetImage2( 
            /* [in] */ IDispatch *ribbonControl,
            /* [retval][out] */ IPictureDisp **ppdispImage) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE GetImage3( 
            /* [in] */ IDispatch *ribbonControl,
            /* [retval][out] */ IPictureDisp **ppdispImage) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE OnLoad( 
            /* [in] */ IDispatch *ribbonControl) = 0;
        
    };
    
    
#else 	/* C style interface */

    typedef struct IRibbonCallbackVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IRibbonCallback * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IRibbonCallback * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IRibbonCallback * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IRibbonCallback * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IRibbonCallback * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IRibbonCallback * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IRibbonCallback * This,
            /* [annotation][in] */ 
            _In_  DISPID dispIdMember,
            /* [annotation][in] */ 
            _In_  REFIID riid,
            /* [annotation][in] */ 
            _In_  LCID lcid,
            /* [annotation][in] */ 
            _In_  WORD wFlags,
            /* [annotation][out][in] */ 
            _In_  DISPPARAMS *pDispParams,
            /* [annotation][out] */ 
            _Out_opt_  VARIANT *pVarResult,
            /* [annotation][out] */ 
            _Out_opt_  EXCEPINFO *pExcepInfo,
            /* [annotation][out] */ 
            _Out_opt_  UINT *puArgErr);
        
        HRESULT ( STDMETHODCALLTYPE *OnAddRevisionButton )( 
            IRibbonCallback * This,
            /* [in] */ IDispatch *ribbonControl);
        
        HRESULT ( STDMETHODCALLTYPE *OnPinBookmarkButton )( 
            IRibbonCallback * This,
            /* [in] */ IDispatch *ribbonControl);
        
        HRESULT ( STDMETHODCALLTYPE *OnRevisionSettingsButton )( 
            IRibbonCallback * This,
            /* [in] */ IDispatch *ribbonControl);
        
        HRESULT ( STDMETHODCALLTYPE *OnRemoveAllRevisionButton )( 
            IRibbonCallback * This,
            /* [in] */ IDispatch *ribbonControl);
        
        HRESULT ( STDMETHODCALLTYPE *GetLabel )( 
            IRibbonCallback * This,
            /* [in] */ IDispatch *ribbonControl,
            /* [retval][out] */ BSTR *Label);
        
        HRESULT ( STDMETHODCALLTYPE *GetImage )( 
            IRibbonCallback * This,
            /* [in] */ IDispatch *ribbonControl,
            /* [retval][out] */ IPictureDisp **ppdispImage);
        
        HRESULT ( STDMETHODCALLTYPE *GetImage2 )( 
            IRibbonCallback * This,
            /* [in] */ IDispatch *ribbonControl,
            /* [retval][out] */ IPictureDisp **ppdispImage);
        
        HRESULT ( STDMETHODCALLTYPE *GetImage3 )( 
            IRibbonCallback * This,
            /* [in] */ IDispatch *ribbonControl,
            /* [retval][out] */ IPictureDisp **ppdispImage);
        
        HRESULT ( STDMETHODCALLTYPE *OnLoad )( 
            IRibbonCallback * This,
            /* [in] */ IDispatch *ribbonControl);
        
        END_INTERFACE
    } IRibbonCallbackVtbl;

    interface IRibbonCallback
    {
        CONST_VTBL struct IRibbonCallbackVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IRibbonCallback_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IRibbonCallback_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IRibbonCallback_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IRibbonCallback_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IRibbonCallback_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IRibbonCallback_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IRibbonCallback_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define IRibbonCallback_OnAddRevisionButton(This,ribbonControl)	\
    ( (This)->lpVtbl -> OnAddRevisionButton(This,ribbonControl) ) 

#define IRibbonCallback_OnPinBookmarkButton(This,ribbonControl)	\
    ( (This)->lpVtbl -> OnPinBookmarkButton(This,ribbonControl) ) 

#define IRibbonCallback_OnRevisionSettingsButton(This,ribbonControl)	\
    ( (This)->lpVtbl -> OnRevisionSettingsButton(This,ribbonControl) ) 

#define IRibbonCallback_OnRemoveAllRevisionButton(This,ribbonControl)	\
    ( (This)->lpVtbl -> OnRemoveAllRevisionButton(This,ribbonControl) ) 

#define IRibbonCallback_GetLabel(This,ribbonControl,Label)	\
    ( (This)->lpVtbl -> GetLabel(This,ribbonControl,Label) ) 

#define IRibbonCallback_GetImage(This,ribbonControl,ppdispImage)	\
    ( (This)->lpVtbl -> GetImage(This,ribbonControl,ppdispImage) ) 

#define IRibbonCallback_GetImage2(This,ribbonControl,ppdispImage)	\
    ( (This)->lpVtbl -> GetImage2(This,ribbonControl,ppdispImage) ) 

#define IRibbonCallback_GetImage3(This,ribbonControl,ppdispImage)	\
    ( (This)->lpVtbl -> GetImage3(This,ribbonControl,ppdispImage) ) 

#define IRibbonCallback_OnLoad(This,ribbonControl)	\
    ( (This)->lpVtbl -> OnLoad(This,ribbonControl) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __IRibbonCallback_INTERFACE_DEFINED__ */


/* Additional Prototypes for ALL interfaces */

unsigned long             __RPC_USER  BSTR_UserSize(     unsigned long *, unsigned long            , BSTR * ); 
unsigned char * __RPC_USER  BSTR_UserMarshal(  unsigned long *, unsigned char *, BSTR * ); 
unsigned char * __RPC_USER  BSTR_UserUnmarshal(unsigned long *, unsigned char *, BSTR * ); 
void                      __RPC_USER  BSTR_UserFree(     unsigned long *, BSTR * ); 

unsigned long             __RPC_USER  BSTR_UserSize64(     unsigned long *, unsigned long            , BSTR * ); 
unsigned char * __RPC_USER  BSTR_UserMarshal64(  unsigned long *, unsigned char *, BSTR * ); 
unsigned char * __RPC_USER  BSTR_UserUnmarshal64(unsigned long *, unsigned char *, BSTR * ); 
void                      __RPC_USER  BSTR_UserFree64(     unsigned long *, BSTR * ); 

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif


