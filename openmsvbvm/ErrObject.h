

/* this ALWAYS GENERATED file contains the definitions for the interfaces */


 /* File created by MIDL compiler version 8.01.0628 */
/* at Tue Jan 19 00:14:07 2038
 */
/* Compiler settings for ErrObject.idl:
    Oicf, W1, Zp8, env=Win32 (32b run), target_arch=X86 8.01.0628 
    protocol : dce , ms_ext, c_ext, robust
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

#ifndef __ErrObject_h__
#define __ErrObject_h__

#if defined(_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

#ifndef DECLSPEC_XFGVIRT
#if defined(_CONTROL_FLOW_GUARD_XFG)
#define DECLSPEC_XFGVIRT(base, func) __declspec(xfg_virtual(base, func))
#else
#define DECLSPEC_XFGVIRT(base, func)
#endif
#endif

/* Forward Declarations */ 

#ifndef __ErrObject_FWD_DEFINED__
#define __ErrObject_FWD_DEFINED__

#ifdef __cplusplus
typedef class ErrObject ErrObject;
#else
typedef struct ErrObject ErrObject;
#endif /* __cplusplus */

#endif 	/* __ErrObject_FWD_DEFINED__ */


#ifndef ___ErrObject_FWD_DEFINED__
#define ___ErrObject_FWD_DEFINED__
typedef interface _ErrObject _ErrObject;

#endif 	/* ___ErrObject_FWD_DEFINED__ */


/* header files for imported files */
#include "oaidl.h"
#include "ocidl.h"

#ifdef __cplusplus
extern "C"{
#endif 


#ifndef ___ErrObject_INTERFACE_DEFINED__
#define ___ErrObject_INTERFACE_DEFINED__

/* interface _ErrObject */
/* [object][nonextensible][hidden][helpstring][uuid] */ 


EXTERN_C const IID IID__ErrObject;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("A4C466B8-499F-101B-BB78-00AA00383CBB")
    _ErrObject : public IDispatch
    {
    public:
        virtual /* [helpstring][propget] */ HRESULT __stdcall get_Number( 
            /* [retval][out] */ long *rhs) = 0;
        
        virtual /* [helpstring][propput] */ HRESULT __stdcall put_Number( 
            /* [in] */ long rhs) = 0;
        
        virtual /* [helpstring][propget] */ HRESULT __stdcall get_Source( 
            /* [retval][out] */ BSTR *rhs) = 0;
        
        virtual /* [helpstring][propput] */ HRESULT __stdcall put_Source( 
            /* [in] */ BSTR rhs) = 0;
        
        virtual /* [helpstring][propget] */ HRESULT __stdcall get_Description( 
            /* [retval][out] */ BSTR *rhs) = 0;
        
        virtual /* [helpstring][propput] */ HRESULT __stdcall put_Description( 
            /* [in] */ BSTR rhs) = 0;
        
        virtual /* [helpstring][propget] */ HRESULT __stdcall get_HelpFile( 
            /* [retval][out] */ BSTR *rhs) = 0;
        
        virtual /* [helpstring][propput] */ HRESULT __stdcall put_HelpFile( 
            /* [in] */ BSTR rhs) = 0;
        
        virtual /* [helpstring][propget] */ HRESULT __stdcall get_HelpContext( 
            /* [retval][out] */ long *rhs) = 0;
        
        virtual /* [helpstring][propput] */ HRESULT __stdcall put_HelpContext( 
            /* [in] */ long rhs) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall Raise( 
            /* [in] */ long Number,
            /* [optional][in] */ VARIANT *Source,
            /* [optional][in] */ VARIANT *Description,
            /* [optional][in] */ VARIANT *HelpFile,
            /* [optional][in] */ VARIANT *HelpContext) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall Clear( void) = 0;
        
        virtual /* [helpstring][propget] */ HRESULT __stdcall get_LastDllError( 
            /* [retval][out] */ long *rhs) = 0;
        
    };
    
    
#else 	/* C style interface */

    typedef struct _ErrObjectVtbl
    {
        BEGIN_INTERFACE
        
        DECLSPEC_XFGVIRT(IUnknown, QueryInterface)
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            _ErrObject * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        DECLSPEC_XFGVIRT(IUnknown, AddRef)
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            _ErrObject * This);
        
        DECLSPEC_XFGVIRT(IUnknown, Release)
        ULONG ( STDMETHODCALLTYPE *Release )( 
            _ErrObject * This);
        
        DECLSPEC_XFGVIRT(IDispatch, GetTypeInfoCount)
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            _ErrObject * This,
            /* [out] */ UINT *pctinfo);
        
        DECLSPEC_XFGVIRT(IDispatch, GetTypeInfo)
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            _ErrObject * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        DECLSPEC_XFGVIRT(IDispatch, GetIDsOfNames)
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            _ErrObject * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        DECLSPEC_XFGVIRT(IDispatch, Invoke)
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            _ErrObject * This,
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
        
        DECLSPEC_XFGVIRT(_ErrObject, get_Number)
        /* [helpstring][propget] */ HRESULT ( __stdcall *get_Number )( 
            _ErrObject * This,
            /* [retval][out] */ long *rhs);
        
        DECLSPEC_XFGVIRT(_ErrObject, put_Number)
        /* [helpstring][propput] */ HRESULT ( __stdcall *put_Number )( 
            _ErrObject * This,
            /* [in] */ long rhs);
        
        DECLSPEC_XFGVIRT(_ErrObject, get_Source)
        /* [helpstring][propget] */ HRESULT ( __stdcall *get_Source )( 
            _ErrObject * This,
            /* [retval][out] */ BSTR *rhs);
        
        DECLSPEC_XFGVIRT(_ErrObject, put_Source)
        /* [helpstring][propput] */ HRESULT ( __stdcall *put_Source )( 
            _ErrObject * This,
            /* [in] */ BSTR rhs);
        
        DECLSPEC_XFGVIRT(_ErrObject, get_Description)
        /* [helpstring][propget] */ HRESULT ( __stdcall *get_Description )( 
            _ErrObject * This,
            /* [retval][out] */ BSTR *rhs);
        
        DECLSPEC_XFGVIRT(_ErrObject, put_Description)
        /* [helpstring][propput] */ HRESULT ( __stdcall *put_Description )( 
            _ErrObject * This,
            /* [in] */ BSTR rhs);
        
        DECLSPEC_XFGVIRT(_ErrObject, get_HelpFile)
        /* [helpstring][propget] */ HRESULT ( __stdcall *get_HelpFile )( 
            _ErrObject * This,
            /* [retval][out] */ BSTR *rhs);
        
        DECLSPEC_XFGVIRT(_ErrObject, put_HelpFile)
        /* [helpstring][propput] */ HRESULT ( __stdcall *put_HelpFile )( 
            _ErrObject * This,
            /* [in] */ BSTR rhs);
        
        DECLSPEC_XFGVIRT(_ErrObject, get_HelpContext)
        /* [helpstring][propget] */ HRESULT ( __stdcall *get_HelpContext )( 
            _ErrObject * This,
            /* [retval][out] */ long *rhs);
        
        DECLSPEC_XFGVIRT(_ErrObject, put_HelpContext)
        /* [helpstring][propput] */ HRESULT ( __stdcall *put_HelpContext )( 
            _ErrObject * This,
            /* [in] */ long rhs);
        
        DECLSPEC_XFGVIRT(_ErrObject, Raise)
        /* [helpstring] */ HRESULT ( __stdcall *Raise )( 
            _ErrObject * This,
            /* [in] */ long Number,
            /* [optional][in] */ VARIANT *Source,
            /* [optional][in] */ VARIANT *Description,
            /* [optional][in] */ VARIANT *HelpFile,
            /* [optional][in] */ VARIANT *HelpContext);
        
        DECLSPEC_XFGVIRT(_ErrObject, Clear)
        /* [helpstring] */ HRESULT ( __stdcall *Clear )( 
            _ErrObject * This);
        
        DECLSPEC_XFGVIRT(_ErrObject, get_LastDllError)
        /* [helpstring][propget] */ HRESULT ( __stdcall *get_LastDllError )( 
            _ErrObject * This,
            /* [retval][out] */ long *rhs);
        
        END_INTERFACE
    } _ErrObjectVtbl;

    interface _ErrObject
    {
        CONST_VTBL struct _ErrObjectVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define _ErrObject_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define _ErrObject_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define _ErrObject_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define _ErrObject_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define _ErrObject_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define _ErrObject_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define _ErrObject_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define _ErrObject_get_Number(This,rhs)	\
    ( (This)->lpVtbl -> get_Number(This,rhs) ) 

#define _ErrObject_put_Number(This,rhs)	\
    ( (This)->lpVtbl -> put_Number(This,rhs) ) 

#define _ErrObject_get_Source(This,rhs)	\
    ( (This)->lpVtbl -> get_Source(This,rhs) ) 

#define _ErrObject_put_Source(This,rhs)	\
    ( (This)->lpVtbl -> put_Source(This,rhs) ) 

#define _ErrObject_get_Description(This,rhs)	\
    ( (This)->lpVtbl -> get_Description(This,rhs) ) 

#define _ErrObject_put_Description(This,rhs)	\
    ( (This)->lpVtbl -> put_Description(This,rhs) ) 

#define _ErrObject_get_HelpFile(This,rhs)	\
    ( (This)->lpVtbl -> get_HelpFile(This,rhs) ) 

#define _ErrObject_put_HelpFile(This,rhs)	\
    ( (This)->lpVtbl -> put_HelpFile(This,rhs) ) 

#define _ErrObject_get_HelpContext(This,rhs)	\
    ( (This)->lpVtbl -> get_HelpContext(This,rhs) ) 

#define _ErrObject_put_HelpContext(This,rhs)	\
    ( (This)->lpVtbl -> put_HelpContext(This,rhs) ) 

#define _ErrObject_Raise(This,Number,Source,Description,HelpFile,HelpContext)	\
    ( (This)->lpVtbl -> Raise(This,Number,Source,Description,HelpFile,HelpContext) ) 

#define _ErrObject_Clear(This)	\
    ( (This)->lpVtbl -> Clear(This) ) 

#define _ErrObject_get_LastDllError(This,rhs)	\
    ( (This)->lpVtbl -> get_LastDllError(This,rhs) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* ___ErrObject_INTERFACE_DEFINED__ */


/* Additional Prototypes for ALL interfaces */

unsigned long             __RPC_USER  BSTR_UserSize(     unsigned long *, unsigned long            , BSTR * ); 
unsigned char * __RPC_USER  BSTR_UserMarshal(  unsigned long *, unsigned char *, BSTR * ); 
unsigned char * __RPC_USER  BSTR_UserUnmarshal(unsigned long *, unsigned char *, BSTR * ); 
void                      __RPC_USER  BSTR_UserFree(     unsigned long *, BSTR * ); 

unsigned long             __RPC_USER  VARIANT_UserSize(     unsigned long *, unsigned long            , VARIANT * ); 
unsigned char * __RPC_USER  VARIANT_UserMarshal(  unsigned long *, unsigned char *, VARIANT * ); 
unsigned char * __RPC_USER  VARIANT_UserUnmarshal(unsigned long *, unsigned char *, VARIANT * ); 
void                      __RPC_USER  VARIANT_UserFree(     unsigned long *, VARIANT * ); 

unsigned long             __RPC_USER  BSTR_UserSize64(     unsigned long *, unsigned long            , BSTR * ); 
unsigned char * __RPC_USER  BSTR_UserMarshal64(  unsigned long *, unsigned char *, BSTR * ); 
unsigned char * __RPC_USER  BSTR_UserUnmarshal64(unsigned long *, unsigned char *, BSTR * ); 
void                      __RPC_USER  BSTR_UserFree64(     unsigned long *, BSTR * ); 

unsigned long             __RPC_USER  VARIANT_UserSize64(     unsigned long *, unsigned long            , VARIANT * ); 
unsigned char * __RPC_USER  VARIANT_UserMarshal64(  unsigned long *, unsigned char *, VARIANT * ); 
unsigned char * __RPC_USER  VARIANT_UserUnmarshal64(unsigned long *, unsigned char *, VARIANT * ); 
void                      __RPC_USER  VARIANT_UserFree64(     unsigned long *, VARIANT * ); 

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif


