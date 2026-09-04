

/* this ALWAYS GENERATED file contains the definitions for the interfaces */


 /* File created by MIDL compiler version 8.01.0628 */
/* at Tue Jan 19 00:14:07 2038
 */
/* Compiler settings for Label.idl:
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

#ifndef __Label_h__
#define __Label_h__

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

#ifndef __Label_FWD_DEFINED__
#define __Label_FWD_DEFINED__

#ifdef __cplusplus
typedef class Label Label;
#else
typedef struct Label Label;
#endif /* __cplusplus */

#endif 	/* __Label_FWD_DEFINED__ */


#ifndef ___Label_FWD_DEFINED__
#define ___Label_FWD_DEFINED__
typedef interface _Label _Label;

#endif 	/* ___Label_FWD_DEFINED__ */


/* header files for imported files */
#include "oaidl.h"
#include "ocidl.h"

#ifdef __cplusplus
extern "C"{
#endif 


#ifndef ___Label_INTERFACE_DEFINED__
#define ___Label_INTERFACE_DEFINED__

/* interface _Label */
/* [object][nonextensible][hidden][helpstring][uuid] */ 


EXTERN_C const IID IID__Label;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("33AD4ED9-6699-11CF-B70C-00AA0060D393")
    _Label : public IDispatch
    {
    public:
        virtual /* [helpstring] */ HRESULT __stdcall Missing1( void) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall Missing2( void) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall Missing3( void) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall Missing4( void) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall Missing5( void) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall Missing6( void) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall Missing7( void) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall Missing8( void) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall Missing9( void) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall Missing10( void) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall Missing11( void) = 0;
        
        virtual /* [helpstring][propget] */ HRESULT __stdcall get_Name( 
            /* [retval][out] */ BSTR *rhs) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall Missing12( void) = 0;
        
        virtual /* [helpstring][propget] */ HRESULT __stdcall get_Caption( 
            /* [retval][out] */ BSTR *rhs) = 0;
        
        virtual /* [helpstring][propput] */ HRESULT __stdcall put_Caption( 
            /* [in] */ BSTR rhs) = 0;
        
    };
    
    
#else 	/* C style interface */

    typedef struct _LabelVtbl
    {
        BEGIN_INTERFACE
        
        DECLSPEC_XFGVIRT(IUnknown, QueryInterface)
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            _Label * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        DECLSPEC_XFGVIRT(IUnknown, AddRef)
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            _Label * This);
        
        DECLSPEC_XFGVIRT(IUnknown, Release)
        ULONG ( STDMETHODCALLTYPE *Release )( 
            _Label * This);
        
        DECLSPEC_XFGVIRT(IDispatch, GetTypeInfoCount)
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            _Label * This,
            /* [out] */ UINT *pctinfo);
        
        DECLSPEC_XFGVIRT(IDispatch, GetTypeInfo)
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            _Label * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        DECLSPEC_XFGVIRT(IDispatch, GetIDsOfNames)
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            _Label * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        DECLSPEC_XFGVIRT(IDispatch, Invoke)
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            _Label * This,
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
        
        DECLSPEC_XFGVIRT(_Label, Missing1)
        /* [helpstring] */ HRESULT ( __stdcall *Missing1 )( 
            _Label * This);
        
        DECLSPEC_XFGVIRT(_Label, Missing2)
        /* [helpstring] */ HRESULT ( __stdcall *Missing2 )( 
            _Label * This);
        
        DECLSPEC_XFGVIRT(_Label, Missing3)
        /* [helpstring] */ HRESULT ( __stdcall *Missing3 )( 
            _Label * This);
        
        DECLSPEC_XFGVIRT(_Label, Missing4)
        /* [helpstring] */ HRESULT ( __stdcall *Missing4 )( 
            _Label * This);
        
        DECLSPEC_XFGVIRT(_Label, Missing5)
        /* [helpstring] */ HRESULT ( __stdcall *Missing5 )( 
            _Label * This);
        
        DECLSPEC_XFGVIRT(_Label, Missing6)
        /* [helpstring] */ HRESULT ( __stdcall *Missing6 )( 
            _Label * This);
        
        DECLSPEC_XFGVIRT(_Label, Missing7)
        /* [helpstring] */ HRESULT ( __stdcall *Missing7 )( 
            _Label * This);
        
        DECLSPEC_XFGVIRT(_Label, Missing8)
        /* [helpstring] */ HRESULT ( __stdcall *Missing8 )( 
            _Label * This);
        
        DECLSPEC_XFGVIRT(_Label, Missing9)
        /* [helpstring] */ HRESULT ( __stdcall *Missing9 )( 
            _Label * This);
        
        DECLSPEC_XFGVIRT(_Label, Missing10)
        /* [helpstring] */ HRESULT ( __stdcall *Missing10 )( 
            _Label * This);
        
        DECLSPEC_XFGVIRT(_Label, Missing11)
        /* [helpstring] */ HRESULT ( __stdcall *Missing11 )( 
            _Label * This);
        
        DECLSPEC_XFGVIRT(_Label, get_Name)
        /* [helpstring][propget] */ HRESULT ( __stdcall *get_Name )( 
            _Label * This,
            /* [retval][out] */ BSTR *rhs);
        
        DECLSPEC_XFGVIRT(_Label, Missing12)
        /* [helpstring] */ HRESULT ( __stdcall *Missing12 )( 
            _Label * This);
        
        DECLSPEC_XFGVIRT(_Label, get_Caption)
        /* [helpstring][propget] */ HRESULT ( __stdcall *get_Caption )( 
            _Label * This,
            /* [retval][out] */ BSTR *rhs);
        
        DECLSPEC_XFGVIRT(_Label, put_Caption)
        /* [helpstring][propput] */ HRESULT ( __stdcall *put_Caption )( 
            _Label * This,
            /* [in] */ BSTR rhs);
        
        END_INTERFACE
    } _LabelVtbl;

    interface _Label
    {
        CONST_VTBL struct _LabelVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define _Label_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define _Label_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define _Label_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define _Label_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define _Label_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define _Label_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define _Label_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define _Label_Missing1(This)	\
    ( (This)->lpVtbl -> Missing1(This) ) 

#define _Label_Missing2(This)	\
    ( (This)->lpVtbl -> Missing2(This) ) 

#define _Label_Missing3(This)	\
    ( (This)->lpVtbl -> Missing3(This) ) 

#define _Label_Missing4(This)	\
    ( (This)->lpVtbl -> Missing4(This) ) 

#define _Label_Missing5(This)	\
    ( (This)->lpVtbl -> Missing5(This) ) 

#define _Label_Missing6(This)	\
    ( (This)->lpVtbl -> Missing6(This) ) 

#define _Label_Missing7(This)	\
    ( (This)->lpVtbl -> Missing7(This) ) 

#define _Label_Missing8(This)	\
    ( (This)->lpVtbl -> Missing8(This) ) 

#define _Label_Missing9(This)	\
    ( (This)->lpVtbl -> Missing9(This) ) 

#define _Label_Missing10(This)	\
    ( (This)->lpVtbl -> Missing10(This) ) 

#define _Label_Missing11(This)	\
    ( (This)->lpVtbl -> Missing11(This) ) 

#define _Label_get_Name(This,rhs)	\
    ( (This)->lpVtbl -> get_Name(This,rhs) ) 

#define _Label_Missing12(This)	\
    ( (This)->lpVtbl -> Missing12(This) ) 

#define _Label_get_Caption(This,rhs)	\
    ( (This)->lpVtbl -> get_Caption(This,rhs) ) 

#define _Label_put_Caption(This,rhs)	\
    ( (This)->lpVtbl -> put_Caption(This,rhs) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* ___Label_INTERFACE_DEFINED__ */


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


