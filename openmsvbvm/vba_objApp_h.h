

/* this ALWAYS GENERATED file contains the definitions for the interfaces */


 /* File created by MIDL compiler version 8.01.0622 */
/* at Tue Jan 19 00:14:07 2038
 */
/* Compiler settings for vba_objApp.idl:
    Oicf, W1, Zp8, env=Win32 (32b run), target_arch=X86 8.01.0622 
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

#ifndef __vba_objApp_h_h__
#define __vba_objApp_h_h__

#if defined(_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

/* Forward Declarations */ 

#ifndef __IApp_FWD_DEFINED__
#define __IApp_FWD_DEFINED__
typedef interface IApp IApp;

#endif 	/* __IApp_FWD_DEFINED__ */


/* header files for imported files */
#include "oaidl.h"
#include "ocidl.h"

#ifdef __cplusplus
extern "C"{
#endif 


#ifndef __IApp_INTERFACE_DEFINED__
#define __IApp_INTERFACE_DEFINED__

/* interface IApp */
/* [object][nonextensible][hidden][helpcontext][helpstring][uuid] */ 


EXTERN_C const IID IID_IApp;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("33ad4f79-6699-11cf-b70c-00aa0060d393")
    IApp : public IDispatch
    {
    public:
        virtual /* [helpstring] */ HRESULT __stdcall HctlDemandLoad( 
            /* [out] */ unsigned int *ctl) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall ChkProp( 
            /* [in] */ unsigned int i,
            /* [out] */ unsigned int *tagData) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall SetPropA( 
            /* [in] */ unsigned int i,
            /* [out] */ unsigned int *tagData) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall GetPropA( 
            /* [in] */ unsigned int i,
            /* [out] */ unsigned int *tagData) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall GetPropHsz( 
            /* [in] */ unsigned int i,
            /* [out] */ unsigned char **hsz) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall LoadProp( 
            /* [in] */ unsigned int i,
            /* [out] */ unsigned int *fref) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall SaveProp( 
            /* [in] */ unsigned int i,
            /* [out] */ unsigned int *fref) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall GetPalette( void) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall Reset( void) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall get_DefaultProp( 
            /* [out] */ VARIANT *var) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall put_DefaultProp( 
            /* [out] */ VARIANT *var) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall get_000x( 
            /* [out] */ VARIANT *var) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall put_000x( 
            /* [in] */ unsigned int i) = 0;
        
        virtual /* [helpcontext][helpstring][propget] */ HRESULT __stdcall get_Path( 
            /* [retval][out] */ BSTR *rhs) = 0;
        
        virtual /* [helpcontext][helpstring][propput] */ HRESULT __stdcall put_Path( 
            /* [in] */ BSTR rhs) = 0;
        
        virtual /* [helpcontext][helpstring][propget] */ HRESULT __stdcall get_EXEName( 
            /* [retval][out] */ BSTR *rhs) = 0;
        
        virtual /* [helpcontext][helpstring][propput] */ HRESULT __stdcall put_EXEName( 
            /* [in] */ BSTR rhs) = 0;
        
        virtual /* [helpcontext][helpstring][propget] */ HRESULT __stdcall get_Title( 
            /* [retval][out] */ BSTR *rhs) = 0;
        
        virtual /* [helpcontext][helpstring][propput] */ HRESULT __stdcall put_Title( 
            /* [in] */ BSTR rhs) = 0;
        
    };
    
    
#else 	/* C style interface */

    typedef struct IAppVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IApp * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IApp * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IApp * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IApp * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IApp * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IApp * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IApp * This,
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
        
        /* [helpstring] */ HRESULT ( __stdcall *HctlDemandLoad )( 
            IApp * This,
            /* [out] */ unsigned int *ctl);
        
        /* [helpstring] */ HRESULT ( __stdcall *ChkProp )( 
            IApp * This,
            /* [in] */ unsigned int i,
            /* [out] */ unsigned int *tagData);
        
        /* [helpstring] */ HRESULT ( __stdcall *SetPropA )( 
            IApp * This,
            /* [in] */ unsigned int i,
            /* [out] */ unsigned int *tagData);
        
        /* [helpstring] */ HRESULT ( __stdcall *GetPropA )( 
            IApp * This,
            /* [in] */ unsigned int i,
            /* [out] */ unsigned int *tagData);
        
        /* [helpstring] */ HRESULT ( __stdcall *GetPropHsz )( 
            IApp * This,
            /* [in] */ unsigned int i,
            /* [out] */ unsigned char **hsz);
        
        /* [helpstring] */ HRESULT ( __stdcall *LoadProp )( 
            IApp * This,
            /* [in] */ unsigned int i,
            /* [out] */ unsigned int *fref);
        
        /* [helpstring] */ HRESULT ( __stdcall *SaveProp )( 
            IApp * This,
            /* [in] */ unsigned int i,
            /* [out] */ unsigned int *fref);
        
        /* [helpstring] */ HRESULT ( __stdcall *GetPalette )( 
            IApp * This);
        
        /* [helpstring] */ HRESULT ( __stdcall *Reset )( 
            IApp * This);
        
        /* [helpstring] */ HRESULT ( __stdcall *get_DefaultProp )( 
            IApp * This,
            /* [out] */ VARIANT *var);
        
        /* [helpstring] */ HRESULT ( __stdcall *put_DefaultProp )( 
            IApp * This,
            /* [out] */ VARIANT *var);
        
        /* [helpstring] */ HRESULT ( __stdcall *get_000x )( 
            IApp * This,
            /* [out] */ VARIANT *var);
        
        /* [helpstring] */ HRESULT ( __stdcall *put_000x )( 
            IApp * This,
            /* [in] */ unsigned int i);
        
        /* [helpcontext][helpstring][propget] */ HRESULT ( __stdcall *get_Path )( 
            IApp * This,
            /* [retval][out] */ BSTR *rhs);
        
        /* [helpcontext][helpstring][propput] */ HRESULT ( __stdcall *put_Path )( 
            IApp * This,
            /* [in] */ BSTR rhs);
        
        /* [helpcontext][helpstring][propget] */ HRESULT ( __stdcall *get_EXEName )( 
            IApp * This,
            /* [retval][out] */ BSTR *rhs);
        
        /* [helpcontext][helpstring][propput] */ HRESULT ( __stdcall *put_EXEName )( 
            IApp * This,
            /* [in] */ BSTR rhs);
        
        /* [helpcontext][helpstring][propget] */ HRESULT ( __stdcall *get_Title )( 
            IApp * This,
            /* [retval][out] */ BSTR *rhs);
        
        /* [helpcontext][helpstring][propput] */ HRESULT ( __stdcall *put_Title )( 
            IApp * This,
            /* [in] */ BSTR rhs);
        
        END_INTERFACE
    } IAppVtbl;

    interface IApp
    {
        CONST_VTBL struct IAppVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IApp_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IApp_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IApp_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IApp_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IApp_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IApp_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IApp_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define IApp_HctlDemandLoad(This,ctl)	\
    ( (This)->lpVtbl -> HctlDemandLoad(This,ctl) ) 

#define IApp_ChkProp(This,i,tagData)	\
    ( (This)->lpVtbl -> ChkProp(This,i,tagData) ) 

#define IApp_SetPropA(This,i,tagData)	\
    ( (This)->lpVtbl -> SetPropA(This,i,tagData) ) 

#define IApp_GetPropA(This,i,tagData)	\
    ( (This)->lpVtbl -> GetPropA(This,i,tagData) ) 

#define IApp_GetPropHsz(This,i,hsz)	\
    ( (This)->lpVtbl -> GetPropHsz(This,i,hsz) ) 

#define IApp_LoadProp(This,i,fref)	\
    ( (This)->lpVtbl -> LoadProp(This,i,fref) ) 

#define IApp_SaveProp(This,i,fref)	\
    ( (This)->lpVtbl -> SaveProp(This,i,fref) ) 

#define IApp_GetPalette(This)	\
    ( (This)->lpVtbl -> GetPalette(This) ) 

#define IApp_Reset(This)	\
    ( (This)->lpVtbl -> Reset(This) ) 

#define IApp_get_DefaultProp(This,var)	\
    ( (This)->lpVtbl -> get_DefaultProp(This,var) ) 

#define IApp_put_DefaultProp(This,var)	\
    ( (This)->lpVtbl -> put_DefaultProp(This,var) ) 

#define IApp_get_000x(This,var)	\
    ( (This)->lpVtbl -> get_000x(This,var) ) 

#define IApp_put_000x(This,i)	\
    ( (This)->lpVtbl -> put_000x(This,i) ) 

#define IApp_get_Path(This,rhs)	\
    ( (This)->lpVtbl -> get_Path(This,rhs) ) 

#define IApp_put_Path(This,rhs)	\
    ( (This)->lpVtbl -> put_Path(This,rhs) ) 

#define IApp_get_EXEName(This,rhs)	\
    ( (This)->lpVtbl -> get_EXEName(This,rhs) ) 

#define IApp_put_EXEName(This,rhs)	\
    ( (This)->lpVtbl -> put_EXEName(This,rhs) ) 

#define IApp_get_Title(This,rhs)	\
    ( (This)->lpVtbl -> get_Title(This,rhs) ) 

#define IApp_put_Title(This,rhs)	\
    ( (This)->lpVtbl -> put_Title(This,rhs) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __IApp_INTERFACE_DEFINED__ */


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


