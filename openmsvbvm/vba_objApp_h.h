

/* this ALWAYS GENERATED file contains the definitions for the interfaces */


 /* File created by MIDL compiler version 8.01.0628 */
/* at Tue Jan 19 00:14:07 2038
 */
/* Compiler settings for vba_objApp.idl:
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

#ifndef __vba_objApp_h_h__
#define __vba_objApp_h_h__

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
        
        virtual /* [helpcontext][helpstring][propget] */ HRESULT __stdcall get_PrevInstance( 
            /* [retval][out] */ VARIANT_BOOL *rhs) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall Missing27( void) = 0;
        
        virtual /* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE get_StartMode( 
            /* [retval][out] */ short *rhs) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall Missing29( void) = 0;
        
        virtual /* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE get_TaskVisible( 
            /* [retval][out] */ VARIANT_BOOL *rhs) = 0;
        
        virtual /* [helpcontext][helpstring][propput] */ HRESULT STDMETHODCALLTYPE put_TaskVisible( 
            /* [in] */ VARIANT_BOOL rhs) = 0;
        
        virtual /* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE get_OleServerBusyTimeout( 
            /* [retval][out] */ long *rhs) = 0;
        
        virtual /* [helpcontext][helpstring][propput] */ HRESULT STDMETHODCALLTYPE put_OleServerBusyTimeout( 
            /* [in] */ long rhs) = 0;
        
        virtual /* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE get_OleServerBusyMsgTitle( 
            /* [retval][out] */ BSTR *rhs) = 0;
        
        virtual /* [helpcontext][helpstring][propput] */ HRESULT STDMETHODCALLTYPE put_OleServerBusyMsgTitle( 
            /* [in] */ BSTR rhs) = 0;
        
        virtual /* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE get_OleServerBusyMsgText( 
            /* [retval][out] */ BSTR *rhs) = 0;
        
        virtual /* [helpcontext][helpstring][propput] */ HRESULT STDMETHODCALLTYPE put_OleServerBusyMsgText( 
            /* [in] */ BSTR rhs) = 0;
        
        virtual /* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE get_OleServerBusyRaiseError( 
            /* [retval][out] */ VARIANT_BOOL *rhs) = 0;
        
        virtual /* [helpcontext][helpstring][propput] */ HRESULT STDMETHODCALLTYPE put_OleServerBusyRaiseError( 
            /* [in] */ VARIANT_BOOL rhs) = 0;
        
        virtual /* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE get_OleRequestPendingTimeout( 
            /* [retval][out] */ long *rhs) = 0;
        
        virtual /* [helpcontext][helpstring][propput] */ HRESULT STDMETHODCALLTYPE put_OleRequestPendingTimeout( 
            /* [in] */ long rhs) = 0;
        
        virtual /* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE get_OleRequestPendingMsgTitle( 
            /* [retval][out] */ BSTR *rhs) = 0;
        
        virtual /* [helpcontext][helpstring][propput] */ HRESULT STDMETHODCALLTYPE put_OleRequestPendingMsgTitle( 
            /* [in] */ BSTR rhs) = 0;
        
        virtual /* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE get_OleRequestPendingMsgText( 
            /* [retval][out] */ BSTR *rhs) = 0;
        
        virtual /* [helpcontext][helpstring][propput] */ HRESULT STDMETHODCALLTYPE put_OleRequestPendingMsgText( 
            /* [in] */ BSTR rhs) = 0;
        
        virtual /* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE get_Major( 
            /* [retval][out] */ short *rhs) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall Missing47( void) = 0;
        
        virtual /* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE get_Minor( 
            /* [retval][out] */ short *rhs) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall Missing49( void) = 0;
        
        virtual /* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE get_Revision( 
            /* [retval][out] */ short *rhs) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall Missing51( void) = 0;
        
        virtual /* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE get_Comments( 
            /* [retval][out] */ BSTR *rhs) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall Missing53( void) = 0;
        
        virtual /* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE get_CompanyName( 
            /* [retval][out] */ BSTR *rhs) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall Missing55( void) = 0;
        
        virtual /* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE get_FileDescription( 
            /* [retval][out] */ BSTR *rhs) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall Missing57( void) = 0;
        
        virtual /* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE get_LegalCopyright( 
            /* [retval][out] */ BSTR *rhs) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall Missing59( void) = 0;
        
        virtual /* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE get_LegalTrademarks( 
            /* [retval][out] */ BSTR *rhs) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall Missing61( void) = 0;
        
        virtual /* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE get_ProductName( 
            /* [retval][out] */ BSTR *rhs) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall Missing63( void) = 0;
        
        virtual /* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE get_hInstance( 
            /* [retval][out] */ long *rhs) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall Missing65( void) = 0;
        
        virtual /* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE get_NonModalAllowed( 
            /* [retval][out] */ VARIANT_BOOL *rhs) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall Missing67( void) = 0;
        
        virtual /* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE get_LogPath( 
            /* [retval][out] */ BSTR *rhs) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall Missing69( void) = 0;
        
        virtual /* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE get_LogMode( 
            /* [retval][out] */ long *rhs) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall Missing71( void) = 0;
        
        virtual /* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE get_UnattendedApp( 
            /* [retval][out] */ VARIANT_BOOL *rhs) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall Missing73( void) = 0;
        
        virtual /* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE get_ThreadID( 
            /* [retval][out] */ long *rhs) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall Missing75( void) = 0;
        
        virtual /* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE get_HelpFile( 
            /* [retval][out] */ BSTR *rhs) = 0;
        
        virtual /* [helpcontext][helpstring][propput] */ HRESULT STDMETHODCALLTYPE put_HelpFile( 
            /* [in] */ BSTR rhs) = 0;
        
        virtual /* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE get_RetainedProject( 
            /* [retval][out] */ VARIANT_BOOL *rhs) = 0;
        
        virtual /* [helpcontext][helpstring] */ HRESULT STDMETHODCALLTYPE StartLogging( 
            /* [in] */ BSTR LogTarget,
            /* [in] */ long LogModes) = 0;
        
        virtual /* [helpcontext][helpstring] */ HRESULT STDMETHODCALLTYPE LogEvent( 
            /* [in] */ BSTR LogBuffer,
            /* [in] */ VARIANT EventType) = 0;
        
    };
    
    
#else 	/* C style interface */

    typedef struct IAppVtbl
    {
        BEGIN_INTERFACE
        
        DECLSPEC_XFGVIRT(IUnknown, QueryInterface)
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IApp * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        DECLSPEC_XFGVIRT(IUnknown, AddRef)
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IApp * This);
        
        DECLSPEC_XFGVIRT(IUnknown, Release)
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IApp * This);
        
        DECLSPEC_XFGVIRT(IDispatch, GetTypeInfoCount)
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IApp * This,
            /* [out] */ UINT *pctinfo);
        
        DECLSPEC_XFGVIRT(IDispatch, GetTypeInfo)
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IApp * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        DECLSPEC_XFGVIRT(IDispatch, GetIDsOfNames)
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IApp * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        DECLSPEC_XFGVIRT(IDispatch, Invoke)
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
        
        DECLSPEC_XFGVIRT(IApp, HctlDemandLoad)
        /* [helpstring] */ HRESULT ( __stdcall *HctlDemandLoad )( 
            IApp * This,
            /* [out] */ unsigned int *ctl);
        
        DECLSPEC_XFGVIRT(IApp, ChkProp)
        /* [helpstring] */ HRESULT ( __stdcall *ChkProp )( 
            IApp * This,
            /* [in] */ unsigned int i,
            /* [out] */ unsigned int *tagData);
        
        DECLSPEC_XFGVIRT(IApp, SetPropA)
        /* [helpstring] */ HRESULT ( __stdcall *SetPropA )( 
            IApp * This,
            /* [in] */ unsigned int i,
            /* [out] */ unsigned int *tagData);
        
        DECLSPEC_XFGVIRT(IApp, GetPropA)
        /* [helpstring] */ HRESULT ( __stdcall *GetPropA )( 
            IApp * This,
            /* [in] */ unsigned int i,
            /* [out] */ unsigned int *tagData);
        
        DECLSPEC_XFGVIRT(IApp, GetPropHsz)
        /* [helpstring] */ HRESULT ( __stdcall *GetPropHsz )( 
            IApp * This,
            /* [in] */ unsigned int i,
            /* [out] */ unsigned char **hsz);
        
        DECLSPEC_XFGVIRT(IApp, LoadProp)
        /* [helpstring] */ HRESULT ( __stdcall *LoadProp )( 
            IApp * This,
            /* [in] */ unsigned int i,
            /* [out] */ unsigned int *fref);
        
        DECLSPEC_XFGVIRT(IApp, SaveProp)
        /* [helpstring] */ HRESULT ( __stdcall *SaveProp )( 
            IApp * This,
            /* [in] */ unsigned int i,
            /* [out] */ unsigned int *fref);
        
        DECLSPEC_XFGVIRT(IApp, GetPalette)
        /* [helpstring] */ HRESULT ( __stdcall *GetPalette )( 
            IApp * This);
        
        DECLSPEC_XFGVIRT(IApp, Reset)
        /* [helpstring] */ HRESULT ( __stdcall *Reset )( 
            IApp * This);
        
        DECLSPEC_XFGVIRT(IApp, get_DefaultProp)
        /* [helpstring] */ HRESULT ( __stdcall *get_DefaultProp )( 
            IApp * This,
            /* [out] */ VARIANT *var);
        
        DECLSPEC_XFGVIRT(IApp, put_DefaultProp)
        /* [helpstring] */ HRESULT ( __stdcall *put_DefaultProp )( 
            IApp * This,
            /* [out] */ VARIANT *var);
        
        DECLSPEC_XFGVIRT(IApp, get_000x)
        /* [helpstring] */ HRESULT ( __stdcall *get_000x )( 
            IApp * This,
            /* [out] */ VARIANT *var);
        
        DECLSPEC_XFGVIRT(IApp, put_000x)
        /* [helpstring] */ HRESULT ( __stdcall *put_000x )( 
            IApp * This,
            /* [in] */ unsigned int i);
        
        DECLSPEC_XFGVIRT(IApp, get_Path)
        /* [helpcontext][helpstring][propget] */ HRESULT ( __stdcall *get_Path )( 
            IApp * This,
            /* [retval][out] */ BSTR *rhs);
        
        DECLSPEC_XFGVIRT(IApp, put_Path)
        /* [helpcontext][helpstring][propput] */ HRESULT ( __stdcall *put_Path )( 
            IApp * This,
            /* [in] */ BSTR rhs);
        
        DECLSPEC_XFGVIRT(IApp, get_EXEName)
        /* [helpcontext][helpstring][propget] */ HRESULT ( __stdcall *get_EXEName )( 
            IApp * This,
            /* [retval][out] */ BSTR *rhs);
        
        DECLSPEC_XFGVIRT(IApp, put_EXEName)
        /* [helpcontext][helpstring][propput] */ HRESULT ( __stdcall *put_EXEName )( 
            IApp * This,
            /* [in] */ BSTR rhs);
        
        DECLSPEC_XFGVIRT(IApp, get_Title)
        /* [helpcontext][helpstring][propget] */ HRESULT ( __stdcall *get_Title )( 
            IApp * This,
            /* [retval][out] */ BSTR *rhs);
        
        DECLSPEC_XFGVIRT(IApp, put_Title)
        /* [helpcontext][helpstring][propput] */ HRESULT ( __stdcall *put_Title )( 
            IApp * This,
            /* [in] */ BSTR rhs);
        
        DECLSPEC_XFGVIRT(IApp, get_PrevInstance)
        /* [helpcontext][helpstring][propget] */ HRESULT ( __stdcall *get_PrevInstance )( 
            IApp * This,
            /* [retval][out] */ VARIANT_BOOL *rhs);
        
        DECLSPEC_XFGVIRT(IApp, Missing27)
        /* [helpstring] */ HRESULT ( __stdcall *Missing27 )( 
            IApp * This);
        
        DECLSPEC_XFGVIRT(IApp, get_StartMode)
        /* [helpcontext][helpstring][propget] */ HRESULT ( STDMETHODCALLTYPE *get_StartMode )( 
            IApp * This,
            /* [retval][out] */ short *rhs);
        
        DECLSPEC_XFGVIRT(IApp, Missing29)
        /* [helpstring] */ HRESULT ( __stdcall *Missing29 )( 
            IApp * This);
        
        DECLSPEC_XFGVIRT(IApp, get_TaskVisible)
        /* [helpcontext][helpstring][propget] */ HRESULT ( STDMETHODCALLTYPE *get_TaskVisible )( 
            IApp * This,
            /* [retval][out] */ VARIANT_BOOL *rhs);
        
        DECLSPEC_XFGVIRT(IApp, put_TaskVisible)
        /* [helpcontext][helpstring][propput] */ HRESULT ( STDMETHODCALLTYPE *put_TaskVisible )( 
            IApp * This,
            /* [in] */ VARIANT_BOOL rhs);
        
        DECLSPEC_XFGVIRT(IApp, get_OleServerBusyTimeout)
        /* [helpcontext][helpstring][propget] */ HRESULT ( STDMETHODCALLTYPE *get_OleServerBusyTimeout )( 
            IApp * This,
            /* [retval][out] */ long *rhs);
        
        DECLSPEC_XFGVIRT(IApp, put_OleServerBusyTimeout)
        /* [helpcontext][helpstring][propput] */ HRESULT ( STDMETHODCALLTYPE *put_OleServerBusyTimeout )( 
            IApp * This,
            /* [in] */ long rhs);
        
        DECLSPEC_XFGVIRT(IApp, get_OleServerBusyMsgTitle)
        /* [helpcontext][helpstring][propget] */ HRESULT ( STDMETHODCALLTYPE *get_OleServerBusyMsgTitle )( 
            IApp * This,
            /* [retval][out] */ BSTR *rhs);
        
        DECLSPEC_XFGVIRT(IApp, put_OleServerBusyMsgTitle)
        /* [helpcontext][helpstring][propput] */ HRESULT ( STDMETHODCALLTYPE *put_OleServerBusyMsgTitle )( 
            IApp * This,
            /* [in] */ BSTR rhs);
        
        DECLSPEC_XFGVIRT(IApp, get_OleServerBusyMsgText)
        /* [helpcontext][helpstring][propget] */ HRESULT ( STDMETHODCALLTYPE *get_OleServerBusyMsgText )( 
            IApp * This,
            /* [retval][out] */ BSTR *rhs);
        
        DECLSPEC_XFGVIRT(IApp, put_OleServerBusyMsgText)
        /* [helpcontext][helpstring][propput] */ HRESULT ( STDMETHODCALLTYPE *put_OleServerBusyMsgText )( 
            IApp * This,
            /* [in] */ BSTR rhs);
        
        DECLSPEC_XFGVIRT(IApp, get_OleServerBusyRaiseError)
        /* [helpcontext][helpstring][propget] */ HRESULT ( STDMETHODCALLTYPE *get_OleServerBusyRaiseError )( 
            IApp * This,
            /* [retval][out] */ VARIANT_BOOL *rhs);
        
        DECLSPEC_XFGVIRT(IApp, put_OleServerBusyRaiseError)
        /* [helpcontext][helpstring][propput] */ HRESULT ( STDMETHODCALLTYPE *put_OleServerBusyRaiseError )( 
            IApp * This,
            /* [in] */ VARIANT_BOOL rhs);
        
        DECLSPEC_XFGVIRT(IApp, get_OleRequestPendingTimeout)
        /* [helpcontext][helpstring][propget] */ HRESULT ( STDMETHODCALLTYPE *get_OleRequestPendingTimeout )( 
            IApp * This,
            /* [retval][out] */ long *rhs);
        
        DECLSPEC_XFGVIRT(IApp, put_OleRequestPendingTimeout)
        /* [helpcontext][helpstring][propput] */ HRESULT ( STDMETHODCALLTYPE *put_OleRequestPendingTimeout )( 
            IApp * This,
            /* [in] */ long rhs);
        
        DECLSPEC_XFGVIRT(IApp, get_OleRequestPendingMsgTitle)
        /* [helpcontext][helpstring][propget] */ HRESULT ( STDMETHODCALLTYPE *get_OleRequestPendingMsgTitle )( 
            IApp * This,
            /* [retval][out] */ BSTR *rhs);
        
        DECLSPEC_XFGVIRT(IApp, put_OleRequestPendingMsgTitle)
        /* [helpcontext][helpstring][propput] */ HRESULT ( STDMETHODCALLTYPE *put_OleRequestPendingMsgTitle )( 
            IApp * This,
            /* [in] */ BSTR rhs);
        
        DECLSPEC_XFGVIRT(IApp, get_OleRequestPendingMsgText)
        /* [helpcontext][helpstring][propget] */ HRESULT ( STDMETHODCALLTYPE *get_OleRequestPendingMsgText )( 
            IApp * This,
            /* [retval][out] */ BSTR *rhs);
        
        DECLSPEC_XFGVIRT(IApp, put_OleRequestPendingMsgText)
        /* [helpcontext][helpstring][propput] */ HRESULT ( STDMETHODCALLTYPE *put_OleRequestPendingMsgText )( 
            IApp * This,
            /* [in] */ BSTR rhs);
        
        DECLSPEC_XFGVIRT(IApp, get_Major)
        /* [helpcontext][helpstring][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Major )( 
            IApp * This,
            /* [retval][out] */ short *rhs);
        
        DECLSPEC_XFGVIRT(IApp, Missing47)
        /* [helpstring] */ HRESULT ( __stdcall *Missing47 )( 
            IApp * This);
        
        DECLSPEC_XFGVIRT(IApp, get_Minor)
        /* [helpcontext][helpstring][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Minor )( 
            IApp * This,
            /* [retval][out] */ short *rhs);
        
        DECLSPEC_XFGVIRT(IApp, Missing49)
        /* [helpstring] */ HRESULT ( __stdcall *Missing49 )( 
            IApp * This);
        
        DECLSPEC_XFGVIRT(IApp, get_Revision)
        /* [helpcontext][helpstring][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Revision )( 
            IApp * This,
            /* [retval][out] */ short *rhs);
        
        DECLSPEC_XFGVIRT(IApp, Missing51)
        /* [helpstring] */ HRESULT ( __stdcall *Missing51 )( 
            IApp * This);
        
        DECLSPEC_XFGVIRT(IApp, get_Comments)
        /* [helpcontext][helpstring][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Comments )( 
            IApp * This,
            /* [retval][out] */ BSTR *rhs);
        
        DECLSPEC_XFGVIRT(IApp, Missing53)
        /* [helpstring] */ HRESULT ( __stdcall *Missing53 )( 
            IApp * This);
        
        DECLSPEC_XFGVIRT(IApp, get_CompanyName)
        /* [helpcontext][helpstring][propget] */ HRESULT ( STDMETHODCALLTYPE *get_CompanyName )( 
            IApp * This,
            /* [retval][out] */ BSTR *rhs);
        
        DECLSPEC_XFGVIRT(IApp, Missing55)
        /* [helpstring] */ HRESULT ( __stdcall *Missing55 )( 
            IApp * This);
        
        DECLSPEC_XFGVIRT(IApp, get_FileDescription)
        /* [helpcontext][helpstring][propget] */ HRESULT ( STDMETHODCALLTYPE *get_FileDescription )( 
            IApp * This,
            /* [retval][out] */ BSTR *rhs);
        
        DECLSPEC_XFGVIRT(IApp, Missing57)
        /* [helpstring] */ HRESULT ( __stdcall *Missing57 )( 
            IApp * This);
        
        DECLSPEC_XFGVIRT(IApp, get_LegalCopyright)
        /* [helpcontext][helpstring][propget] */ HRESULT ( STDMETHODCALLTYPE *get_LegalCopyright )( 
            IApp * This,
            /* [retval][out] */ BSTR *rhs);
        
        DECLSPEC_XFGVIRT(IApp, Missing59)
        /* [helpstring] */ HRESULT ( __stdcall *Missing59 )( 
            IApp * This);
        
        DECLSPEC_XFGVIRT(IApp, get_LegalTrademarks)
        /* [helpcontext][helpstring][propget] */ HRESULT ( STDMETHODCALLTYPE *get_LegalTrademarks )( 
            IApp * This,
            /* [retval][out] */ BSTR *rhs);
        
        DECLSPEC_XFGVIRT(IApp, Missing61)
        /* [helpstring] */ HRESULT ( __stdcall *Missing61 )( 
            IApp * This);
        
        DECLSPEC_XFGVIRT(IApp, get_ProductName)
        /* [helpcontext][helpstring][propget] */ HRESULT ( STDMETHODCALLTYPE *get_ProductName )( 
            IApp * This,
            /* [retval][out] */ BSTR *rhs);
        
        DECLSPEC_XFGVIRT(IApp, Missing63)
        /* [helpstring] */ HRESULT ( __stdcall *Missing63 )( 
            IApp * This);
        
        DECLSPEC_XFGVIRT(IApp, get_hInstance)
        /* [helpcontext][helpstring][propget] */ HRESULT ( STDMETHODCALLTYPE *get_hInstance )( 
            IApp * This,
            /* [retval][out] */ long *rhs);
        
        DECLSPEC_XFGVIRT(IApp, Missing65)
        /* [helpstring] */ HRESULT ( __stdcall *Missing65 )( 
            IApp * This);
        
        DECLSPEC_XFGVIRT(IApp, get_NonModalAllowed)
        /* [helpcontext][helpstring][propget] */ HRESULT ( STDMETHODCALLTYPE *get_NonModalAllowed )( 
            IApp * This,
            /* [retval][out] */ VARIANT_BOOL *rhs);
        
        DECLSPEC_XFGVIRT(IApp, Missing67)
        /* [helpstring] */ HRESULT ( __stdcall *Missing67 )( 
            IApp * This);
        
        DECLSPEC_XFGVIRT(IApp, get_LogPath)
        /* [helpcontext][helpstring][propget] */ HRESULT ( STDMETHODCALLTYPE *get_LogPath )( 
            IApp * This,
            /* [retval][out] */ BSTR *rhs);
        
        DECLSPEC_XFGVIRT(IApp, Missing69)
        /* [helpstring] */ HRESULT ( __stdcall *Missing69 )( 
            IApp * This);
        
        DECLSPEC_XFGVIRT(IApp, get_LogMode)
        /* [helpcontext][helpstring][propget] */ HRESULT ( STDMETHODCALLTYPE *get_LogMode )( 
            IApp * This,
            /* [retval][out] */ long *rhs);
        
        DECLSPEC_XFGVIRT(IApp, Missing71)
        /* [helpstring] */ HRESULT ( __stdcall *Missing71 )( 
            IApp * This);
        
        DECLSPEC_XFGVIRT(IApp, get_UnattendedApp)
        /* [helpcontext][helpstring][propget] */ HRESULT ( STDMETHODCALLTYPE *get_UnattendedApp )( 
            IApp * This,
            /* [retval][out] */ VARIANT_BOOL *rhs);
        
        DECLSPEC_XFGVIRT(IApp, Missing73)
        /* [helpstring] */ HRESULT ( __stdcall *Missing73 )( 
            IApp * This);
        
        DECLSPEC_XFGVIRT(IApp, get_ThreadID)
        /* [helpcontext][helpstring][propget] */ HRESULT ( STDMETHODCALLTYPE *get_ThreadID )( 
            IApp * This,
            /* [retval][out] */ long *rhs);
        
        DECLSPEC_XFGVIRT(IApp, Missing75)
        /* [helpstring] */ HRESULT ( __stdcall *Missing75 )( 
            IApp * This);
        
        DECLSPEC_XFGVIRT(IApp, get_HelpFile)
        /* [helpcontext][helpstring][propget] */ HRESULT ( STDMETHODCALLTYPE *get_HelpFile )( 
            IApp * This,
            /* [retval][out] */ BSTR *rhs);
        
        DECLSPEC_XFGVIRT(IApp, put_HelpFile)
        /* [helpcontext][helpstring][propput] */ HRESULT ( STDMETHODCALLTYPE *put_HelpFile )( 
            IApp * This,
            /* [in] */ BSTR rhs);
        
        DECLSPEC_XFGVIRT(IApp, get_RetainedProject)
        /* [helpcontext][helpstring][propget] */ HRESULT ( STDMETHODCALLTYPE *get_RetainedProject )( 
            IApp * This,
            /* [retval][out] */ VARIANT_BOOL *rhs);
        
        DECLSPEC_XFGVIRT(IApp, StartLogging)
        /* [helpcontext][helpstring] */ HRESULT ( STDMETHODCALLTYPE *StartLogging )( 
            IApp * This,
            /* [in] */ BSTR LogTarget,
            /* [in] */ long LogModes);
        
        DECLSPEC_XFGVIRT(IApp, LogEvent)
        /* [helpcontext][helpstring] */ HRESULT ( STDMETHODCALLTYPE *LogEvent )( 
            IApp * This,
            /* [in] */ BSTR LogBuffer,
            /* [in] */ VARIANT EventType);
        
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

#define IApp_get_PrevInstance(This,rhs)	\
    ( (This)->lpVtbl -> get_PrevInstance(This,rhs) ) 

#define IApp_Missing27(This)	\
    ( (This)->lpVtbl -> Missing27(This) ) 

#define IApp_get_StartMode(This,rhs)	\
    ( (This)->lpVtbl -> get_StartMode(This,rhs) ) 

#define IApp_Missing29(This)	\
    ( (This)->lpVtbl -> Missing29(This) ) 

#define IApp_get_TaskVisible(This,rhs)	\
    ( (This)->lpVtbl -> get_TaskVisible(This,rhs) ) 

#define IApp_put_TaskVisible(This,rhs)	\
    ( (This)->lpVtbl -> put_TaskVisible(This,rhs) ) 

#define IApp_get_OleServerBusyTimeout(This,rhs)	\
    ( (This)->lpVtbl -> get_OleServerBusyTimeout(This,rhs) ) 

#define IApp_put_OleServerBusyTimeout(This,rhs)	\
    ( (This)->lpVtbl -> put_OleServerBusyTimeout(This,rhs) ) 

#define IApp_get_OleServerBusyMsgTitle(This,rhs)	\
    ( (This)->lpVtbl -> get_OleServerBusyMsgTitle(This,rhs) ) 

#define IApp_put_OleServerBusyMsgTitle(This,rhs)	\
    ( (This)->lpVtbl -> put_OleServerBusyMsgTitle(This,rhs) ) 

#define IApp_get_OleServerBusyMsgText(This,rhs)	\
    ( (This)->lpVtbl -> get_OleServerBusyMsgText(This,rhs) ) 

#define IApp_put_OleServerBusyMsgText(This,rhs)	\
    ( (This)->lpVtbl -> put_OleServerBusyMsgText(This,rhs) ) 

#define IApp_get_OleServerBusyRaiseError(This,rhs)	\
    ( (This)->lpVtbl -> get_OleServerBusyRaiseError(This,rhs) ) 

#define IApp_put_OleServerBusyRaiseError(This,rhs)	\
    ( (This)->lpVtbl -> put_OleServerBusyRaiseError(This,rhs) ) 

#define IApp_get_OleRequestPendingTimeout(This,rhs)	\
    ( (This)->lpVtbl -> get_OleRequestPendingTimeout(This,rhs) ) 

#define IApp_put_OleRequestPendingTimeout(This,rhs)	\
    ( (This)->lpVtbl -> put_OleRequestPendingTimeout(This,rhs) ) 

#define IApp_get_OleRequestPendingMsgTitle(This,rhs)	\
    ( (This)->lpVtbl -> get_OleRequestPendingMsgTitle(This,rhs) ) 

#define IApp_put_OleRequestPendingMsgTitle(This,rhs)	\
    ( (This)->lpVtbl -> put_OleRequestPendingMsgTitle(This,rhs) ) 

#define IApp_get_OleRequestPendingMsgText(This,rhs)	\
    ( (This)->lpVtbl -> get_OleRequestPendingMsgText(This,rhs) ) 

#define IApp_put_OleRequestPendingMsgText(This,rhs)	\
    ( (This)->lpVtbl -> put_OleRequestPendingMsgText(This,rhs) ) 

#define IApp_get_Major(This,rhs)	\
    ( (This)->lpVtbl -> get_Major(This,rhs) ) 

#define IApp_Missing47(This)	\
    ( (This)->lpVtbl -> Missing47(This) ) 

#define IApp_get_Minor(This,rhs)	\
    ( (This)->lpVtbl -> get_Minor(This,rhs) ) 

#define IApp_Missing49(This)	\
    ( (This)->lpVtbl -> Missing49(This) ) 

#define IApp_get_Revision(This,rhs)	\
    ( (This)->lpVtbl -> get_Revision(This,rhs) ) 

#define IApp_Missing51(This)	\
    ( (This)->lpVtbl -> Missing51(This) ) 

#define IApp_get_Comments(This,rhs)	\
    ( (This)->lpVtbl -> get_Comments(This,rhs) ) 

#define IApp_Missing53(This)	\
    ( (This)->lpVtbl -> Missing53(This) ) 

#define IApp_get_CompanyName(This,rhs)	\
    ( (This)->lpVtbl -> get_CompanyName(This,rhs) ) 

#define IApp_Missing55(This)	\
    ( (This)->lpVtbl -> Missing55(This) ) 

#define IApp_get_FileDescription(This,rhs)	\
    ( (This)->lpVtbl -> get_FileDescription(This,rhs) ) 

#define IApp_Missing57(This)	\
    ( (This)->lpVtbl -> Missing57(This) ) 

#define IApp_get_LegalCopyright(This,rhs)	\
    ( (This)->lpVtbl -> get_LegalCopyright(This,rhs) ) 

#define IApp_Missing59(This)	\
    ( (This)->lpVtbl -> Missing59(This) ) 

#define IApp_get_LegalTrademarks(This,rhs)	\
    ( (This)->lpVtbl -> get_LegalTrademarks(This,rhs) ) 

#define IApp_Missing61(This)	\
    ( (This)->lpVtbl -> Missing61(This) ) 

#define IApp_get_ProductName(This,rhs)	\
    ( (This)->lpVtbl -> get_ProductName(This,rhs) ) 

#define IApp_Missing63(This)	\
    ( (This)->lpVtbl -> Missing63(This) ) 

#define IApp_get_hInstance(This,rhs)	\
    ( (This)->lpVtbl -> get_hInstance(This,rhs) ) 

#define IApp_Missing65(This)	\
    ( (This)->lpVtbl -> Missing65(This) ) 

#define IApp_get_NonModalAllowed(This,rhs)	\
    ( (This)->lpVtbl -> get_NonModalAllowed(This,rhs) ) 

#define IApp_Missing67(This)	\
    ( (This)->lpVtbl -> Missing67(This) ) 

#define IApp_get_LogPath(This,rhs)	\
    ( (This)->lpVtbl -> get_LogPath(This,rhs) ) 

#define IApp_Missing69(This)	\
    ( (This)->lpVtbl -> Missing69(This) ) 

#define IApp_get_LogMode(This,rhs)	\
    ( (This)->lpVtbl -> get_LogMode(This,rhs) ) 

#define IApp_Missing71(This)	\
    ( (This)->lpVtbl -> Missing71(This) ) 

#define IApp_get_UnattendedApp(This,rhs)	\
    ( (This)->lpVtbl -> get_UnattendedApp(This,rhs) ) 

#define IApp_Missing73(This)	\
    ( (This)->lpVtbl -> Missing73(This) ) 

#define IApp_get_ThreadID(This,rhs)	\
    ( (This)->lpVtbl -> get_ThreadID(This,rhs) ) 

#define IApp_Missing75(This)	\
    ( (This)->lpVtbl -> Missing75(This) ) 

#define IApp_get_HelpFile(This,rhs)	\
    ( (This)->lpVtbl -> get_HelpFile(This,rhs) ) 

#define IApp_put_HelpFile(This,rhs)	\
    ( (This)->lpVtbl -> put_HelpFile(This,rhs) ) 

#define IApp_get_RetainedProject(This,rhs)	\
    ( (This)->lpVtbl -> get_RetainedProject(This,rhs) ) 

#define IApp_StartLogging(This,LogTarget,LogModes)	\
    ( (This)->lpVtbl -> StartLogging(This,LogTarget,LogModes) ) 

#define IApp_LogEvent(This,LogBuffer,EventType)	\
    ( (This)->lpVtbl -> LogEvent(This,LogBuffer,EventType) ) 

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


