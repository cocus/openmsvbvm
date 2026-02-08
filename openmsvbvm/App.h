

/* this ALWAYS GENERATED file contains the definitions for the interfaces */


 /* File created by MIDL compiler version 8.01.0628 */
/* at Tue Jan 19 00:14:07 2038
 */
/* Compiler settings for App.idl:
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

#ifndef __App_h__
#define __App_h__

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

#ifndef __App_FWD_DEFINED__
#define __App_FWD_DEFINED__

#ifdef __cplusplus
typedef class App App;
#else
typedef struct App App;
#endif /* __cplusplus */

#endif 	/* __App_FWD_DEFINED__ */


#ifndef ___App_FWD_DEFINED__
#define ___App_FWD_DEFINED__
typedef interface _App _App;

#endif 	/* ___App_FWD_DEFINED__ */


/* header files for imported files */
#include "oaidl.h"
#include "ocidl.h"

#ifdef __cplusplus
extern "C"{
#endif 


#ifndef ___App_INTERFACE_DEFINED__
#define ___App_INTERFACE_DEFINED__

/* interface _App */
/* [object][nonextensible][hidden][helpcontext][helpstring][uuid] */ 


EXTERN_C const IID IID__App;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("33AD4F79-6699-11CF-B70C-00AA0060D393")
    _App : public IDispatch
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

    typedef struct _AppVtbl
    {
        BEGIN_INTERFACE
        
        DECLSPEC_XFGVIRT(IUnknown, QueryInterface)
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            _App * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        DECLSPEC_XFGVIRT(IUnknown, AddRef)
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            _App * This);
        
        DECLSPEC_XFGVIRT(IUnknown, Release)
        ULONG ( STDMETHODCALLTYPE *Release )( 
            _App * This);
        
        DECLSPEC_XFGVIRT(IDispatch, GetTypeInfoCount)
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            _App * This,
            /* [out] */ UINT *pctinfo);
        
        DECLSPEC_XFGVIRT(IDispatch, GetTypeInfo)
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            _App * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        DECLSPEC_XFGVIRT(IDispatch, GetIDsOfNames)
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            _App * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        DECLSPEC_XFGVIRT(IDispatch, Invoke)
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            _App * This,
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
        
        DECLSPEC_XFGVIRT(_App, HctlDemandLoad)
        /* [helpstring] */ HRESULT ( __stdcall *HctlDemandLoad )( 
            _App * This,
            /* [out] */ unsigned int *ctl);
        
        DECLSPEC_XFGVIRT(_App, ChkProp)
        /* [helpstring] */ HRESULT ( __stdcall *ChkProp )( 
            _App * This,
            /* [in] */ unsigned int i,
            /* [out] */ unsigned int *tagData);
        
        DECLSPEC_XFGVIRT(_App, SetPropA)
        /* [helpstring] */ HRESULT ( __stdcall *SetPropA )( 
            _App * This,
            /* [in] */ unsigned int i,
            /* [out] */ unsigned int *tagData);
        
        DECLSPEC_XFGVIRT(_App, GetPropA)
        /* [helpstring] */ HRESULT ( __stdcall *GetPropA )( 
            _App * This,
            /* [in] */ unsigned int i,
            /* [out] */ unsigned int *tagData);
        
        DECLSPEC_XFGVIRT(_App, GetPropHsz)
        /* [helpstring] */ HRESULT ( __stdcall *GetPropHsz )( 
            _App * This,
            /* [in] */ unsigned int i,
            /* [out] */ unsigned char **hsz);
        
        DECLSPEC_XFGVIRT(_App, LoadProp)
        /* [helpstring] */ HRESULT ( __stdcall *LoadProp )( 
            _App * This,
            /* [in] */ unsigned int i,
            /* [out] */ unsigned int *fref);
        
        DECLSPEC_XFGVIRT(_App, SaveProp)
        /* [helpstring] */ HRESULT ( __stdcall *SaveProp )( 
            _App * This,
            /* [in] */ unsigned int i,
            /* [out] */ unsigned int *fref);
        
        DECLSPEC_XFGVIRT(_App, GetPalette)
        /* [helpstring] */ HRESULT ( __stdcall *GetPalette )( 
            _App * This);
        
        DECLSPEC_XFGVIRT(_App, Reset)
        /* [helpstring] */ HRESULT ( __stdcall *Reset )( 
            _App * This);
        
        DECLSPEC_XFGVIRT(_App, get_DefaultProp)
        /* [helpstring] */ HRESULT ( __stdcall *get_DefaultProp )( 
            _App * This,
            /* [out] */ VARIANT *var);
        
        DECLSPEC_XFGVIRT(_App, put_DefaultProp)
        /* [helpstring] */ HRESULT ( __stdcall *put_DefaultProp )( 
            _App * This,
            /* [out] */ VARIANT *var);
        
        DECLSPEC_XFGVIRT(_App, get_000x)
        /* [helpstring] */ HRESULT ( __stdcall *get_000x )( 
            _App * This,
            /* [out] */ VARIANT *var);
        
        DECLSPEC_XFGVIRT(_App, put_000x)
        /* [helpstring] */ HRESULT ( __stdcall *put_000x )( 
            _App * This,
            /* [in] */ unsigned int i);
        
        DECLSPEC_XFGVIRT(_App, get_Path)
        /* [helpcontext][helpstring][propget] */ HRESULT ( __stdcall *get_Path )( 
            _App * This,
            /* [retval][out] */ BSTR *rhs);
        
        DECLSPEC_XFGVIRT(_App, put_Path)
        /* [helpcontext][helpstring][propput] */ HRESULT ( __stdcall *put_Path )( 
            _App * This,
            /* [in] */ BSTR rhs);
        
        DECLSPEC_XFGVIRT(_App, get_EXEName)
        /* [helpcontext][helpstring][propget] */ HRESULT ( __stdcall *get_EXEName )( 
            _App * This,
            /* [retval][out] */ BSTR *rhs);
        
        DECLSPEC_XFGVIRT(_App, put_EXEName)
        /* [helpcontext][helpstring][propput] */ HRESULT ( __stdcall *put_EXEName )( 
            _App * This,
            /* [in] */ BSTR rhs);
        
        DECLSPEC_XFGVIRT(_App, get_Title)
        /* [helpcontext][helpstring][propget] */ HRESULT ( __stdcall *get_Title )( 
            _App * This,
            /* [retval][out] */ BSTR *rhs);
        
        DECLSPEC_XFGVIRT(_App, put_Title)
        /* [helpcontext][helpstring][propput] */ HRESULT ( __stdcall *put_Title )( 
            _App * This,
            /* [in] */ BSTR rhs);
        
        DECLSPEC_XFGVIRT(_App, get_PrevInstance)
        /* [helpcontext][helpstring][propget] */ HRESULT ( __stdcall *get_PrevInstance )( 
            _App * This,
            /* [retval][out] */ VARIANT_BOOL *rhs);
        
        DECLSPEC_XFGVIRT(_App, Missing27)
        /* [helpstring] */ HRESULT ( __stdcall *Missing27 )( 
            _App * This);
        
        DECLSPEC_XFGVIRT(_App, get_StartMode)
        /* [helpcontext][helpstring][propget] */ HRESULT ( STDMETHODCALLTYPE *get_StartMode )( 
            _App * This,
            /* [retval][out] */ short *rhs);
        
        DECLSPEC_XFGVIRT(_App, Missing29)
        /* [helpstring] */ HRESULT ( __stdcall *Missing29 )( 
            _App * This);
        
        DECLSPEC_XFGVIRT(_App, get_TaskVisible)
        /* [helpcontext][helpstring][propget] */ HRESULT ( STDMETHODCALLTYPE *get_TaskVisible )( 
            _App * This,
            /* [retval][out] */ VARIANT_BOOL *rhs);
        
        DECLSPEC_XFGVIRT(_App, put_TaskVisible)
        /* [helpcontext][helpstring][propput] */ HRESULT ( STDMETHODCALLTYPE *put_TaskVisible )( 
            _App * This,
            /* [in] */ VARIANT_BOOL rhs);
        
        DECLSPEC_XFGVIRT(_App, get_OleServerBusyTimeout)
        /* [helpcontext][helpstring][propget] */ HRESULT ( STDMETHODCALLTYPE *get_OleServerBusyTimeout )( 
            _App * This,
            /* [retval][out] */ long *rhs);
        
        DECLSPEC_XFGVIRT(_App, put_OleServerBusyTimeout)
        /* [helpcontext][helpstring][propput] */ HRESULT ( STDMETHODCALLTYPE *put_OleServerBusyTimeout )( 
            _App * This,
            /* [in] */ long rhs);
        
        DECLSPEC_XFGVIRT(_App, get_OleServerBusyMsgTitle)
        /* [helpcontext][helpstring][propget] */ HRESULT ( STDMETHODCALLTYPE *get_OleServerBusyMsgTitle )( 
            _App * This,
            /* [retval][out] */ BSTR *rhs);
        
        DECLSPEC_XFGVIRT(_App, put_OleServerBusyMsgTitle)
        /* [helpcontext][helpstring][propput] */ HRESULT ( STDMETHODCALLTYPE *put_OleServerBusyMsgTitle )( 
            _App * This,
            /* [in] */ BSTR rhs);
        
        DECLSPEC_XFGVIRT(_App, get_OleServerBusyMsgText)
        /* [helpcontext][helpstring][propget] */ HRESULT ( STDMETHODCALLTYPE *get_OleServerBusyMsgText )( 
            _App * This,
            /* [retval][out] */ BSTR *rhs);
        
        DECLSPEC_XFGVIRT(_App, put_OleServerBusyMsgText)
        /* [helpcontext][helpstring][propput] */ HRESULT ( STDMETHODCALLTYPE *put_OleServerBusyMsgText )( 
            _App * This,
            /* [in] */ BSTR rhs);
        
        DECLSPEC_XFGVIRT(_App, get_OleServerBusyRaiseError)
        /* [helpcontext][helpstring][propget] */ HRESULT ( STDMETHODCALLTYPE *get_OleServerBusyRaiseError )( 
            _App * This,
            /* [retval][out] */ VARIANT_BOOL *rhs);
        
        DECLSPEC_XFGVIRT(_App, put_OleServerBusyRaiseError)
        /* [helpcontext][helpstring][propput] */ HRESULT ( STDMETHODCALLTYPE *put_OleServerBusyRaiseError )( 
            _App * This,
            /* [in] */ VARIANT_BOOL rhs);
        
        DECLSPEC_XFGVIRT(_App, get_OleRequestPendingTimeout)
        /* [helpcontext][helpstring][propget] */ HRESULT ( STDMETHODCALLTYPE *get_OleRequestPendingTimeout )( 
            _App * This,
            /* [retval][out] */ long *rhs);
        
        DECLSPEC_XFGVIRT(_App, put_OleRequestPendingTimeout)
        /* [helpcontext][helpstring][propput] */ HRESULT ( STDMETHODCALLTYPE *put_OleRequestPendingTimeout )( 
            _App * This,
            /* [in] */ long rhs);
        
        DECLSPEC_XFGVIRT(_App, get_OleRequestPendingMsgTitle)
        /* [helpcontext][helpstring][propget] */ HRESULT ( STDMETHODCALLTYPE *get_OleRequestPendingMsgTitle )( 
            _App * This,
            /* [retval][out] */ BSTR *rhs);
        
        DECLSPEC_XFGVIRT(_App, put_OleRequestPendingMsgTitle)
        /* [helpcontext][helpstring][propput] */ HRESULT ( STDMETHODCALLTYPE *put_OleRequestPendingMsgTitle )( 
            _App * This,
            /* [in] */ BSTR rhs);
        
        DECLSPEC_XFGVIRT(_App, get_OleRequestPendingMsgText)
        /* [helpcontext][helpstring][propget] */ HRESULT ( STDMETHODCALLTYPE *get_OleRequestPendingMsgText )( 
            _App * This,
            /* [retval][out] */ BSTR *rhs);
        
        DECLSPEC_XFGVIRT(_App, put_OleRequestPendingMsgText)
        /* [helpcontext][helpstring][propput] */ HRESULT ( STDMETHODCALLTYPE *put_OleRequestPendingMsgText )( 
            _App * This,
            /* [in] */ BSTR rhs);
        
        DECLSPEC_XFGVIRT(_App, get_Major)
        /* [helpcontext][helpstring][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Major )( 
            _App * This,
            /* [retval][out] */ short *rhs);
        
        DECLSPEC_XFGVIRT(_App, Missing47)
        /* [helpstring] */ HRESULT ( __stdcall *Missing47 )( 
            _App * This);
        
        DECLSPEC_XFGVIRT(_App, get_Minor)
        /* [helpcontext][helpstring][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Minor )( 
            _App * This,
            /* [retval][out] */ short *rhs);
        
        DECLSPEC_XFGVIRT(_App, Missing49)
        /* [helpstring] */ HRESULT ( __stdcall *Missing49 )( 
            _App * This);
        
        DECLSPEC_XFGVIRT(_App, get_Revision)
        /* [helpcontext][helpstring][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Revision )( 
            _App * This,
            /* [retval][out] */ short *rhs);
        
        DECLSPEC_XFGVIRT(_App, Missing51)
        /* [helpstring] */ HRESULT ( __stdcall *Missing51 )( 
            _App * This);
        
        DECLSPEC_XFGVIRT(_App, get_Comments)
        /* [helpcontext][helpstring][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Comments )( 
            _App * This,
            /* [retval][out] */ BSTR *rhs);
        
        DECLSPEC_XFGVIRT(_App, Missing53)
        /* [helpstring] */ HRESULT ( __stdcall *Missing53 )( 
            _App * This);
        
        DECLSPEC_XFGVIRT(_App, get_CompanyName)
        /* [helpcontext][helpstring][propget] */ HRESULT ( STDMETHODCALLTYPE *get_CompanyName )( 
            _App * This,
            /* [retval][out] */ BSTR *rhs);
        
        DECLSPEC_XFGVIRT(_App, Missing55)
        /* [helpstring] */ HRESULT ( __stdcall *Missing55 )( 
            _App * This);
        
        DECLSPEC_XFGVIRT(_App, get_FileDescription)
        /* [helpcontext][helpstring][propget] */ HRESULT ( STDMETHODCALLTYPE *get_FileDescription )( 
            _App * This,
            /* [retval][out] */ BSTR *rhs);
        
        DECLSPEC_XFGVIRT(_App, Missing57)
        /* [helpstring] */ HRESULT ( __stdcall *Missing57 )( 
            _App * This);
        
        DECLSPEC_XFGVIRT(_App, get_LegalCopyright)
        /* [helpcontext][helpstring][propget] */ HRESULT ( STDMETHODCALLTYPE *get_LegalCopyright )( 
            _App * This,
            /* [retval][out] */ BSTR *rhs);
        
        DECLSPEC_XFGVIRT(_App, Missing59)
        /* [helpstring] */ HRESULT ( __stdcall *Missing59 )( 
            _App * This);
        
        DECLSPEC_XFGVIRT(_App, get_LegalTrademarks)
        /* [helpcontext][helpstring][propget] */ HRESULT ( STDMETHODCALLTYPE *get_LegalTrademarks )( 
            _App * This,
            /* [retval][out] */ BSTR *rhs);
        
        DECLSPEC_XFGVIRT(_App, Missing61)
        /* [helpstring] */ HRESULT ( __stdcall *Missing61 )( 
            _App * This);
        
        DECLSPEC_XFGVIRT(_App, get_ProductName)
        /* [helpcontext][helpstring][propget] */ HRESULT ( STDMETHODCALLTYPE *get_ProductName )( 
            _App * This,
            /* [retval][out] */ BSTR *rhs);
        
        DECLSPEC_XFGVIRT(_App, Missing63)
        /* [helpstring] */ HRESULT ( __stdcall *Missing63 )( 
            _App * This);
        
        DECLSPEC_XFGVIRT(_App, get_hInstance)
        /* [helpcontext][helpstring][propget] */ HRESULT ( STDMETHODCALLTYPE *get_hInstance )( 
            _App * This,
            /* [retval][out] */ long *rhs);
        
        DECLSPEC_XFGVIRT(_App, Missing65)
        /* [helpstring] */ HRESULT ( __stdcall *Missing65 )( 
            _App * This);
        
        DECLSPEC_XFGVIRT(_App, get_NonModalAllowed)
        /* [helpcontext][helpstring][propget] */ HRESULT ( STDMETHODCALLTYPE *get_NonModalAllowed )( 
            _App * This,
            /* [retval][out] */ VARIANT_BOOL *rhs);
        
        DECLSPEC_XFGVIRT(_App, Missing67)
        /* [helpstring] */ HRESULT ( __stdcall *Missing67 )( 
            _App * This);
        
        DECLSPEC_XFGVIRT(_App, get_LogPath)
        /* [helpcontext][helpstring][propget] */ HRESULT ( STDMETHODCALLTYPE *get_LogPath )( 
            _App * This,
            /* [retval][out] */ BSTR *rhs);
        
        DECLSPEC_XFGVIRT(_App, Missing69)
        /* [helpstring] */ HRESULT ( __stdcall *Missing69 )( 
            _App * This);
        
        DECLSPEC_XFGVIRT(_App, get_LogMode)
        /* [helpcontext][helpstring][propget] */ HRESULT ( STDMETHODCALLTYPE *get_LogMode )( 
            _App * This,
            /* [retval][out] */ long *rhs);
        
        DECLSPEC_XFGVIRT(_App, Missing71)
        /* [helpstring] */ HRESULT ( __stdcall *Missing71 )( 
            _App * This);
        
        DECLSPEC_XFGVIRT(_App, get_UnattendedApp)
        /* [helpcontext][helpstring][propget] */ HRESULT ( STDMETHODCALLTYPE *get_UnattendedApp )( 
            _App * This,
            /* [retval][out] */ VARIANT_BOOL *rhs);
        
        DECLSPEC_XFGVIRT(_App, Missing73)
        /* [helpstring] */ HRESULT ( __stdcall *Missing73 )( 
            _App * This);
        
        DECLSPEC_XFGVIRT(_App, get_ThreadID)
        /* [helpcontext][helpstring][propget] */ HRESULT ( STDMETHODCALLTYPE *get_ThreadID )( 
            _App * This,
            /* [retval][out] */ long *rhs);
        
        DECLSPEC_XFGVIRT(_App, Missing75)
        /* [helpstring] */ HRESULT ( __stdcall *Missing75 )( 
            _App * This);
        
        DECLSPEC_XFGVIRT(_App, get_HelpFile)
        /* [helpcontext][helpstring][propget] */ HRESULT ( STDMETHODCALLTYPE *get_HelpFile )( 
            _App * This,
            /* [retval][out] */ BSTR *rhs);
        
        DECLSPEC_XFGVIRT(_App, put_HelpFile)
        /* [helpcontext][helpstring][propput] */ HRESULT ( STDMETHODCALLTYPE *put_HelpFile )( 
            _App * This,
            /* [in] */ BSTR rhs);
        
        DECLSPEC_XFGVIRT(_App, get_RetainedProject)
        /* [helpcontext][helpstring][propget] */ HRESULT ( STDMETHODCALLTYPE *get_RetainedProject )( 
            _App * This,
            /* [retval][out] */ VARIANT_BOOL *rhs);
        
        DECLSPEC_XFGVIRT(_App, StartLogging)
        /* [helpcontext][helpstring] */ HRESULT ( STDMETHODCALLTYPE *StartLogging )( 
            _App * This,
            /* [in] */ BSTR LogTarget,
            /* [in] */ long LogModes);
        
        DECLSPEC_XFGVIRT(_App, LogEvent)
        /* [helpcontext][helpstring] */ HRESULT ( STDMETHODCALLTYPE *LogEvent )( 
            _App * This,
            /* [in] */ BSTR LogBuffer,
            /* [in] */ VARIANT EventType);
        
        END_INTERFACE
    } _AppVtbl;

    interface _App
    {
        CONST_VTBL struct _AppVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define _App_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define _App_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define _App_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define _App_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define _App_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define _App_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define _App_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define _App_HctlDemandLoad(This,ctl)	\
    ( (This)->lpVtbl -> HctlDemandLoad(This,ctl) ) 

#define _App_ChkProp(This,i,tagData)	\
    ( (This)->lpVtbl -> ChkProp(This,i,tagData) ) 

#define _App_SetPropA(This,i,tagData)	\
    ( (This)->lpVtbl -> SetPropA(This,i,tagData) ) 

#define _App_GetPropA(This,i,tagData)	\
    ( (This)->lpVtbl -> GetPropA(This,i,tagData) ) 

#define _App_GetPropHsz(This,i,hsz)	\
    ( (This)->lpVtbl -> GetPropHsz(This,i,hsz) ) 

#define _App_LoadProp(This,i,fref)	\
    ( (This)->lpVtbl -> LoadProp(This,i,fref) ) 

#define _App_SaveProp(This,i,fref)	\
    ( (This)->lpVtbl -> SaveProp(This,i,fref) ) 

#define _App_GetPalette(This)	\
    ( (This)->lpVtbl -> GetPalette(This) ) 

#define _App_Reset(This)	\
    ( (This)->lpVtbl -> Reset(This) ) 

#define _App_get_DefaultProp(This,var)	\
    ( (This)->lpVtbl -> get_DefaultProp(This,var) ) 

#define _App_put_DefaultProp(This,var)	\
    ( (This)->lpVtbl -> put_DefaultProp(This,var) ) 

#define _App_get_000x(This,var)	\
    ( (This)->lpVtbl -> get_000x(This,var) ) 

#define _App_put_000x(This,i)	\
    ( (This)->lpVtbl -> put_000x(This,i) ) 

#define _App_get_Path(This,rhs)	\
    ( (This)->lpVtbl -> get_Path(This,rhs) ) 

#define _App_put_Path(This,rhs)	\
    ( (This)->lpVtbl -> put_Path(This,rhs) ) 

#define _App_get_EXEName(This,rhs)	\
    ( (This)->lpVtbl -> get_EXEName(This,rhs) ) 

#define _App_put_EXEName(This,rhs)	\
    ( (This)->lpVtbl -> put_EXEName(This,rhs) ) 

#define _App_get_Title(This,rhs)	\
    ( (This)->lpVtbl -> get_Title(This,rhs) ) 

#define _App_put_Title(This,rhs)	\
    ( (This)->lpVtbl -> put_Title(This,rhs) ) 

#define _App_get_PrevInstance(This,rhs)	\
    ( (This)->lpVtbl -> get_PrevInstance(This,rhs) ) 

#define _App_Missing27(This)	\
    ( (This)->lpVtbl -> Missing27(This) ) 

#define _App_get_StartMode(This,rhs)	\
    ( (This)->lpVtbl -> get_StartMode(This,rhs) ) 

#define _App_Missing29(This)	\
    ( (This)->lpVtbl -> Missing29(This) ) 

#define _App_get_TaskVisible(This,rhs)	\
    ( (This)->lpVtbl -> get_TaskVisible(This,rhs) ) 

#define _App_put_TaskVisible(This,rhs)	\
    ( (This)->lpVtbl -> put_TaskVisible(This,rhs) ) 

#define _App_get_OleServerBusyTimeout(This,rhs)	\
    ( (This)->lpVtbl -> get_OleServerBusyTimeout(This,rhs) ) 

#define _App_put_OleServerBusyTimeout(This,rhs)	\
    ( (This)->lpVtbl -> put_OleServerBusyTimeout(This,rhs) ) 

#define _App_get_OleServerBusyMsgTitle(This,rhs)	\
    ( (This)->lpVtbl -> get_OleServerBusyMsgTitle(This,rhs) ) 

#define _App_put_OleServerBusyMsgTitle(This,rhs)	\
    ( (This)->lpVtbl -> put_OleServerBusyMsgTitle(This,rhs) ) 

#define _App_get_OleServerBusyMsgText(This,rhs)	\
    ( (This)->lpVtbl -> get_OleServerBusyMsgText(This,rhs) ) 

#define _App_put_OleServerBusyMsgText(This,rhs)	\
    ( (This)->lpVtbl -> put_OleServerBusyMsgText(This,rhs) ) 

#define _App_get_OleServerBusyRaiseError(This,rhs)	\
    ( (This)->lpVtbl -> get_OleServerBusyRaiseError(This,rhs) ) 

#define _App_put_OleServerBusyRaiseError(This,rhs)	\
    ( (This)->lpVtbl -> put_OleServerBusyRaiseError(This,rhs) ) 

#define _App_get_OleRequestPendingTimeout(This,rhs)	\
    ( (This)->lpVtbl -> get_OleRequestPendingTimeout(This,rhs) ) 

#define _App_put_OleRequestPendingTimeout(This,rhs)	\
    ( (This)->lpVtbl -> put_OleRequestPendingTimeout(This,rhs) ) 

#define _App_get_OleRequestPendingMsgTitle(This,rhs)	\
    ( (This)->lpVtbl -> get_OleRequestPendingMsgTitle(This,rhs) ) 

#define _App_put_OleRequestPendingMsgTitle(This,rhs)	\
    ( (This)->lpVtbl -> put_OleRequestPendingMsgTitle(This,rhs) ) 

#define _App_get_OleRequestPendingMsgText(This,rhs)	\
    ( (This)->lpVtbl -> get_OleRequestPendingMsgText(This,rhs) ) 

#define _App_put_OleRequestPendingMsgText(This,rhs)	\
    ( (This)->lpVtbl -> put_OleRequestPendingMsgText(This,rhs) ) 

#define _App_get_Major(This,rhs)	\
    ( (This)->lpVtbl -> get_Major(This,rhs) ) 

#define _App_Missing47(This)	\
    ( (This)->lpVtbl -> Missing47(This) ) 

#define _App_get_Minor(This,rhs)	\
    ( (This)->lpVtbl -> get_Minor(This,rhs) ) 

#define _App_Missing49(This)	\
    ( (This)->lpVtbl -> Missing49(This) ) 

#define _App_get_Revision(This,rhs)	\
    ( (This)->lpVtbl -> get_Revision(This,rhs) ) 

#define _App_Missing51(This)	\
    ( (This)->lpVtbl -> Missing51(This) ) 

#define _App_get_Comments(This,rhs)	\
    ( (This)->lpVtbl -> get_Comments(This,rhs) ) 

#define _App_Missing53(This)	\
    ( (This)->lpVtbl -> Missing53(This) ) 

#define _App_get_CompanyName(This,rhs)	\
    ( (This)->lpVtbl -> get_CompanyName(This,rhs) ) 

#define _App_Missing55(This)	\
    ( (This)->lpVtbl -> Missing55(This) ) 

#define _App_get_FileDescription(This,rhs)	\
    ( (This)->lpVtbl -> get_FileDescription(This,rhs) ) 

#define _App_Missing57(This)	\
    ( (This)->lpVtbl -> Missing57(This) ) 

#define _App_get_LegalCopyright(This,rhs)	\
    ( (This)->lpVtbl -> get_LegalCopyright(This,rhs) ) 

#define _App_Missing59(This)	\
    ( (This)->lpVtbl -> Missing59(This) ) 

#define _App_get_LegalTrademarks(This,rhs)	\
    ( (This)->lpVtbl -> get_LegalTrademarks(This,rhs) ) 

#define _App_Missing61(This)	\
    ( (This)->lpVtbl -> Missing61(This) ) 

#define _App_get_ProductName(This,rhs)	\
    ( (This)->lpVtbl -> get_ProductName(This,rhs) ) 

#define _App_Missing63(This)	\
    ( (This)->lpVtbl -> Missing63(This) ) 

#define _App_get_hInstance(This,rhs)	\
    ( (This)->lpVtbl -> get_hInstance(This,rhs) ) 

#define _App_Missing65(This)	\
    ( (This)->lpVtbl -> Missing65(This) ) 

#define _App_get_NonModalAllowed(This,rhs)	\
    ( (This)->lpVtbl -> get_NonModalAllowed(This,rhs) ) 

#define _App_Missing67(This)	\
    ( (This)->lpVtbl -> Missing67(This) ) 

#define _App_get_LogPath(This,rhs)	\
    ( (This)->lpVtbl -> get_LogPath(This,rhs) ) 

#define _App_Missing69(This)	\
    ( (This)->lpVtbl -> Missing69(This) ) 

#define _App_get_LogMode(This,rhs)	\
    ( (This)->lpVtbl -> get_LogMode(This,rhs) ) 

#define _App_Missing71(This)	\
    ( (This)->lpVtbl -> Missing71(This) ) 

#define _App_get_UnattendedApp(This,rhs)	\
    ( (This)->lpVtbl -> get_UnattendedApp(This,rhs) ) 

#define _App_Missing73(This)	\
    ( (This)->lpVtbl -> Missing73(This) ) 

#define _App_get_ThreadID(This,rhs)	\
    ( (This)->lpVtbl -> get_ThreadID(This,rhs) ) 

#define _App_Missing75(This)	\
    ( (This)->lpVtbl -> Missing75(This) ) 

#define _App_get_HelpFile(This,rhs)	\
    ( (This)->lpVtbl -> get_HelpFile(This,rhs) ) 

#define _App_put_HelpFile(This,rhs)	\
    ( (This)->lpVtbl -> put_HelpFile(This,rhs) ) 

#define _App_get_RetainedProject(This,rhs)	\
    ( (This)->lpVtbl -> get_RetainedProject(This,rhs) ) 

#define _App_StartLogging(This,LogTarget,LogModes)	\
    ( (This)->lpVtbl -> StartLogging(This,LogTarget,LogModes) ) 

#define _App_LogEvent(This,LogBuffer,EventType)	\
    ( (This)->lpVtbl -> LogEvent(This,LogBuffer,EventType) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* ___App_INTERFACE_DEFINED__ */


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


