

/* this ALWAYS GENERATED file contains the definitions for the interfaces */


 /* File created by MIDL compiler version 8.01.0628 */
/* at Tue Jan 19 00:14:07 2038
 */
/* Compiler settings for VBGlobal.idl:
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

#ifndef __VBGlobal_h__
#define __VBGlobal_h__

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

#ifndef __Global_FWD_DEFINED__
#define __Global_FWD_DEFINED__

#ifdef __cplusplus
typedef class Global Global;
#else
typedef struct Global Global;
#endif /* __cplusplus */

#endif 	/* __Global_FWD_DEFINED__ */


#ifndef __VBGlobal_FWD_DEFINED__
#define __VBGlobal_FWD_DEFINED__
typedef interface VBGlobal VBGlobal;

#endif 	/* __VBGlobal_FWD_DEFINED__ */


/* header files for imported files */
#include "oaidl.h"
#include "ocidl.h"
#include "App.h"

#ifdef __cplusplus
extern "C"{
#endif 


#ifndef __VBGlobal_INTERFACE_DEFINED__
#define __VBGlobal_INTERFACE_DEFINED__

/* interface VBGlobal */
/* [object][helpcontext][helpstring][uuid] */ 


EXTERN_C const IID IID_VBGlobal;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("fcfb3d22-a0fa-1068-a738-08002b3371b5")
    VBGlobal : public IUnknown
    {
    public:
        virtual /* [helpcontext][helpstring] */ HRESULT __stdcall Load( 
            /* [in] */ IDispatch *object) = 0;
        
        virtual /* [helpcontext][helpstring] */ HRESULT __stdcall Unload( 
            /* [in] */ IDispatch *object) = 0;
        
        virtual /* [helpcontext][helpstring][propget] */ HRESULT __stdcall get_App( 
            /* [retval][out] */ _App **pdispRetVal) = 0;
        
    };
    
    
#else 	/* C style interface */

    typedef struct VBGlobalVtbl
    {
        BEGIN_INTERFACE
        
        DECLSPEC_XFGVIRT(IUnknown, QueryInterface)
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            VBGlobal * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        DECLSPEC_XFGVIRT(IUnknown, AddRef)
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            VBGlobal * This);
        
        DECLSPEC_XFGVIRT(IUnknown, Release)
        ULONG ( STDMETHODCALLTYPE *Release )( 
            VBGlobal * This);
        
        DECLSPEC_XFGVIRT(VBGlobal, Load)
        /* [helpcontext][helpstring] */ HRESULT ( __stdcall *Load )( 
            VBGlobal * This,
            /* [in] */ IDispatch *object);
        
        DECLSPEC_XFGVIRT(VBGlobal, Unload)
        /* [helpcontext][helpstring] */ HRESULT ( __stdcall *Unload )( 
            VBGlobal * This,
            /* [in] */ IDispatch *object);
        
        DECLSPEC_XFGVIRT(VBGlobal, get_App)
        /* [helpcontext][helpstring][propget] */ HRESULT ( __stdcall *get_App )( 
            VBGlobal * This,
            /* [retval][out] */ _App **pdispRetVal);
        
        END_INTERFACE
    } VBGlobalVtbl;

    interface VBGlobal
    {
        CONST_VTBL struct VBGlobalVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define VBGlobal_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define VBGlobal_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define VBGlobal_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define VBGlobal_Load(This,object)	\
    ( (This)->lpVtbl -> Load(This,object) ) 

#define VBGlobal_Unload(This,object)	\
    ( (This)->lpVtbl -> Unload(This,object) ) 

#define VBGlobal_get_App(This,pdispRetVal)	\
    ( (This)->lpVtbl -> get_App(This,pdispRetVal) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __VBGlobal_INTERFACE_DEFINED__ */


/* Additional Prototypes for ALL interfaces */

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif


