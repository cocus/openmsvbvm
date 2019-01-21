

/* this ALWAYS GENERATED file contains the definitions for the interfaces */


 /* File created by MIDL compiler version 8.01.0622 */
/* at Tue Jan 19 00:14:07 2038
 */
/* Compiler settings for vba_objVBGlobal.idl:
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

#ifndef __vba_objVBGlobal_h_h__
#define __vba_objVBGlobal_h_h__

#if defined(_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

/* Forward Declarations */ 

#ifndef __IVBGlobal_FWD_DEFINED__
#define __IVBGlobal_FWD_DEFINED__
typedef interface IVBGlobal IVBGlobal;

#endif 	/* __IVBGlobal_FWD_DEFINED__ */


/* header files for imported files */
#include "oaidl.h"
#include "ocidl.h"
#include "vba_objApp.h"

#ifdef __cplusplus
extern "C"{
#endif 


#ifndef __IVBGlobal_INTERFACE_DEFINED__
#define __IVBGlobal_INTERFACE_DEFINED__

/* interface IVBGlobal */
/* [object][helpcontext][helpstring][uuid] */ 


EXTERN_C const IID IID_IVBGlobal;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("fcfb3d22-a0fa-1068-a738-08002b3371b5")
    IVBGlobal : public IUnknown
    {
    public:
        virtual /* [helpcontext][helpstring] */ HRESULT __stdcall Load( 
            /* [in] */ IDispatch *object) = 0;
        
        virtual /* [helpcontext][helpstring] */ HRESULT __stdcall Unload( 
            /* [in] */ IDispatch *object) = 0;
        
        virtual /* [helpcontext][helpstring][propget] */ HRESULT __stdcall get_App( 
            /* [retval][out] */ IApp **pdispRetVal) = 0;
        
    };
    
    
#else 	/* C style interface */

    typedef struct IVBGlobalVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IVBGlobal * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IVBGlobal * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IVBGlobal * This);
        
        /* [helpcontext][helpstring] */ HRESULT ( __stdcall *Load )( 
            IVBGlobal * This,
            /* [in] */ IDispatch *object);
        
        /* [helpcontext][helpstring] */ HRESULT ( __stdcall *Unload )( 
            IVBGlobal * This,
            /* [in] */ IDispatch *object);
        
        /* [helpcontext][helpstring][propget] */ HRESULT ( __stdcall *get_App )( 
            IVBGlobal * This,
            /* [retval][out] */ IApp **pdispRetVal);
        
        END_INTERFACE
    } IVBGlobalVtbl;

    interface IVBGlobal
    {
        CONST_VTBL struct IVBGlobalVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IVBGlobal_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IVBGlobal_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IVBGlobal_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IVBGlobal_Load(This,object)	\
    ( (This)->lpVtbl -> Load(This,object) ) 

#define IVBGlobal_Unload(This,object)	\
    ( (This)->lpVtbl -> Unload(This,object) ) 

#define IVBGlobal_get_App(This,pdispRetVal)	\
    ( (This)->lpVtbl -> get_App(This,pdispRetVal) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __IVBGlobal_INTERFACE_DEFINED__ */


/* Additional Prototypes for ALL interfaces */

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif


