

/* this ALWAYS GENERATED file contains the definitions for the interfaces */


 /* File created by MIDL compiler version 8.01.0628 */
/* at Tue Jan 19 00:14:07 2038
 */
/* Compiler settings for Form.idl:
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

#ifndef __Form_h__
#define __Form_h__

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

#ifndef __Form_FWD_DEFINED__
#define __Form_FWD_DEFINED__

#ifdef __cplusplus
typedef class Form Form;
#else
typedef struct Form Form;
#endif /* __cplusplus */

#endif 	/* __Form_FWD_DEFINED__ */


#ifndef ___Form_FWD_DEFINED__
#define ___Form_FWD_DEFINED__
typedef interface _Form _Form;

#endif 	/* ___Form_FWD_DEFINED__ */


/* header files for imported files */
#include "oaidl.h"
#include "ocidl.h"

#ifdef __cplusplus
extern "C"{
#endif 


#ifndef ___Form_INTERFACE_DEFINED__
#define ___Form_INTERFACE_DEFINED__

/* interface _Form */
/* [object][nonextensible][hidden][helpstring][uuid] */ 


EXTERN_C const IID IID__Form;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("33AD4F39-6699-11CF-B70C-00AA0060D393")
    _Form : public IDispatch
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
        
        virtual /* [helpstring][propget] */ HRESULT __stdcall get_hWnd( 
            /* [retval][out] */ long *rhs) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall Missing13( void) = 0;
        
        virtual /* [helpstring][propget] */ HRESULT __stdcall get_BackColor( 
            /* [retval][out] */ long *rhs) = 0;
        
        virtual /* [helpstring][propput] */ HRESULT __stdcall put_BackColor( 
            /* [in] */ long rhs) = 0;
        
        virtual /* [helpstring][propget] */ HRESULT __stdcall get_ForeColor( 
            /* [retval][out] */ long *rhs) = 0;
        
        virtual /* [helpstring][propput] */ HRESULT __stdcall put_ForeColor( 
            /* [in] */ long rhs) = 0;
        
        virtual /* [helpstring][propget] */ HRESULT __stdcall get_Left( 
            /* [retval][out] */ long *rhs) = 0;
        
        virtual /* [helpstring][propput] */ HRESULT __stdcall put_Left( 
            /* [in] */ long rhs) = 0;
        
        virtual /* [helpstring][propget] */ HRESULT __stdcall get_Top( 
            /* [retval][out] */ long *rhs) = 0;
        
        virtual /* [helpstring][propput] */ HRESULT __stdcall put_Top( 
            /* [in] */ long rhs) = 0;
        
        virtual /* [helpstring][propget] */ HRESULT __stdcall get_Width( 
            /* [retval][out] */ long *rhs) = 0;
        
        virtual /* [helpstring][propput] */ HRESULT __stdcall put_Width( 
            /* [in] */ long rhs) = 0;
        
        virtual /* [helpstring][propget] */ HRESULT __stdcall get_Height( 
            /* [retval][out] */ long *rhs) = 0;
        
        virtual /* [helpstring][propput] */ HRESULT __stdcall put_Height( 
            /* [in] */ long rhs) = 0;
        
        virtual /* [helpstring][propget] */ HRESULT __stdcall get_Enabled( 
            /* [retval][out] */ VARIANT_BOOL *rhs) = 0;
        
        virtual /* [helpstring][propput] */ HRESULT __stdcall put_Enabled( 
            /* [in] */ VARIANT_BOOL rhs) = 0;
        
        virtual /* [helpstring][propget] */ HRESULT __stdcall get_WindowState( 
            /* [retval][out] */ long *rhs) = 0;
        
        virtual /* [helpstring][propput] */ HRESULT __stdcall put_WindowState( 
            /* [in] */ long rhs) = 0;
        
        virtual /* [helpstring][propget] */ HRESULT __stdcall get_MousePointer( 
            /* [retval][out] */ long *rhs) = 0;
        
        virtual /* [helpstring][propput] */ HRESULT __stdcall put_MousePointer( 
            /* [in] */ long rhs) = 0;
        
        virtual /* [helpstring][propget] */ HRESULT __stdcall get_FontName( 
            /* [retval][out] */ BSTR *rhs) = 0;
        
        virtual /* [helpstring][propput] */ HRESULT __stdcall put_FontName( 
            /* [in] */ BSTR rhs) = 0;
        
        virtual /* [helpstring][propget] */ HRESULT __stdcall get_FontSize( 
            /* [retval][out] */ float *rhs) = 0;
        
        virtual /* [helpstring][propput] */ HRESULT __stdcall put_FontSize( 
            /* [in] */ float rhs) = 0;
        
        virtual /* [helpstring][propget] */ HRESULT __stdcall get_FontBold( 
            /* [retval][out] */ VARIANT_BOOL *rhs) = 0;
        
        virtual /* [helpstring][propput] */ HRESULT __stdcall put_FontBold( 
            /* [in] */ VARIANT_BOOL rhs) = 0;
        
        virtual /* [helpstring][propget] */ HRESULT __stdcall get_FontItalic( 
            /* [retval][out] */ VARIANT_BOOL *rhs) = 0;
        
        virtual /* [helpstring][propput] */ HRESULT __stdcall put_FontItalic( 
            /* [in] */ VARIANT_BOOL rhs) = 0;
        
        virtual /* [helpstring][propget] */ HRESULT __stdcall get_FontStrikethru( 
            /* [retval][out] */ VARIANT_BOOL *rhs) = 0;
        
        virtual /* [helpstring][propput] */ HRESULT __stdcall put_FontStrikethru( 
            /* [in] */ VARIANT_BOOL rhs) = 0;
        
        virtual /* [helpstring][propget] */ HRESULT __stdcall get_FontUnderline( 
            /* [retval][out] */ VARIANT_BOOL *rhs) = 0;
        
        virtual /* [helpstring][propput] */ HRESULT __stdcall put_FontUnderline( 
            /* [in] */ VARIANT_BOOL rhs) = 0;
        
        virtual /* [helpstring][propget] */ HRESULT __stdcall get_hDC( 
            /* [retval][out] */ long *rhs) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall Missing14( void) = 0;
        
        virtual /* [helpstring][propget] */ HRESULT __stdcall get_CurrentX( 
            /* [retval][out] */ float *rhs) = 0;
        
        virtual /* [helpstring][propput] */ HRESULT __stdcall put_CurrentX( 
            /* [in] */ float rhs) = 0;
        
        virtual /* [helpstring][propget] */ HRESULT __stdcall get_CurrentY( 
            /* [retval][out] */ float *rhs) = 0;
        
        virtual /* [helpstring][propput] */ HRESULT __stdcall put_CurrentY( 
            /* [in] */ float rhs) = 0;
        
        virtual /* [helpstring][propget] */ HRESULT __stdcall get_ScaleLeft( 
            /* [retval][out] */ float *rhs) = 0;
        
        virtual /* [helpstring][propput] */ HRESULT __stdcall put_ScaleLeft( 
            /* [in] */ float rhs) = 0;
        
        virtual /* [helpstring][propget] */ HRESULT __stdcall get_ScaleTop( 
            /* [retval][out] */ float *rhs) = 0;
        
        virtual /* [helpstring][propput] */ HRESULT __stdcall put_ScaleTop( 
            /* [in] */ float rhs) = 0;
        
        virtual /* [helpstring][propget] */ HRESULT __stdcall get_ScaleWidth( 
            /* [retval][out] */ float *rhs) = 0;
        
        virtual /* [helpstring][propput] */ HRESULT __stdcall put_ScaleWidth( 
            /* [in] */ float rhs) = 0;
        
        virtual /* [helpstring][propget] */ HRESULT __stdcall get_ScaleHeight( 
            /* [retval][out] */ float *rhs) = 0;
        
        virtual /* [helpstring][propput] */ HRESULT __stdcall put_ScaleHeight( 
            /* [in] */ float rhs) = 0;
        
        virtual /* [helpstring][propget] */ HRESULT __stdcall get_ScaleMode( 
            /* [retval][out] */ short *rhs) = 0;
        
        virtual /* [helpstring][propput] */ HRESULT __stdcall put_ScaleMode( 
            /* [in] */ short rhs) = 0;
        
        virtual /* [helpstring][propget] */ HRESULT __stdcall get_FontTransparent( 
            /* [retval][out] */ VARIANT_BOOL *rhs) = 0;
        
        virtual /* [helpstring][propput] */ HRESULT __stdcall put_FontTransparent( 
            /* [in] */ VARIANT_BOOL rhs) = 0;
        
        virtual /* [helpstring][propget] */ HRESULT __stdcall get_DrawStyle( 
            /* [retval][out] */ short *rhs) = 0;
        
        virtual /* [helpstring][propput] */ HRESULT __stdcall put_DrawStyle( 
            /* [in] */ short rhs) = 0;
        
        virtual /* [helpstring][propget] */ HRESULT __stdcall get_DrawWidth( 
            /* [retval][out] */ short *rhs) = 0;
        
        virtual /* [helpstring][propput] */ HRESULT __stdcall put_DrawWidth( 
            /* [in] */ short rhs) = 0;
        
        virtual /* [helpstring][propget] */ HRESULT __stdcall get_FillStyle( 
            /* [retval][out] */ short *rhs) = 0;
        
        virtual /* [helpstring][propput] */ HRESULT __stdcall put_FillStyle( 
            /* [in] */ short rhs) = 0;
        
        virtual /* [helpstring][propget] */ HRESULT __stdcall get_FillColor( 
            /* [retval][out] */ long *rhs) = 0;
        
        virtual /* [helpstring][propput] */ HRESULT __stdcall put_FillColor( 
            /* [in] */ long rhs) = 0;
        
        virtual /* [helpstring][propget] */ HRESULT __stdcall get_DrawMode( 
            /* [retval][out] */ short *rhs) = 0;
        
        virtual /* [helpstring][propput] */ HRESULT __stdcall put_DrawMode( 
            /* [in] */ short rhs) = 0;
        
        virtual /* [helpstring][propget] */ HRESULT __stdcall get_AutoRedraw( 
            /* [retval][out] */ VARIANT_BOOL *rhs) = 0;
        
        virtual /* [helpstring][propput] */ HRESULT __stdcall put_AutoRedraw( 
            /* [in] */ VARIANT_BOOL rhs) = 0;
        
        virtual /* [helpstring][propget] */ HRESULT __stdcall get_Picture( 
            /* [retval][out] */ long *rhs) = 0;
        
        virtual /* [helpstring][propput] */ HRESULT __stdcall put_Picture( 
            /* [in] */ long rhs) = 0;
        
        virtual /* [helpstring][propget] */ HRESULT __stdcall get_BorderStyle( 
            /* [retval][out] */ short *rhs) = 0;
        
        virtual /* [helpstring][propput] */ HRESULT __stdcall put_BorderStyle( 
            /* [in] */ short rhs) = 0;
        
        virtual /* [helpstring][propget] */ HRESULT __stdcall get_Icon( 
            /* [retval][out] */ long *rhs) = 0;
        
        virtual /* [helpstring][propput] */ HRESULT __stdcall put_Icon( 
            /* [in] */ long rhs) = 0;
        
        virtual /* [helpstring][propget] */ HRESULT __stdcall get_LinkTopic( 
            /* [retval][out] */ BSTR *rhs) = 0;
        
        virtual /* [helpstring][propput] */ HRESULT __stdcall put_LinkTopic( 
            /* [in] */ BSTR rhs) = 0;
        
        virtual /* [helpstring][propget] */ HRESULT __stdcall get_LinkMode( 
            /* [retval][out] */ short *rhs) = 0;
        
        virtual /* [helpstring][propput] */ HRESULT __stdcall put_LinkMode( 
            /* [in] */ short rhs) = 0;
        
        virtual /* [helpstring][propget] */ HRESULT __stdcall get_MaxButton( 
            /* [retval][out] */ VARIANT_BOOL *rhs) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall Missing15( void) = 0;
        
        virtual /* [helpstring][propget] */ HRESULT __stdcall get_MinButton( 
            /* [retval][out] */ VARIANT_BOOL *rhs) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall Missing16( void) = 0;
        
        virtual /* [helpstring][propget] */ HRESULT __stdcall get_ControlBox( 
            /* [retval][out] */ VARIANT_BOOL *rhs) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall Missing17( void) = 0;
        
        virtual /* [helpstring][propget] */ HRESULT __stdcall get_Image( 
            /* [retval][out] */ long *rhs) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall Missing18( void) = 0;
        
        virtual /* [helpstring][propget] */ HRESULT __stdcall get_HasDC( 
            /* [retval][out] */ VARIANT_BOOL *rhs) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall Missing19( void) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall Missing20( void) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall Missing21( void) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall Missing22( void) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall Missing23( void) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall Missing24( void) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall Missing25( void) = 0;
        
        virtual /* [helpstring][propget] */ HRESULT __stdcall get_Visible( 
            /* [retval][out] */ VARIANT_BOOL *rhs) = 0;
        
        virtual /* [helpstring][propput] */ HRESULT __stdcall put_Visible( 
            /* [in] */ VARIANT_BOOL rhs) = 0;
        
        virtual /* [helpstring][propget] */ HRESULT __stdcall get_Tag( 
            /* [retval][out] */ BSTR *rhs) = 0;
        
        virtual /* [helpstring][propput] */ HRESULT __stdcall put_Tag( 
            /* [in] */ BSTR rhs) = 0;
        
        virtual /* [helpstring][propget] */ HRESULT __stdcall get_MDIChild( 
            /* [retval][out] */ VARIANT_BOOL *rhs) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall Missing26( void) = 0;
        
        virtual /* [helpstring][propget] */ HRESULT __stdcall get_KeyPreview( 
            /* [retval][out] */ VARIANT_BOOL *rhs) = 0;
        
        virtual /* [helpstring][propput] */ HRESULT __stdcall put_KeyPreview( 
            /* [in] */ VARIANT_BOOL rhs) = 0;
        
        virtual /* [helpstring][propget] */ HRESULT __stdcall get_ClipControls( 
            /* [retval][out] */ VARIANT_BOOL *rhs) = 0;
        
        virtual /* [helpstring][propput] */ HRESULT __stdcall put_ClipControls( 
            /* [in] */ VARIANT_BOOL rhs) = 0;
        
        virtual /* [helpstring][propget] */ HRESULT __stdcall get_HelpContextID( 
            /* [retval][out] */ long *rhs) = 0;
        
        virtual /* [helpstring][propput] */ HRESULT __stdcall put_HelpContextID( 
            /* [in] */ long rhs) = 0;
        
        virtual /* [helpstring][propget] */ HRESULT __stdcall get_ActiveControl( 
            /* [retval][out] */ long *rhs) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall Missing27( void) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall Missing28( void) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall Missing29( void) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall Missing30( void) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall Missing31( void) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall Missing32( void) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall Missing33( void) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall Missing34( void) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall Missing35( void) = 0;
        
        virtual /* [helpstring][propget] */ HRESULT __stdcall get_Count( 
            /* [retval][out] */ short *rhs) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall Missing36( void) = 0;
        
        virtual /* [helpstring][propget] */ HRESULT __stdcall get_Controls( 
            /* [retval][out] */ long *rhs) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall Missing37( void) = 0;
        
        virtual /* [helpstring][propget] */ HRESULT __stdcall get_MouseIcon( 
            /* [retval][out] */ long *rhs) = 0;
        
        virtual /* [helpstring][propput] */ HRESULT __stdcall put_MouseIcon( 
            /* [in] */ long rhs) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall Missing38( void) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall Missing39( void) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall Missing40( void) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall Missing41( void) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall Missing42( void) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall Missing43( void) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall Missing44( void) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall Missing45( void) = 0;
        
        virtual /* [helpstring][propget] */ HRESULT __stdcall get_Font( 
            /* [retval][out] */ long *rhs) = 0;
        
        virtual /* [helpstring][propputref] */ HRESULT __stdcall putref_Font( 
            /* [in] */ long rhs) = 0;
        
        virtual /* [helpstring][propget] */ HRESULT __stdcall get_Appearance( 
            /* [retval][out] */ short *rhs) = 0;
        
        virtual /* [helpstring][propput] */ HRESULT __stdcall put_Appearance( 
            /* [in] */ short rhs) = 0;
        
        virtual /* [helpstring][propget] */ HRESULT __stdcall get_WhatsThisButton( 
            /* [retval][out] */ VARIANT_BOOL *rhs) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall Missing46( void) = 0;
        
        virtual /* [helpstring][propget] */ HRESULT __stdcall get_WhatsThisHelp( 
            /* [retval][out] */ VARIANT_BOOL *rhs) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall Missing47( void) = 0;
        
        virtual /* [helpstring][propget] */ HRESULT __stdcall get_ShowInTaskbar( 
            /* [retval][out] */ VARIANT_BOOL *rhs) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall Missing48( void) = 0;
        
        virtual /* [helpstring][propget] */ HRESULT __stdcall get_RightToLeft( 
            /* [retval][out] */ VARIANT_BOOL *rhs) = 0;
        
        virtual /* [helpstring][propput] */ HRESULT __stdcall put_RightToLeft( 
            /* [in] */ VARIANT_BOOL rhs) = 0;
        
        virtual /* [helpstring][propget] */ HRESULT __stdcall get_StartUpPosition( 
            /* [retval][out] */ short *rhs) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall Missing49( void) = 0;
        
        virtual /* [helpstring][propget] */ HRESULT __stdcall get_OLEDropMode( 
            /* [retval][out] */ short *rhs) = 0;
        
        virtual /* [helpstring][propput] */ HRESULT __stdcall put_OLEDropMode( 
            /* [in] */ short rhs) = 0;
        
        virtual /* [helpstring][propget] */ HRESULT __stdcall get_Palette( 
            /* [retval][out] */ long *rhs) = 0;
        
        virtual /* [helpstring][propputref] */ HRESULT __stdcall putref_Palette( 
            /* [in] */ long rhs) = 0;
        
        virtual /* [helpstring][propget] */ HRESULT __stdcall get_PaletteMode( 
            /* [retval][out] */ short *rhs) = 0;
        
        virtual /* [helpstring][propput] */ HRESULT __stdcall put_PaletteMode( 
            /* [in] */ short rhs) = 0;
        
        virtual /* [helpstring][propget] */ HRESULT __stdcall get_Moveable( 
            /* [retval][out] */ VARIANT_BOOL *rhs) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall Missing50( void) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall Refresh( void) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall Move( 
            /* [in] */ float Left,
            /* [optional][in] */ VARIANT Top,
            /* [optional][in] */ VARIANT Width,
            /* [optional][in] */ VARIANT Height) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall SetFocus( void) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall ZOrder( 
            /* [optional][in] */ VARIANT Position) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall Show( 
            /* [in] */ VARIANT modal,
            /* [in] */ VARIANT OwnerForm) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall Hide( void) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall PrintForm( void) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall PopupMenu( 
            /* [in] */ IDispatch *Menu,
            /* [optional][in] */ VARIANT Flags,
            /* [optional][in] */ VARIANT x,
            /* [optional][in] */ VARIANT y,
            /* [optional][in] */ VARIANT BoldCommand) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall Circle( 
            /* [in] */ float x,
            /* [in] */ float y,
            /* [in] */ float Radius,
            /* [optional][in] */ VARIANT Color,
            /* [optional][in] */ VARIANT StartAngle,
            /* [optional][in] */ VARIANT EndAngle,
            /* [optional][in] */ VARIANT Aspect) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall Cls( void) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall Line( 
            /* [in] */ float x1,
            /* [in] */ float y1,
            /* [in] */ float x2,
            /* [in] */ float y2,
            /* [optional][in] */ VARIANT Color,
            /* [optional][in] */ VARIANT BorF) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall PaintPicture( 
            /* [in] */ IDispatch *Picture,
            /* [in] */ float x1,
            /* [in] */ float y1,
            /* [optional][in] */ VARIANT width1,
            /* [optional][in] */ VARIANT height1,
            /* [optional][in] */ VARIANT x2,
            /* [optional][in] */ VARIANT y2,
            /* [optional][in] */ VARIANT width2,
            /* [optional][in] */ VARIANT height2,
            /* [optional][in] */ VARIANT Opcode) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall Point( 
            /* [in] */ float x,
            /* [in] */ float y,
            /* [retval][out] */ long *rhs) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall PSet( 
            /* [in] */ float x,
            /* [in] */ float y,
            /* [optional][in] */ VARIANT Color) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall Scale( 
            /* [optional][in] */ VARIANT x1,
            /* [optional][in] */ VARIANT y1,
            /* [optional][in] */ VARIANT x2,
            /* [optional][in] */ VARIANT y2) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall ScaleX( 
            /* [in] */ float Value,
            /* [optional][in] */ VARIANT fromScale,
            /* [optional][in] */ VARIANT toScale,
            /* [retval][out] */ float *rhs) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall ScaleY( 
            /* [in] */ float Value,
            /* [optional][in] */ VARIANT fromScale,
            /* [optional][in] */ VARIANT toScale,
            /* [retval][out] */ float *rhs) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall TextWidth( 
            /* [in] */ BSTR str,
            /* [retval][out] */ float *rhs) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall TextHeight( 
            /* [in] */ BSTR str,
            /* [retval][out] */ float *rhs) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall ValidateControls( 
            /* [retval][out] */ VARIANT_BOOL *rhs) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall WhatsThisMode( void) = 0;
        
        virtual /* [helpstring] */ HRESULT __stdcall OLEDrag( void) = 0;
        
    };
    
    
#else 	/* C style interface */

    typedef struct _FormVtbl
    {
        BEGIN_INTERFACE
        
        DECLSPEC_XFGVIRT(IUnknown, QueryInterface)
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            _Form * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        DECLSPEC_XFGVIRT(IUnknown, AddRef)
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            _Form * This);
        
        DECLSPEC_XFGVIRT(IUnknown, Release)
        ULONG ( STDMETHODCALLTYPE *Release )( 
            _Form * This);
        
        DECLSPEC_XFGVIRT(IDispatch, GetTypeInfoCount)
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            _Form * This,
            /* [out] */ UINT *pctinfo);
        
        DECLSPEC_XFGVIRT(IDispatch, GetTypeInfo)
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            _Form * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        DECLSPEC_XFGVIRT(IDispatch, GetIDsOfNames)
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            _Form * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        DECLSPEC_XFGVIRT(IDispatch, Invoke)
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            _Form * This,
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
        
        DECLSPEC_XFGVIRT(_Form, Missing1)
        /* [helpstring] */ HRESULT ( __stdcall *Missing1 )( 
            _Form * This);
        
        DECLSPEC_XFGVIRT(_Form, Missing2)
        /* [helpstring] */ HRESULT ( __stdcall *Missing2 )( 
            _Form * This);
        
        DECLSPEC_XFGVIRT(_Form, Missing3)
        /* [helpstring] */ HRESULT ( __stdcall *Missing3 )( 
            _Form * This);
        
        DECLSPEC_XFGVIRT(_Form, Missing4)
        /* [helpstring] */ HRESULT ( __stdcall *Missing4 )( 
            _Form * This);
        
        DECLSPEC_XFGVIRT(_Form, Missing5)
        /* [helpstring] */ HRESULT ( __stdcall *Missing5 )( 
            _Form * This);
        
        DECLSPEC_XFGVIRT(_Form, Missing6)
        /* [helpstring] */ HRESULT ( __stdcall *Missing6 )( 
            _Form * This);
        
        DECLSPEC_XFGVIRT(_Form, Missing7)
        /* [helpstring] */ HRESULT ( __stdcall *Missing7 )( 
            _Form * This);
        
        DECLSPEC_XFGVIRT(_Form, Missing8)
        /* [helpstring] */ HRESULT ( __stdcall *Missing8 )( 
            _Form * This);
        
        DECLSPEC_XFGVIRT(_Form, Missing9)
        /* [helpstring] */ HRESULT ( __stdcall *Missing9 )( 
            _Form * This);
        
        DECLSPEC_XFGVIRT(_Form, Missing10)
        /* [helpstring] */ HRESULT ( __stdcall *Missing10 )( 
            _Form * This);
        
        DECLSPEC_XFGVIRT(_Form, Missing11)
        /* [helpstring] */ HRESULT ( __stdcall *Missing11 )( 
            _Form * This);
        
        DECLSPEC_XFGVIRT(_Form, get_Name)
        /* [helpstring][propget] */ HRESULT ( __stdcall *get_Name )( 
            _Form * This,
            /* [retval][out] */ BSTR *rhs);
        
        DECLSPEC_XFGVIRT(_Form, Missing12)
        /* [helpstring] */ HRESULT ( __stdcall *Missing12 )( 
            _Form * This);
        
        DECLSPEC_XFGVIRT(_Form, get_Caption)
        /* [helpstring][propget] */ HRESULT ( __stdcall *get_Caption )( 
            _Form * This,
            /* [retval][out] */ BSTR *rhs);
        
        DECLSPEC_XFGVIRT(_Form, put_Caption)
        /* [helpstring][propput] */ HRESULT ( __stdcall *put_Caption )( 
            _Form * This,
            /* [in] */ BSTR rhs);
        
        DECLSPEC_XFGVIRT(_Form, get_hWnd)
        /* [helpstring][propget] */ HRESULT ( __stdcall *get_hWnd )( 
            _Form * This,
            /* [retval][out] */ long *rhs);
        
        DECLSPEC_XFGVIRT(_Form, Missing13)
        /* [helpstring] */ HRESULT ( __stdcall *Missing13 )( 
            _Form * This);
        
        DECLSPEC_XFGVIRT(_Form, get_BackColor)
        /* [helpstring][propget] */ HRESULT ( __stdcall *get_BackColor )( 
            _Form * This,
            /* [retval][out] */ long *rhs);
        
        DECLSPEC_XFGVIRT(_Form, put_BackColor)
        /* [helpstring][propput] */ HRESULT ( __stdcall *put_BackColor )( 
            _Form * This,
            /* [in] */ long rhs);
        
        DECLSPEC_XFGVIRT(_Form, get_ForeColor)
        /* [helpstring][propget] */ HRESULT ( __stdcall *get_ForeColor )( 
            _Form * This,
            /* [retval][out] */ long *rhs);
        
        DECLSPEC_XFGVIRT(_Form, put_ForeColor)
        /* [helpstring][propput] */ HRESULT ( __stdcall *put_ForeColor )( 
            _Form * This,
            /* [in] */ long rhs);
        
        DECLSPEC_XFGVIRT(_Form, get_Left)
        /* [helpstring][propget] */ HRESULT ( __stdcall *get_Left )( 
            _Form * This,
            /* [retval][out] */ long *rhs);
        
        DECLSPEC_XFGVIRT(_Form, put_Left)
        /* [helpstring][propput] */ HRESULT ( __stdcall *put_Left )( 
            _Form * This,
            /* [in] */ long rhs);
        
        DECLSPEC_XFGVIRT(_Form, get_Top)
        /* [helpstring][propget] */ HRESULT ( __stdcall *get_Top )( 
            _Form * This,
            /* [retval][out] */ long *rhs);
        
        DECLSPEC_XFGVIRT(_Form, put_Top)
        /* [helpstring][propput] */ HRESULT ( __stdcall *put_Top )( 
            _Form * This,
            /* [in] */ long rhs);
        
        DECLSPEC_XFGVIRT(_Form, get_Width)
        /* [helpstring][propget] */ HRESULT ( __stdcall *get_Width )( 
            _Form * This,
            /* [retval][out] */ long *rhs);
        
        DECLSPEC_XFGVIRT(_Form, put_Width)
        /* [helpstring][propput] */ HRESULT ( __stdcall *put_Width )( 
            _Form * This,
            /* [in] */ long rhs);
        
        DECLSPEC_XFGVIRT(_Form, get_Height)
        /* [helpstring][propget] */ HRESULT ( __stdcall *get_Height )( 
            _Form * This,
            /* [retval][out] */ long *rhs);
        
        DECLSPEC_XFGVIRT(_Form, put_Height)
        /* [helpstring][propput] */ HRESULT ( __stdcall *put_Height )( 
            _Form * This,
            /* [in] */ long rhs);
        
        DECLSPEC_XFGVIRT(_Form, get_Enabled)
        /* [helpstring][propget] */ HRESULT ( __stdcall *get_Enabled )( 
            _Form * This,
            /* [retval][out] */ VARIANT_BOOL *rhs);
        
        DECLSPEC_XFGVIRT(_Form, put_Enabled)
        /* [helpstring][propput] */ HRESULT ( __stdcall *put_Enabled )( 
            _Form * This,
            /* [in] */ VARIANT_BOOL rhs);
        
        DECLSPEC_XFGVIRT(_Form, get_WindowState)
        /* [helpstring][propget] */ HRESULT ( __stdcall *get_WindowState )( 
            _Form * This,
            /* [retval][out] */ long *rhs);
        
        DECLSPEC_XFGVIRT(_Form, put_WindowState)
        /* [helpstring][propput] */ HRESULT ( __stdcall *put_WindowState )( 
            _Form * This,
            /* [in] */ long rhs);
        
        DECLSPEC_XFGVIRT(_Form, get_MousePointer)
        /* [helpstring][propget] */ HRESULT ( __stdcall *get_MousePointer )( 
            _Form * This,
            /* [retval][out] */ long *rhs);
        
        DECLSPEC_XFGVIRT(_Form, put_MousePointer)
        /* [helpstring][propput] */ HRESULT ( __stdcall *put_MousePointer )( 
            _Form * This,
            /* [in] */ long rhs);
        
        DECLSPEC_XFGVIRT(_Form, get_FontName)
        /* [helpstring][propget] */ HRESULT ( __stdcall *get_FontName )( 
            _Form * This,
            /* [retval][out] */ BSTR *rhs);
        
        DECLSPEC_XFGVIRT(_Form, put_FontName)
        /* [helpstring][propput] */ HRESULT ( __stdcall *put_FontName )( 
            _Form * This,
            /* [in] */ BSTR rhs);
        
        DECLSPEC_XFGVIRT(_Form, get_FontSize)
        /* [helpstring][propget] */ HRESULT ( __stdcall *get_FontSize )( 
            _Form * This,
            /* [retval][out] */ float *rhs);
        
        DECLSPEC_XFGVIRT(_Form, put_FontSize)
        /* [helpstring][propput] */ HRESULT ( __stdcall *put_FontSize )( 
            _Form * This,
            /* [in] */ float rhs);
        
        DECLSPEC_XFGVIRT(_Form, get_FontBold)
        /* [helpstring][propget] */ HRESULT ( __stdcall *get_FontBold )( 
            _Form * This,
            /* [retval][out] */ VARIANT_BOOL *rhs);
        
        DECLSPEC_XFGVIRT(_Form, put_FontBold)
        /* [helpstring][propput] */ HRESULT ( __stdcall *put_FontBold )( 
            _Form * This,
            /* [in] */ VARIANT_BOOL rhs);
        
        DECLSPEC_XFGVIRT(_Form, get_FontItalic)
        /* [helpstring][propget] */ HRESULT ( __stdcall *get_FontItalic )( 
            _Form * This,
            /* [retval][out] */ VARIANT_BOOL *rhs);
        
        DECLSPEC_XFGVIRT(_Form, put_FontItalic)
        /* [helpstring][propput] */ HRESULT ( __stdcall *put_FontItalic )( 
            _Form * This,
            /* [in] */ VARIANT_BOOL rhs);
        
        DECLSPEC_XFGVIRT(_Form, get_FontStrikethru)
        /* [helpstring][propget] */ HRESULT ( __stdcall *get_FontStrikethru )( 
            _Form * This,
            /* [retval][out] */ VARIANT_BOOL *rhs);
        
        DECLSPEC_XFGVIRT(_Form, put_FontStrikethru)
        /* [helpstring][propput] */ HRESULT ( __stdcall *put_FontStrikethru )( 
            _Form * This,
            /* [in] */ VARIANT_BOOL rhs);
        
        DECLSPEC_XFGVIRT(_Form, get_FontUnderline)
        /* [helpstring][propget] */ HRESULT ( __stdcall *get_FontUnderline )( 
            _Form * This,
            /* [retval][out] */ VARIANT_BOOL *rhs);
        
        DECLSPEC_XFGVIRT(_Form, put_FontUnderline)
        /* [helpstring][propput] */ HRESULT ( __stdcall *put_FontUnderline )( 
            _Form * This,
            /* [in] */ VARIANT_BOOL rhs);
        
        DECLSPEC_XFGVIRT(_Form, get_hDC)
        /* [helpstring][propget] */ HRESULT ( __stdcall *get_hDC )( 
            _Form * This,
            /* [retval][out] */ long *rhs);
        
        DECLSPEC_XFGVIRT(_Form, Missing14)
        /* [helpstring] */ HRESULT ( __stdcall *Missing14 )( 
            _Form * This);
        
        DECLSPEC_XFGVIRT(_Form, get_CurrentX)
        /* [helpstring][propget] */ HRESULT ( __stdcall *get_CurrentX )( 
            _Form * This,
            /* [retval][out] */ float *rhs);
        
        DECLSPEC_XFGVIRT(_Form, put_CurrentX)
        /* [helpstring][propput] */ HRESULT ( __stdcall *put_CurrentX )( 
            _Form * This,
            /* [in] */ float rhs);
        
        DECLSPEC_XFGVIRT(_Form, get_CurrentY)
        /* [helpstring][propget] */ HRESULT ( __stdcall *get_CurrentY )( 
            _Form * This,
            /* [retval][out] */ float *rhs);
        
        DECLSPEC_XFGVIRT(_Form, put_CurrentY)
        /* [helpstring][propput] */ HRESULT ( __stdcall *put_CurrentY )( 
            _Form * This,
            /* [in] */ float rhs);
        
        DECLSPEC_XFGVIRT(_Form, get_ScaleLeft)
        /* [helpstring][propget] */ HRESULT ( __stdcall *get_ScaleLeft )( 
            _Form * This,
            /* [retval][out] */ float *rhs);
        
        DECLSPEC_XFGVIRT(_Form, put_ScaleLeft)
        /* [helpstring][propput] */ HRESULT ( __stdcall *put_ScaleLeft )( 
            _Form * This,
            /* [in] */ float rhs);
        
        DECLSPEC_XFGVIRT(_Form, get_ScaleTop)
        /* [helpstring][propget] */ HRESULT ( __stdcall *get_ScaleTop )( 
            _Form * This,
            /* [retval][out] */ float *rhs);
        
        DECLSPEC_XFGVIRT(_Form, put_ScaleTop)
        /* [helpstring][propput] */ HRESULT ( __stdcall *put_ScaleTop )( 
            _Form * This,
            /* [in] */ float rhs);
        
        DECLSPEC_XFGVIRT(_Form, get_ScaleWidth)
        /* [helpstring][propget] */ HRESULT ( __stdcall *get_ScaleWidth )( 
            _Form * This,
            /* [retval][out] */ float *rhs);
        
        DECLSPEC_XFGVIRT(_Form, put_ScaleWidth)
        /* [helpstring][propput] */ HRESULT ( __stdcall *put_ScaleWidth )( 
            _Form * This,
            /* [in] */ float rhs);
        
        DECLSPEC_XFGVIRT(_Form, get_ScaleHeight)
        /* [helpstring][propget] */ HRESULT ( __stdcall *get_ScaleHeight )( 
            _Form * This,
            /* [retval][out] */ float *rhs);
        
        DECLSPEC_XFGVIRT(_Form, put_ScaleHeight)
        /* [helpstring][propput] */ HRESULT ( __stdcall *put_ScaleHeight )( 
            _Form * This,
            /* [in] */ float rhs);
        
        DECLSPEC_XFGVIRT(_Form, get_ScaleMode)
        /* [helpstring][propget] */ HRESULT ( __stdcall *get_ScaleMode )( 
            _Form * This,
            /* [retval][out] */ short *rhs);
        
        DECLSPEC_XFGVIRT(_Form, put_ScaleMode)
        /* [helpstring][propput] */ HRESULT ( __stdcall *put_ScaleMode )( 
            _Form * This,
            /* [in] */ short rhs);
        
        DECLSPEC_XFGVIRT(_Form, get_FontTransparent)
        /* [helpstring][propget] */ HRESULT ( __stdcall *get_FontTransparent )( 
            _Form * This,
            /* [retval][out] */ VARIANT_BOOL *rhs);
        
        DECLSPEC_XFGVIRT(_Form, put_FontTransparent)
        /* [helpstring][propput] */ HRESULT ( __stdcall *put_FontTransparent )( 
            _Form * This,
            /* [in] */ VARIANT_BOOL rhs);
        
        DECLSPEC_XFGVIRT(_Form, get_DrawStyle)
        /* [helpstring][propget] */ HRESULT ( __stdcall *get_DrawStyle )( 
            _Form * This,
            /* [retval][out] */ short *rhs);
        
        DECLSPEC_XFGVIRT(_Form, put_DrawStyle)
        /* [helpstring][propput] */ HRESULT ( __stdcall *put_DrawStyle )( 
            _Form * This,
            /* [in] */ short rhs);
        
        DECLSPEC_XFGVIRT(_Form, get_DrawWidth)
        /* [helpstring][propget] */ HRESULT ( __stdcall *get_DrawWidth )( 
            _Form * This,
            /* [retval][out] */ short *rhs);
        
        DECLSPEC_XFGVIRT(_Form, put_DrawWidth)
        /* [helpstring][propput] */ HRESULT ( __stdcall *put_DrawWidth )( 
            _Form * This,
            /* [in] */ short rhs);
        
        DECLSPEC_XFGVIRT(_Form, get_FillStyle)
        /* [helpstring][propget] */ HRESULT ( __stdcall *get_FillStyle )( 
            _Form * This,
            /* [retval][out] */ short *rhs);
        
        DECLSPEC_XFGVIRT(_Form, put_FillStyle)
        /* [helpstring][propput] */ HRESULT ( __stdcall *put_FillStyle )( 
            _Form * This,
            /* [in] */ short rhs);
        
        DECLSPEC_XFGVIRT(_Form, get_FillColor)
        /* [helpstring][propget] */ HRESULT ( __stdcall *get_FillColor )( 
            _Form * This,
            /* [retval][out] */ long *rhs);
        
        DECLSPEC_XFGVIRT(_Form, put_FillColor)
        /* [helpstring][propput] */ HRESULT ( __stdcall *put_FillColor )( 
            _Form * This,
            /* [in] */ long rhs);
        
        DECLSPEC_XFGVIRT(_Form, get_DrawMode)
        /* [helpstring][propget] */ HRESULT ( __stdcall *get_DrawMode )( 
            _Form * This,
            /* [retval][out] */ short *rhs);
        
        DECLSPEC_XFGVIRT(_Form, put_DrawMode)
        /* [helpstring][propput] */ HRESULT ( __stdcall *put_DrawMode )( 
            _Form * This,
            /* [in] */ short rhs);
        
        DECLSPEC_XFGVIRT(_Form, get_AutoRedraw)
        /* [helpstring][propget] */ HRESULT ( __stdcall *get_AutoRedraw )( 
            _Form * This,
            /* [retval][out] */ VARIANT_BOOL *rhs);
        
        DECLSPEC_XFGVIRT(_Form, put_AutoRedraw)
        /* [helpstring][propput] */ HRESULT ( __stdcall *put_AutoRedraw )( 
            _Form * This,
            /* [in] */ VARIANT_BOOL rhs);
        
        DECLSPEC_XFGVIRT(_Form, get_Picture)
        /* [helpstring][propget] */ HRESULT ( __stdcall *get_Picture )( 
            _Form * This,
            /* [retval][out] */ long *rhs);
        
        DECLSPEC_XFGVIRT(_Form, put_Picture)
        /* [helpstring][propput] */ HRESULT ( __stdcall *put_Picture )( 
            _Form * This,
            /* [in] */ long rhs);
        
        DECLSPEC_XFGVIRT(_Form, get_BorderStyle)
        /* [helpstring][propget] */ HRESULT ( __stdcall *get_BorderStyle )( 
            _Form * This,
            /* [retval][out] */ short *rhs);
        
        DECLSPEC_XFGVIRT(_Form, put_BorderStyle)
        /* [helpstring][propput] */ HRESULT ( __stdcall *put_BorderStyle )( 
            _Form * This,
            /* [in] */ short rhs);
        
        DECLSPEC_XFGVIRT(_Form, get_Icon)
        /* [helpstring][propget] */ HRESULT ( __stdcall *get_Icon )( 
            _Form * This,
            /* [retval][out] */ long *rhs);
        
        DECLSPEC_XFGVIRT(_Form, put_Icon)
        /* [helpstring][propput] */ HRESULT ( __stdcall *put_Icon )( 
            _Form * This,
            /* [in] */ long rhs);
        
        DECLSPEC_XFGVIRT(_Form, get_LinkTopic)
        /* [helpstring][propget] */ HRESULT ( __stdcall *get_LinkTopic )( 
            _Form * This,
            /* [retval][out] */ BSTR *rhs);
        
        DECLSPEC_XFGVIRT(_Form, put_LinkTopic)
        /* [helpstring][propput] */ HRESULT ( __stdcall *put_LinkTopic )( 
            _Form * This,
            /* [in] */ BSTR rhs);
        
        DECLSPEC_XFGVIRT(_Form, get_LinkMode)
        /* [helpstring][propget] */ HRESULT ( __stdcall *get_LinkMode )( 
            _Form * This,
            /* [retval][out] */ short *rhs);
        
        DECLSPEC_XFGVIRT(_Form, put_LinkMode)
        /* [helpstring][propput] */ HRESULT ( __stdcall *put_LinkMode )( 
            _Form * This,
            /* [in] */ short rhs);
        
        DECLSPEC_XFGVIRT(_Form, get_MaxButton)
        /* [helpstring][propget] */ HRESULT ( __stdcall *get_MaxButton )( 
            _Form * This,
            /* [retval][out] */ VARIANT_BOOL *rhs);
        
        DECLSPEC_XFGVIRT(_Form, Missing15)
        /* [helpstring] */ HRESULT ( __stdcall *Missing15 )( 
            _Form * This);
        
        DECLSPEC_XFGVIRT(_Form, get_MinButton)
        /* [helpstring][propget] */ HRESULT ( __stdcall *get_MinButton )( 
            _Form * This,
            /* [retval][out] */ VARIANT_BOOL *rhs);
        
        DECLSPEC_XFGVIRT(_Form, Missing16)
        /* [helpstring] */ HRESULT ( __stdcall *Missing16 )( 
            _Form * This);
        
        DECLSPEC_XFGVIRT(_Form, get_ControlBox)
        /* [helpstring][propget] */ HRESULT ( __stdcall *get_ControlBox )( 
            _Form * This,
            /* [retval][out] */ VARIANT_BOOL *rhs);
        
        DECLSPEC_XFGVIRT(_Form, Missing17)
        /* [helpstring] */ HRESULT ( __stdcall *Missing17 )( 
            _Form * This);
        
        DECLSPEC_XFGVIRT(_Form, get_Image)
        /* [helpstring][propget] */ HRESULT ( __stdcall *get_Image )( 
            _Form * This,
            /* [retval][out] */ long *rhs);
        
        DECLSPEC_XFGVIRT(_Form, Missing18)
        /* [helpstring] */ HRESULT ( __stdcall *Missing18 )( 
            _Form * This);
        
        DECLSPEC_XFGVIRT(_Form, get_HasDC)
        /* [helpstring][propget] */ HRESULT ( __stdcall *get_HasDC )( 
            _Form * This,
            /* [retval][out] */ VARIANT_BOOL *rhs);
        
        DECLSPEC_XFGVIRT(_Form, Missing19)
        /* [helpstring] */ HRESULT ( __stdcall *Missing19 )( 
            _Form * This);
        
        DECLSPEC_XFGVIRT(_Form, Missing20)
        /* [helpstring] */ HRESULT ( __stdcall *Missing20 )( 
            _Form * This);
        
        DECLSPEC_XFGVIRT(_Form, Missing21)
        /* [helpstring] */ HRESULT ( __stdcall *Missing21 )( 
            _Form * This);
        
        DECLSPEC_XFGVIRT(_Form, Missing22)
        /* [helpstring] */ HRESULT ( __stdcall *Missing22 )( 
            _Form * This);
        
        DECLSPEC_XFGVIRT(_Form, Missing23)
        /* [helpstring] */ HRESULT ( __stdcall *Missing23 )( 
            _Form * This);
        
        DECLSPEC_XFGVIRT(_Form, Missing24)
        /* [helpstring] */ HRESULT ( __stdcall *Missing24 )( 
            _Form * This);
        
        DECLSPEC_XFGVIRT(_Form, Missing25)
        /* [helpstring] */ HRESULT ( __stdcall *Missing25 )( 
            _Form * This);
        
        DECLSPEC_XFGVIRT(_Form, get_Visible)
        /* [helpstring][propget] */ HRESULT ( __stdcall *get_Visible )( 
            _Form * This,
            /* [retval][out] */ VARIANT_BOOL *rhs);
        
        DECLSPEC_XFGVIRT(_Form, put_Visible)
        /* [helpstring][propput] */ HRESULT ( __stdcall *put_Visible )( 
            _Form * This,
            /* [in] */ VARIANT_BOOL rhs);
        
        DECLSPEC_XFGVIRT(_Form, get_Tag)
        /* [helpstring][propget] */ HRESULT ( __stdcall *get_Tag )( 
            _Form * This,
            /* [retval][out] */ BSTR *rhs);
        
        DECLSPEC_XFGVIRT(_Form, put_Tag)
        /* [helpstring][propput] */ HRESULT ( __stdcall *put_Tag )( 
            _Form * This,
            /* [in] */ BSTR rhs);
        
        DECLSPEC_XFGVIRT(_Form, get_MDIChild)
        /* [helpstring][propget] */ HRESULT ( __stdcall *get_MDIChild )( 
            _Form * This,
            /* [retval][out] */ VARIANT_BOOL *rhs);
        
        DECLSPEC_XFGVIRT(_Form, Missing26)
        /* [helpstring] */ HRESULT ( __stdcall *Missing26 )( 
            _Form * This);
        
        DECLSPEC_XFGVIRT(_Form, get_KeyPreview)
        /* [helpstring][propget] */ HRESULT ( __stdcall *get_KeyPreview )( 
            _Form * This,
            /* [retval][out] */ VARIANT_BOOL *rhs);
        
        DECLSPEC_XFGVIRT(_Form, put_KeyPreview)
        /* [helpstring][propput] */ HRESULT ( __stdcall *put_KeyPreview )( 
            _Form * This,
            /* [in] */ VARIANT_BOOL rhs);
        
        DECLSPEC_XFGVIRT(_Form, get_ClipControls)
        /* [helpstring][propget] */ HRESULT ( __stdcall *get_ClipControls )( 
            _Form * This,
            /* [retval][out] */ VARIANT_BOOL *rhs);
        
        DECLSPEC_XFGVIRT(_Form, put_ClipControls)
        /* [helpstring][propput] */ HRESULT ( __stdcall *put_ClipControls )( 
            _Form * This,
            /* [in] */ VARIANT_BOOL rhs);
        
        DECLSPEC_XFGVIRT(_Form, get_HelpContextID)
        /* [helpstring][propget] */ HRESULT ( __stdcall *get_HelpContextID )( 
            _Form * This,
            /* [retval][out] */ long *rhs);
        
        DECLSPEC_XFGVIRT(_Form, put_HelpContextID)
        /* [helpstring][propput] */ HRESULT ( __stdcall *put_HelpContextID )( 
            _Form * This,
            /* [in] */ long rhs);
        
        DECLSPEC_XFGVIRT(_Form, get_ActiveControl)
        /* [helpstring][propget] */ HRESULT ( __stdcall *get_ActiveControl )( 
            _Form * This,
            /* [retval][out] */ long *rhs);
        
        DECLSPEC_XFGVIRT(_Form, Missing27)
        /* [helpstring] */ HRESULT ( __stdcall *Missing27 )( 
            _Form * This);
        
        DECLSPEC_XFGVIRT(_Form, Missing28)
        /* [helpstring] */ HRESULT ( __stdcall *Missing28 )( 
            _Form * This);
        
        DECLSPEC_XFGVIRT(_Form, Missing29)
        /* [helpstring] */ HRESULT ( __stdcall *Missing29 )( 
            _Form * This);
        
        DECLSPEC_XFGVIRT(_Form, Missing30)
        /* [helpstring] */ HRESULT ( __stdcall *Missing30 )( 
            _Form * This);
        
        DECLSPEC_XFGVIRT(_Form, Missing31)
        /* [helpstring] */ HRESULT ( __stdcall *Missing31 )( 
            _Form * This);
        
        DECLSPEC_XFGVIRT(_Form, Missing32)
        /* [helpstring] */ HRESULT ( __stdcall *Missing32 )( 
            _Form * This);
        
        DECLSPEC_XFGVIRT(_Form, Missing33)
        /* [helpstring] */ HRESULT ( __stdcall *Missing33 )( 
            _Form * This);
        
        DECLSPEC_XFGVIRT(_Form, Missing34)
        /* [helpstring] */ HRESULT ( __stdcall *Missing34 )( 
            _Form * This);
        
        DECLSPEC_XFGVIRT(_Form, Missing35)
        /* [helpstring] */ HRESULT ( __stdcall *Missing35 )( 
            _Form * This);
        
        DECLSPEC_XFGVIRT(_Form, get_Count)
        /* [helpstring][propget] */ HRESULT ( __stdcall *get_Count )( 
            _Form * This,
            /* [retval][out] */ short *rhs);
        
        DECLSPEC_XFGVIRT(_Form, Missing36)
        /* [helpstring] */ HRESULT ( __stdcall *Missing36 )( 
            _Form * This);
        
        DECLSPEC_XFGVIRT(_Form, get_Controls)
        /* [helpstring][propget] */ HRESULT ( __stdcall *get_Controls )( 
            _Form * This,
            /* [retval][out] */ long *rhs);
        
        DECLSPEC_XFGVIRT(_Form, Missing37)
        /* [helpstring] */ HRESULT ( __stdcall *Missing37 )( 
            _Form * This);
        
        DECLSPEC_XFGVIRT(_Form, get_MouseIcon)
        /* [helpstring][propget] */ HRESULT ( __stdcall *get_MouseIcon )( 
            _Form * This,
            /* [retval][out] */ long *rhs);
        
        DECLSPEC_XFGVIRT(_Form, put_MouseIcon)
        /* [helpstring][propput] */ HRESULT ( __stdcall *put_MouseIcon )( 
            _Form * This,
            /* [in] */ long rhs);
        
        DECLSPEC_XFGVIRT(_Form, Missing38)
        /* [helpstring] */ HRESULT ( __stdcall *Missing38 )( 
            _Form * This);
        
        DECLSPEC_XFGVIRT(_Form, Missing39)
        /* [helpstring] */ HRESULT ( __stdcall *Missing39 )( 
            _Form * This);
        
        DECLSPEC_XFGVIRT(_Form, Missing40)
        /* [helpstring] */ HRESULT ( __stdcall *Missing40 )( 
            _Form * This);
        
        DECLSPEC_XFGVIRT(_Form, Missing41)
        /* [helpstring] */ HRESULT ( __stdcall *Missing41 )( 
            _Form * This);
        
        DECLSPEC_XFGVIRT(_Form, Missing42)
        /* [helpstring] */ HRESULT ( __stdcall *Missing42 )( 
            _Form * This);
        
        DECLSPEC_XFGVIRT(_Form, Missing43)
        /* [helpstring] */ HRESULT ( __stdcall *Missing43 )( 
            _Form * This);
        
        DECLSPEC_XFGVIRT(_Form, Missing44)
        /* [helpstring] */ HRESULT ( __stdcall *Missing44 )( 
            _Form * This);
        
        DECLSPEC_XFGVIRT(_Form, Missing45)
        /* [helpstring] */ HRESULT ( __stdcall *Missing45 )( 
            _Form * This);
        
        DECLSPEC_XFGVIRT(_Form, get_Font)
        /* [helpstring][propget] */ HRESULT ( __stdcall *get_Font )( 
            _Form * This,
            /* [retval][out] */ long *rhs);
        
        DECLSPEC_XFGVIRT(_Form, putref_Font)
        /* [helpstring][propputref] */ HRESULT ( __stdcall *putref_Font )( 
            _Form * This,
            /* [in] */ long rhs);
        
        DECLSPEC_XFGVIRT(_Form, get_Appearance)
        /* [helpstring][propget] */ HRESULT ( __stdcall *get_Appearance )( 
            _Form * This,
            /* [retval][out] */ short *rhs);
        
        DECLSPEC_XFGVIRT(_Form, put_Appearance)
        /* [helpstring][propput] */ HRESULT ( __stdcall *put_Appearance )( 
            _Form * This,
            /* [in] */ short rhs);
        
        DECLSPEC_XFGVIRT(_Form, get_WhatsThisButton)
        /* [helpstring][propget] */ HRESULT ( __stdcall *get_WhatsThisButton )( 
            _Form * This,
            /* [retval][out] */ VARIANT_BOOL *rhs);
        
        DECLSPEC_XFGVIRT(_Form, Missing46)
        /* [helpstring] */ HRESULT ( __stdcall *Missing46 )( 
            _Form * This);
        
        DECLSPEC_XFGVIRT(_Form, get_WhatsThisHelp)
        /* [helpstring][propget] */ HRESULT ( __stdcall *get_WhatsThisHelp )( 
            _Form * This,
            /* [retval][out] */ VARIANT_BOOL *rhs);
        
        DECLSPEC_XFGVIRT(_Form, Missing47)
        /* [helpstring] */ HRESULT ( __stdcall *Missing47 )( 
            _Form * This);
        
        DECLSPEC_XFGVIRT(_Form, get_ShowInTaskbar)
        /* [helpstring][propget] */ HRESULT ( __stdcall *get_ShowInTaskbar )( 
            _Form * This,
            /* [retval][out] */ VARIANT_BOOL *rhs);
        
        DECLSPEC_XFGVIRT(_Form, Missing48)
        /* [helpstring] */ HRESULT ( __stdcall *Missing48 )( 
            _Form * This);
        
        DECLSPEC_XFGVIRT(_Form, get_RightToLeft)
        /* [helpstring][propget] */ HRESULT ( __stdcall *get_RightToLeft )( 
            _Form * This,
            /* [retval][out] */ VARIANT_BOOL *rhs);
        
        DECLSPEC_XFGVIRT(_Form, put_RightToLeft)
        /* [helpstring][propput] */ HRESULT ( __stdcall *put_RightToLeft )( 
            _Form * This,
            /* [in] */ VARIANT_BOOL rhs);
        
        DECLSPEC_XFGVIRT(_Form, get_StartUpPosition)
        /* [helpstring][propget] */ HRESULT ( __stdcall *get_StartUpPosition )( 
            _Form * This,
            /* [retval][out] */ short *rhs);
        
        DECLSPEC_XFGVIRT(_Form, Missing49)
        /* [helpstring] */ HRESULT ( __stdcall *Missing49 )( 
            _Form * This);
        
        DECLSPEC_XFGVIRT(_Form, get_OLEDropMode)
        /* [helpstring][propget] */ HRESULT ( __stdcall *get_OLEDropMode )( 
            _Form * This,
            /* [retval][out] */ short *rhs);
        
        DECLSPEC_XFGVIRT(_Form, put_OLEDropMode)
        /* [helpstring][propput] */ HRESULT ( __stdcall *put_OLEDropMode )( 
            _Form * This,
            /* [in] */ short rhs);
        
        DECLSPEC_XFGVIRT(_Form, get_Palette)
        /* [helpstring][propget] */ HRESULT ( __stdcall *get_Palette )( 
            _Form * This,
            /* [retval][out] */ long *rhs);
        
        DECLSPEC_XFGVIRT(_Form, putref_Palette)
        /* [helpstring][propputref] */ HRESULT ( __stdcall *putref_Palette )( 
            _Form * This,
            /* [in] */ long rhs);
        
        DECLSPEC_XFGVIRT(_Form, get_PaletteMode)
        /* [helpstring][propget] */ HRESULT ( __stdcall *get_PaletteMode )( 
            _Form * This,
            /* [retval][out] */ short *rhs);
        
        DECLSPEC_XFGVIRT(_Form, put_PaletteMode)
        /* [helpstring][propput] */ HRESULT ( __stdcall *put_PaletteMode )( 
            _Form * This,
            /* [in] */ short rhs);
        
        DECLSPEC_XFGVIRT(_Form, get_Moveable)
        /* [helpstring][propget] */ HRESULT ( __stdcall *get_Moveable )( 
            _Form * This,
            /* [retval][out] */ VARIANT_BOOL *rhs);
        
        DECLSPEC_XFGVIRT(_Form, Missing50)
        /* [helpstring] */ HRESULT ( __stdcall *Missing50 )( 
            _Form * This);
        
        DECLSPEC_XFGVIRT(_Form, Refresh)
        /* [helpstring] */ HRESULT ( __stdcall *Refresh )( 
            _Form * This);
        
        DECLSPEC_XFGVIRT(_Form, Move)
        /* [helpstring] */ HRESULT ( __stdcall *Move )( 
            _Form * This,
            /* [in] */ float Left,
            /* [optional][in] */ VARIANT Top,
            /* [optional][in] */ VARIANT Width,
            /* [optional][in] */ VARIANT Height);
        
        DECLSPEC_XFGVIRT(_Form, SetFocus)
        /* [helpstring] */ HRESULT ( __stdcall *SetFocus )( 
            _Form * This);
        
        DECLSPEC_XFGVIRT(_Form, ZOrder)
        /* [helpstring] */ HRESULT ( __stdcall *ZOrder )( 
            _Form * This,
            /* [optional][in] */ VARIANT Position);
        
        DECLSPEC_XFGVIRT(_Form, Show)
        /* [helpstring] */ HRESULT ( __stdcall *Show )( 
            _Form * This,
            /* [in] */ VARIANT modal,
            /* [in] */ VARIANT OwnerForm);
        
        DECLSPEC_XFGVIRT(_Form, Hide)
        /* [helpstring] */ HRESULT ( __stdcall *Hide )( 
            _Form * This);
        
        DECLSPEC_XFGVIRT(_Form, PrintForm)
        /* [helpstring] */ HRESULT ( __stdcall *PrintForm )( 
            _Form * This);
        
        DECLSPEC_XFGVIRT(_Form, PopupMenu)
        /* [helpstring] */ HRESULT ( __stdcall *PopupMenu )( 
            _Form * This,
            /* [in] */ IDispatch *Menu,
            /* [optional][in] */ VARIANT Flags,
            /* [optional][in] */ VARIANT x,
            /* [optional][in] */ VARIANT y,
            /* [optional][in] */ VARIANT BoldCommand);
        
        DECLSPEC_XFGVIRT(_Form, Circle)
        /* [helpstring] */ HRESULT ( __stdcall *Circle )( 
            _Form * This,
            /* [in] */ float x,
            /* [in] */ float y,
            /* [in] */ float Radius,
            /* [optional][in] */ VARIANT Color,
            /* [optional][in] */ VARIANT StartAngle,
            /* [optional][in] */ VARIANT EndAngle,
            /* [optional][in] */ VARIANT Aspect);
        
        DECLSPEC_XFGVIRT(_Form, Cls)
        /* [helpstring] */ HRESULT ( __stdcall *Cls )( 
            _Form * This);
        
        DECLSPEC_XFGVIRT(_Form, Line)
        /* [helpstring] */ HRESULT ( __stdcall *Line )( 
            _Form * This,
            /* [in] */ float x1,
            /* [in] */ float y1,
            /* [in] */ float x2,
            /* [in] */ float y2,
            /* [optional][in] */ VARIANT Color,
            /* [optional][in] */ VARIANT BorF);
        
        DECLSPEC_XFGVIRT(_Form, PaintPicture)
        /* [helpstring] */ HRESULT ( __stdcall *PaintPicture )( 
            _Form * This,
            /* [in] */ IDispatch *Picture,
            /* [in] */ float x1,
            /* [in] */ float y1,
            /* [optional][in] */ VARIANT width1,
            /* [optional][in] */ VARIANT height1,
            /* [optional][in] */ VARIANT x2,
            /* [optional][in] */ VARIANT y2,
            /* [optional][in] */ VARIANT width2,
            /* [optional][in] */ VARIANT height2,
            /* [optional][in] */ VARIANT Opcode);
        
        DECLSPEC_XFGVIRT(_Form, Point)
        /* [helpstring] */ HRESULT ( __stdcall *Point )( 
            _Form * This,
            /* [in] */ float x,
            /* [in] */ float y,
            /* [retval][out] */ long *rhs);
        
        DECLSPEC_XFGVIRT(_Form, PSet)
        /* [helpstring] */ HRESULT ( __stdcall *PSet )( 
            _Form * This,
            /* [in] */ float x,
            /* [in] */ float y,
            /* [optional][in] */ VARIANT Color);
        
        DECLSPEC_XFGVIRT(_Form, Scale)
        /* [helpstring] */ HRESULT ( __stdcall *Scale )( 
            _Form * This,
            /* [optional][in] */ VARIANT x1,
            /* [optional][in] */ VARIANT y1,
            /* [optional][in] */ VARIANT x2,
            /* [optional][in] */ VARIANT y2);
        
        DECLSPEC_XFGVIRT(_Form, ScaleX)
        /* [helpstring] */ HRESULT ( __stdcall *ScaleX )( 
            _Form * This,
            /* [in] */ float Value,
            /* [optional][in] */ VARIANT fromScale,
            /* [optional][in] */ VARIANT toScale,
            /* [retval][out] */ float *rhs);
        
        DECLSPEC_XFGVIRT(_Form, ScaleY)
        /* [helpstring] */ HRESULT ( __stdcall *ScaleY )( 
            _Form * This,
            /* [in] */ float Value,
            /* [optional][in] */ VARIANT fromScale,
            /* [optional][in] */ VARIANT toScale,
            /* [retval][out] */ float *rhs);
        
        DECLSPEC_XFGVIRT(_Form, TextWidth)
        /* [helpstring] */ HRESULT ( __stdcall *TextWidth )( 
            _Form * This,
            /* [in] */ BSTR str,
            /* [retval][out] */ float *rhs);
        
        DECLSPEC_XFGVIRT(_Form, TextHeight)
        /* [helpstring] */ HRESULT ( __stdcall *TextHeight )( 
            _Form * This,
            /* [in] */ BSTR str,
            /* [retval][out] */ float *rhs);
        
        DECLSPEC_XFGVIRT(_Form, ValidateControls)
        /* [helpstring] */ HRESULT ( __stdcall *ValidateControls )( 
            _Form * This,
            /* [retval][out] */ VARIANT_BOOL *rhs);
        
        DECLSPEC_XFGVIRT(_Form, WhatsThisMode)
        /* [helpstring] */ HRESULT ( __stdcall *WhatsThisMode )( 
            _Form * This);
        
        DECLSPEC_XFGVIRT(_Form, OLEDrag)
        /* [helpstring] */ HRESULT ( __stdcall *OLEDrag )( 
            _Form * This);
        
        END_INTERFACE
    } _FormVtbl;

    interface _Form
    {
        CONST_VTBL struct _FormVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define _Form_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define _Form_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define _Form_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define _Form_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define _Form_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define _Form_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define _Form_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define _Form_Missing1(This)	\
    ( (This)->lpVtbl -> Missing1(This) ) 

#define _Form_Missing2(This)	\
    ( (This)->lpVtbl -> Missing2(This) ) 

#define _Form_Missing3(This)	\
    ( (This)->lpVtbl -> Missing3(This) ) 

#define _Form_Missing4(This)	\
    ( (This)->lpVtbl -> Missing4(This) ) 

#define _Form_Missing5(This)	\
    ( (This)->lpVtbl -> Missing5(This) ) 

#define _Form_Missing6(This)	\
    ( (This)->lpVtbl -> Missing6(This) ) 

#define _Form_Missing7(This)	\
    ( (This)->lpVtbl -> Missing7(This) ) 

#define _Form_Missing8(This)	\
    ( (This)->lpVtbl -> Missing8(This) ) 

#define _Form_Missing9(This)	\
    ( (This)->lpVtbl -> Missing9(This) ) 

#define _Form_Missing10(This)	\
    ( (This)->lpVtbl -> Missing10(This) ) 

#define _Form_Missing11(This)	\
    ( (This)->lpVtbl -> Missing11(This) ) 

#define _Form_get_Name(This,rhs)	\
    ( (This)->lpVtbl -> get_Name(This,rhs) ) 

#define _Form_Missing12(This)	\
    ( (This)->lpVtbl -> Missing12(This) ) 

#define _Form_get_Caption(This,rhs)	\
    ( (This)->lpVtbl -> get_Caption(This,rhs) ) 

#define _Form_put_Caption(This,rhs)	\
    ( (This)->lpVtbl -> put_Caption(This,rhs) ) 

#define _Form_get_hWnd(This,rhs)	\
    ( (This)->lpVtbl -> get_hWnd(This,rhs) ) 

#define _Form_Missing13(This)	\
    ( (This)->lpVtbl -> Missing13(This) ) 

#define _Form_get_BackColor(This,rhs)	\
    ( (This)->lpVtbl -> get_BackColor(This,rhs) ) 

#define _Form_put_BackColor(This,rhs)	\
    ( (This)->lpVtbl -> put_BackColor(This,rhs) ) 

#define _Form_get_ForeColor(This,rhs)	\
    ( (This)->lpVtbl -> get_ForeColor(This,rhs) ) 

#define _Form_put_ForeColor(This,rhs)	\
    ( (This)->lpVtbl -> put_ForeColor(This,rhs) ) 

#define _Form_get_Left(This,rhs)	\
    ( (This)->lpVtbl -> get_Left(This,rhs) ) 

#define _Form_put_Left(This,rhs)	\
    ( (This)->lpVtbl -> put_Left(This,rhs) ) 

#define _Form_get_Top(This,rhs)	\
    ( (This)->lpVtbl -> get_Top(This,rhs) ) 

#define _Form_put_Top(This,rhs)	\
    ( (This)->lpVtbl -> put_Top(This,rhs) ) 

#define _Form_get_Width(This,rhs)	\
    ( (This)->lpVtbl -> get_Width(This,rhs) ) 

#define _Form_put_Width(This,rhs)	\
    ( (This)->lpVtbl -> put_Width(This,rhs) ) 

#define _Form_get_Height(This,rhs)	\
    ( (This)->lpVtbl -> get_Height(This,rhs) ) 

#define _Form_put_Height(This,rhs)	\
    ( (This)->lpVtbl -> put_Height(This,rhs) ) 

#define _Form_get_Enabled(This,rhs)	\
    ( (This)->lpVtbl -> get_Enabled(This,rhs) ) 

#define _Form_put_Enabled(This,rhs)	\
    ( (This)->lpVtbl -> put_Enabled(This,rhs) ) 

#define _Form_get_WindowState(This,rhs)	\
    ( (This)->lpVtbl -> get_WindowState(This,rhs) ) 

#define _Form_put_WindowState(This,rhs)	\
    ( (This)->lpVtbl -> put_WindowState(This,rhs) ) 

#define _Form_get_MousePointer(This,rhs)	\
    ( (This)->lpVtbl -> get_MousePointer(This,rhs) ) 

#define _Form_put_MousePointer(This,rhs)	\
    ( (This)->lpVtbl -> put_MousePointer(This,rhs) ) 

#define _Form_get_FontName(This,rhs)	\
    ( (This)->lpVtbl -> get_FontName(This,rhs) ) 

#define _Form_put_FontName(This,rhs)	\
    ( (This)->lpVtbl -> put_FontName(This,rhs) ) 

#define _Form_get_FontSize(This,rhs)	\
    ( (This)->lpVtbl -> get_FontSize(This,rhs) ) 

#define _Form_put_FontSize(This,rhs)	\
    ( (This)->lpVtbl -> put_FontSize(This,rhs) ) 

#define _Form_get_FontBold(This,rhs)	\
    ( (This)->lpVtbl -> get_FontBold(This,rhs) ) 

#define _Form_put_FontBold(This,rhs)	\
    ( (This)->lpVtbl -> put_FontBold(This,rhs) ) 

#define _Form_get_FontItalic(This,rhs)	\
    ( (This)->lpVtbl -> get_FontItalic(This,rhs) ) 

#define _Form_put_FontItalic(This,rhs)	\
    ( (This)->lpVtbl -> put_FontItalic(This,rhs) ) 

#define _Form_get_FontStrikethru(This,rhs)	\
    ( (This)->lpVtbl -> get_FontStrikethru(This,rhs) ) 

#define _Form_put_FontStrikethru(This,rhs)	\
    ( (This)->lpVtbl -> put_FontStrikethru(This,rhs) ) 

#define _Form_get_FontUnderline(This,rhs)	\
    ( (This)->lpVtbl -> get_FontUnderline(This,rhs) ) 

#define _Form_put_FontUnderline(This,rhs)	\
    ( (This)->lpVtbl -> put_FontUnderline(This,rhs) ) 

#define _Form_get_hDC(This,rhs)	\
    ( (This)->lpVtbl -> get_hDC(This,rhs) ) 

#define _Form_Missing14(This)	\
    ( (This)->lpVtbl -> Missing14(This) ) 

#define _Form_get_CurrentX(This,rhs)	\
    ( (This)->lpVtbl -> get_CurrentX(This,rhs) ) 

#define _Form_put_CurrentX(This,rhs)	\
    ( (This)->lpVtbl -> put_CurrentX(This,rhs) ) 

#define _Form_get_CurrentY(This,rhs)	\
    ( (This)->lpVtbl -> get_CurrentY(This,rhs) ) 

#define _Form_put_CurrentY(This,rhs)	\
    ( (This)->lpVtbl -> put_CurrentY(This,rhs) ) 

#define _Form_get_ScaleLeft(This,rhs)	\
    ( (This)->lpVtbl -> get_ScaleLeft(This,rhs) ) 

#define _Form_put_ScaleLeft(This,rhs)	\
    ( (This)->lpVtbl -> put_ScaleLeft(This,rhs) ) 

#define _Form_get_ScaleTop(This,rhs)	\
    ( (This)->lpVtbl -> get_ScaleTop(This,rhs) ) 

#define _Form_put_ScaleTop(This,rhs)	\
    ( (This)->lpVtbl -> put_ScaleTop(This,rhs) ) 

#define _Form_get_ScaleWidth(This,rhs)	\
    ( (This)->lpVtbl -> get_ScaleWidth(This,rhs) ) 

#define _Form_put_ScaleWidth(This,rhs)	\
    ( (This)->lpVtbl -> put_ScaleWidth(This,rhs) ) 

#define _Form_get_ScaleHeight(This,rhs)	\
    ( (This)->lpVtbl -> get_ScaleHeight(This,rhs) ) 

#define _Form_put_ScaleHeight(This,rhs)	\
    ( (This)->lpVtbl -> put_ScaleHeight(This,rhs) ) 

#define _Form_get_ScaleMode(This,rhs)	\
    ( (This)->lpVtbl -> get_ScaleMode(This,rhs) ) 

#define _Form_put_ScaleMode(This,rhs)	\
    ( (This)->lpVtbl -> put_ScaleMode(This,rhs) ) 

#define _Form_get_FontTransparent(This,rhs)	\
    ( (This)->lpVtbl -> get_FontTransparent(This,rhs) ) 

#define _Form_put_FontTransparent(This,rhs)	\
    ( (This)->lpVtbl -> put_FontTransparent(This,rhs) ) 

#define _Form_get_DrawStyle(This,rhs)	\
    ( (This)->lpVtbl -> get_DrawStyle(This,rhs) ) 

#define _Form_put_DrawStyle(This,rhs)	\
    ( (This)->lpVtbl -> put_DrawStyle(This,rhs) ) 

#define _Form_get_DrawWidth(This,rhs)	\
    ( (This)->lpVtbl -> get_DrawWidth(This,rhs) ) 

#define _Form_put_DrawWidth(This,rhs)	\
    ( (This)->lpVtbl -> put_DrawWidth(This,rhs) ) 

#define _Form_get_FillStyle(This,rhs)	\
    ( (This)->lpVtbl -> get_FillStyle(This,rhs) ) 

#define _Form_put_FillStyle(This,rhs)	\
    ( (This)->lpVtbl -> put_FillStyle(This,rhs) ) 

#define _Form_get_FillColor(This,rhs)	\
    ( (This)->lpVtbl -> get_FillColor(This,rhs) ) 

#define _Form_put_FillColor(This,rhs)	\
    ( (This)->lpVtbl -> put_FillColor(This,rhs) ) 

#define _Form_get_DrawMode(This,rhs)	\
    ( (This)->lpVtbl -> get_DrawMode(This,rhs) ) 

#define _Form_put_DrawMode(This,rhs)	\
    ( (This)->lpVtbl -> put_DrawMode(This,rhs) ) 

#define _Form_get_AutoRedraw(This,rhs)	\
    ( (This)->lpVtbl -> get_AutoRedraw(This,rhs) ) 

#define _Form_put_AutoRedraw(This,rhs)	\
    ( (This)->lpVtbl -> put_AutoRedraw(This,rhs) ) 

#define _Form_get_Picture(This,rhs)	\
    ( (This)->lpVtbl -> get_Picture(This,rhs) ) 

#define _Form_put_Picture(This,rhs)	\
    ( (This)->lpVtbl -> put_Picture(This,rhs) ) 

#define _Form_get_BorderStyle(This,rhs)	\
    ( (This)->lpVtbl -> get_BorderStyle(This,rhs) ) 

#define _Form_put_BorderStyle(This,rhs)	\
    ( (This)->lpVtbl -> put_BorderStyle(This,rhs) ) 

#define _Form_get_Icon(This,rhs)	\
    ( (This)->lpVtbl -> get_Icon(This,rhs) ) 

#define _Form_put_Icon(This,rhs)	\
    ( (This)->lpVtbl -> put_Icon(This,rhs) ) 

#define _Form_get_LinkTopic(This,rhs)	\
    ( (This)->lpVtbl -> get_LinkTopic(This,rhs) ) 

#define _Form_put_LinkTopic(This,rhs)	\
    ( (This)->lpVtbl -> put_LinkTopic(This,rhs) ) 

#define _Form_get_LinkMode(This,rhs)	\
    ( (This)->lpVtbl -> get_LinkMode(This,rhs) ) 

#define _Form_put_LinkMode(This,rhs)	\
    ( (This)->lpVtbl -> put_LinkMode(This,rhs) ) 

#define _Form_get_MaxButton(This,rhs)	\
    ( (This)->lpVtbl -> get_MaxButton(This,rhs) ) 

#define _Form_Missing15(This)	\
    ( (This)->lpVtbl -> Missing15(This) ) 

#define _Form_get_MinButton(This,rhs)	\
    ( (This)->lpVtbl -> get_MinButton(This,rhs) ) 

#define _Form_Missing16(This)	\
    ( (This)->lpVtbl -> Missing16(This) ) 

#define _Form_get_ControlBox(This,rhs)	\
    ( (This)->lpVtbl -> get_ControlBox(This,rhs) ) 

#define _Form_Missing17(This)	\
    ( (This)->lpVtbl -> Missing17(This) ) 

#define _Form_get_Image(This,rhs)	\
    ( (This)->lpVtbl -> get_Image(This,rhs) ) 

#define _Form_Missing18(This)	\
    ( (This)->lpVtbl -> Missing18(This) ) 

#define _Form_get_HasDC(This,rhs)	\
    ( (This)->lpVtbl -> get_HasDC(This,rhs) ) 

#define _Form_Missing19(This)	\
    ( (This)->lpVtbl -> Missing19(This) ) 

#define _Form_Missing20(This)	\
    ( (This)->lpVtbl -> Missing20(This) ) 

#define _Form_Missing21(This)	\
    ( (This)->lpVtbl -> Missing21(This) ) 

#define _Form_Missing22(This)	\
    ( (This)->lpVtbl -> Missing22(This) ) 

#define _Form_Missing23(This)	\
    ( (This)->lpVtbl -> Missing23(This) ) 

#define _Form_Missing24(This)	\
    ( (This)->lpVtbl -> Missing24(This) ) 

#define _Form_Missing25(This)	\
    ( (This)->lpVtbl -> Missing25(This) ) 

#define _Form_get_Visible(This,rhs)	\
    ( (This)->lpVtbl -> get_Visible(This,rhs) ) 

#define _Form_put_Visible(This,rhs)	\
    ( (This)->lpVtbl -> put_Visible(This,rhs) ) 

#define _Form_get_Tag(This,rhs)	\
    ( (This)->lpVtbl -> get_Tag(This,rhs) ) 

#define _Form_put_Tag(This,rhs)	\
    ( (This)->lpVtbl -> put_Tag(This,rhs) ) 

#define _Form_get_MDIChild(This,rhs)	\
    ( (This)->lpVtbl -> get_MDIChild(This,rhs) ) 

#define _Form_Missing26(This)	\
    ( (This)->lpVtbl -> Missing26(This) ) 

#define _Form_get_KeyPreview(This,rhs)	\
    ( (This)->lpVtbl -> get_KeyPreview(This,rhs) ) 

#define _Form_put_KeyPreview(This,rhs)	\
    ( (This)->lpVtbl -> put_KeyPreview(This,rhs) ) 

#define _Form_get_ClipControls(This,rhs)	\
    ( (This)->lpVtbl -> get_ClipControls(This,rhs) ) 

#define _Form_put_ClipControls(This,rhs)	\
    ( (This)->lpVtbl -> put_ClipControls(This,rhs) ) 

#define _Form_get_HelpContextID(This,rhs)	\
    ( (This)->lpVtbl -> get_HelpContextID(This,rhs) ) 

#define _Form_put_HelpContextID(This,rhs)	\
    ( (This)->lpVtbl -> put_HelpContextID(This,rhs) ) 

#define _Form_get_ActiveControl(This,rhs)	\
    ( (This)->lpVtbl -> get_ActiveControl(This,rhs) ) 

#define _Form_Missing27(This)	\
    ( (This)->lpVtbl -> Missing27(This) ) 

#define _Form_Missing28(This)	\
    ( (This)->lpVtbl -> Missing28(This) ) 

#define _Form_Missing29(This)	\
    ( (This)->lpVtbl -> Missing29(This) ) 

#define _Form_Missing30(This)	\
    ( (This)->lpVtbl -> Missing30(This) ) 

#define _Form_Missing31(This)	\
    ( (This)->lpVtbl -> Missing31(This) ) 

#define _Form_Missing32(This)	\
    ( (This)->lpVtbl -> Missing32(This) ) 

#define _Form_Missing33(This)	\
    ( (This)->lpVtbl -> Missing33(This) ) 

#define _Form_Missing34(This)	\
    ( (This)->lpVtbl -> Missing34(This) ) 

#define _Form_Missing35(This)	\
    ( (This)->lpVtbl -> Missing35(This) ) 

#define _Form_get_Count(This,rhs)	\
    ( (This)->lpVtbl -> get_Count(This,rhs) ) 

#define _Form_Missing36(This)	\
    ( (This)->lpVtbl -> Missing36(This) ) 

#define _Form_get_Controls(This,rhs)	\
    ( (This)->lpVtbl -> get_Controls(This,rhs) ) 

#define _Form_Missing37(This)	\
    ( (This)->lpVtbl -> Missing37(This) ) 

#define _Form_get_MouseIcon(This,rhs)	\
    ( (This)->lpVtbl -> get_MouseIcon(This,rhs) ) 

#define _Form_put_MouseIcon(This,rhs)	\
    ( (This)->lpVtbl -> put_MouseIcon(This,rhs) ) 

#define _Form_Missing38(This)	\
    ( (This)->lpVtbl -> Missing38(This) ) 

#define _Form_Missing39(This)	\
    ( (This)->lpVtbl -> Missing39(This) ) 

#define _Form_Missing40(This)	\
    ( (This)->lpVtbl -> Missing40(This) ) 

#define _Form_Missing41(This)	\
    ( (This)->lpVtbl -> Missing41(This) ) 

#define _Form_Missing42(This)	\
    ( (This)->lpVtbl -> Missing42(This) ) 

#define _Form_Missing43(This)	\
    ( (This)->lpVtbl -> Missing43(This) ) 

#define _Form_Missing44(This)	\
    ( (This)->lpVtbl -> Missing44(This) ) 

#define _Form_Missing45(This)	\
    ( (This)->lpVtbl -> Missing45(This) ) 

#define _Form_get_Font(This,rhs)	\
    ( (This)->lpVtbl -> get_Font(This,rhs) ) 

#define _Form_putref_Font(This,rhs)	\
    ( (This)->lpVtbl -> putref_Font(This,rhs) ) 

#define _Form_get_Appearance(This,rhs)	\
    ( (This)->lpVtbl -> get_Appearance(This,rhs) ) 

#define _Form_put_Appearance(This,rhs)	\
    ( (This)->lpVtbl -> put_Appearance(This,rhs) ) 

#define _Form_get_WhatsThisButton(This,rhs)	\
    ( (This)->lpVtbl -> get_WhatsThisButton(This,rhs) ) 

#define _Form_Missing46(This)	\
    ( (This)->lpVtbl -> Missing46(This) ) 

#define _Form_get_WhatsThisHelp(This,rhs)	\
    ( (This)->lpVtbl -> get_WhatsThisHelp(This,rhs) ) 

#define _Form_Missing47(This)	\
    ( (This)->lpVtbl -> Missing47(This) ) 

#define _Form_get_ShowInTaskbar(This,rhs)	\
    ( (This)->lpVtbl -> get_ShowInTaskbar(This,rhs) ) 

#define _Form_Missing48(This)	\
    ( (This)->lpVtbl -> Missing48(This) ) 

#define _Form_get_RightToLeft(This,rhs)	\
    ( (This)->lpVtbl -> get_RightToLeft(This,rhs) ) 

#define _Form_put_RightToLeft(This,rhs)	\
    ( (This)->lpVtbl -> put_RightToLeft(This,rhs) ) 

#define _Form_get_StartUpPosition(This,rhs)	\
    ( (This)->lpVtbl -> get_StartUpPosition(This,rhs) ) 

#define _Form_Missing49(This)	\
    ( (This)->lpVtbl -> Missing49(This) ) 

#define _Form_get_OLEDropMode(This,rhs)	\
    ( (This)->lpVtbl -> get_OLEDropMode(This,rhs) ) 

#define _Form_put_OLEDropMode(This,rhs)	\
    ( (This)->lpVtbl -> put_OLEDropMode(This,rhs) ) 

#define _Form_get_Palette(This,rhs)	\
    ( (This)->lpVtbl -> get_Palette(This,rhs) ) 

#define _Form_putref_Palette(This,rhs)	\
    ( (This)->lpVtbl -> putref_Palette(This,rhs) ) 

#define _Form_get_PaletteMode(This,rhs)	\
    ( (This)->lpVtbl -> get_PaletteMode(This,rhs) ) 

#define _Form_put_PaletteMode(This,rhs)	\
    ( (This)->lpVtbl -> put_PaletteMode(This,rhs) ) 

#define _Form_get_Moveable(This,rhs)	\
    ( (This)->lpVtbl -> get_Moveable(This,rhs) ) 

#define _Form_Missing50(This)	\
    ( (This)->lpVtbl -> Missing50(This) ) 

#define _Form_Refresh(This)	\
    ( (This)->lpVtbl -> Refresh(This) ) 

#define _Form_Move(This,Left,Top,Width,Height)	\
    ( (This)->lpVtbl -> Move(This,Left,Top,Width,Height) ) 

#define _Form_SetFocus(This)	\
    ( (This)->lpVtbl -> SetFocus(This) ) 

#define _Form_ZOrder(This,Position)	\
    ( (This)->lpVtbl -> ZOrder(This,Position) ) 

#define _Form_Show(This,modal,OwnerForm)	\
    ( (This)->lpVtbl -> Show(This,modal,OwnerForm) ) 

#define _Form_Hide(This)	\
    ( (This)->lpVtbl -> Hide(This) ) 

#define _Form_PrintForm(This)	\
    ( (This)->lpVtbl -> PrintForm(This) ) 

#define _Form_PopupMenu(This,Menu,Flags,x,y,BoldCommand)	\
    ( (This)->lpVtbl -> PopupMenu(This,Menu,Flags,x,y,BoldCommand) ) 

#define _Form_Circle(This,x,y,Radius,Color,StartAngle,EndAngle,Aspect)	\
    ( (This)->lpVtbl -> Circle(This,x,y,Radius,Color,StartAngle,EndAngle,Aspect) ) 

#define _Form_Cls(This)	\
    ( (This)->lpVtbl -> Cls(This) ) 

#define _Form_Line(This,x1,y1,x2,y2,Color,BorF)	\
    ( (This)->lpVtbl -> Line(This,x1,y1,x2,y2,Color,BorF) ) 

#define _Form_PaintPicture(This,Picture,x1,y1,width1,height1,x2,y2,width2,height2,Opcode)	\
    ( (This)->lpVtbl -> PaintPicture(This,Picture,x1,y1,width1,height1,x2,y2,width2,height2,Opcode) ) 

#define _Form_Point(This,x,y,rhs)	\
    ( (This)->lpVtbl -> Point(This,x,y,rhs) ) 

#define _Form_PSet(This,x,y,Color)	\
    ( (This)->lpVtbl -> PSet(This,x,y,Color) ) 

#define _Form_Scale(This,x1,y1,x2,y2)	\
    ( (This)->lpVtbl -> Scale(This,x1,y1,x2,y2) ) 

#define _Form_ScaleX(This,Value,fromScale,toScale,rhs)	\
    ( (This)->lpVtbl -> ScaleX(This,Value,fromScale,toScale,rhs) ) 

#define _Form_ScaleY(This,Value,fromScale,toScale,rhs)	\
    ( (This)->lpVtbl -> ScaleY(This,Value,fromScale,toScale,rhs) ) 

#define _Form_TextWidth(This,str,rhs)	\
    ( (This)->lpVtbl -> TextWidth(This,str,rhs) ) 

#define _Form_TextHeight(This,str,rhs)	\
    ( (This)->lpVtbl -> TextHeight(This,str,rhs) ) 

#define _Form_ValidateControls(This,rhs)	\
    ( (This)->lpVtbl -> ValidateControls(This,rhs) ) 

#define _Form_WhatsThisMode(This)	\
    ( (This)->lpVtbl -> WhatsThisMode(This) ) 

#define _Form_OLEDrag(This)	\
    ( (This)->lpVtbl -> OLEDrag(This) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [helpstring] */ HRESULT __stdcall _Form_Missing32_Proxy( 
    _Form * This);


void __RPC_STUB _Form_Missing32_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring] */ HRESULT __stdcall _Form_Missing33_Proxy( 
    _Form * This);


void __RPC_STUB _Form_Missing33_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring] */ HRESULT __stdcall _Form_Missing34_Proxy( 
    _Form * This);


void __RPC_STUB _Form_Missing34_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring] */ HRESULT __stdcall _Form_Missing35_Proxy( 
    _Form * This);


void __RPC_STUB _Form_Missing35_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][propget] */ HRESULT __stdcall _Form_get_Count_Proxy( 
    _Form * This,
    /* [retval][out] */ short *rhs);


void __RPC_STUB _Form_get_Count_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring] */ HRESULT __stdcall _Form_Missing36_Proxy( 
    _Form * This);


void __RPC_STUB _Form_Missing36_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][propget] */ HRESULT __stdcall _Form_get_Controls_Proxy( 
    _Form * This,
    /* [retval][out] */ long *rhs);


void __RPC_STUB _Form_get_Controls_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring] */ HRESULT __stdcall _Form_Missing37_Proxy( 
    _Form * This);


void __RPC_STUB _Form_Missing37_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][propget] */ HRESULT __stdcall _Form_get_MouseIcon_Proxy( 
    _Form * This,
    /* [retval][out] */ long *rhs);


void __RPC_STUB _Form_get_MouseIcon_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][propput] */ HRESULT __stdcall _Form_put_MouseIcon_Proxy( 
    _Form * This,
    /* [in] */ long rhs);


void __RPC_STUB _Form_put_MouseIcon_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring] */ HRESULT __stdcall _Form_Missing38_Proxy( 
    _Form * This);


void __RPC_STUB _Form_Missing38_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring] */ HRESULT __stdcall _Form_Missing39_Proxy( 
    _Form * This);


void __RPC_STUB _Form_Missing39_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring] */ HRESULT __stdcall _Form_Missing40_Proxy( 
    _Form * This);


void __RPC_STUB _Form_Missing40_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring] */ HRESULT __stdcall _Form_Missing41_Proxy( 
    _Form * This);


void __RPC_STUB _Form_Missing41_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring] */ HRESULT __stdcall _Form_Missing42_Proxy( 
    _Form * This);


void __RPC_STUB _Form_Missing42_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring] */ HRESULT __stdcall _Form_Missing43_Proxy( 
    _Form * This);


void __RPC_STUB _Form_Missing43_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring] */ HRESULT __stdcall _Form_Missing44_Proxy( 
    _Form * This);


void __RPC_STUB _Form_Missing44_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring] */ HRESULT __stdcall _Form_Missing45_Proxy( 
    _Form * This);


void __RPC_STUB _Form_Missing45_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][propget] */ HRESULT __stdcall _Form_get_Font_Proxy( 
    _Form * This,
    /* [retval][out] */ long *rhs);


void __RPC_STUB _Form_get_Font_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][propputref] */ HRESULT __stdcall _Form_putref_Font_Proxy( 
    _Form * This,
    /* [in] */ long rhs);


void __RPC_STUB _Form_putref_Font_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][propget] */ HRESULT __stdcall _Form_get_Appearance_Proxy( 
    _Form * This,
    /* [retval][out] */ short *rhs);


void __RPC_STUB _Form_get_Appearance_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][propput] */ HRESULT __stdcall _Form_put_Appearance_Proxy( 
    _Form * This,
    /* [in] */ short rhs);


void __RPC_STUB _Form_put_Appearance_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][propget] */ HRESULT __stdcall _Form_get_WhatsThisButton_Proxy( 
    _Form * This,
    /* [retval][out] */ VARIANT_BOOL *rhs);


void __RPC_STUB _Form_get_WhatsThisButton_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring] */ HRESULT __stdcall _Form_Missing46_Proxy( 
    _Form * This);


void __RPC_STUB _Form_Missing46_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][propget] */ HRESULT __stdcall _Form_get_WhatsThisHelp_Proxy( 
    _Form * This,
    /* [retval][out] */ VARIANT_BOOL *rhs);


void __RPC_STUB _Form_get_WhatsThisHelp_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring] */ HRESULT __stdcall _Form_Missing47_Proxy( 
    _Form * This);


void __RPC_STUB _Form_Missing47_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][propget] */ HRESULT __stdcall _Form_get_ShowInTaskbar_Proxy( 
    _Form * This,
    /* [retval][out] */ VARIANT_BOOL *rhs);


void __RPC_STUB _Form_get_ShowInTaskbar_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring] */ HRESULT __stdcall _Form_Missing48_Proxy( 
    _Form * This);


void __RPC_STUB _Form_Missing48_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][propget] */ HRESULT __stdcall _Form_get_RightToLeft_Proxy( 
    _Form * This,
    /* [retval][out] */ VARIANT_BOOL *rhs);


void __RPC_STUB _Form_get_RightToLeft_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][propput] */ HRESULT __stdcall _Form_put_RightToLeft_Proxy( 
    _Form * This,
    /* [in] */ VARIANT_BOOL rhs);


void __RPC_STUB _Form_put_RightToLeft_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][propget] */ HRESULT __stdcall _Form_get_StartUpPosition_Proxy( 
    _Form * This,
    /* [retval][out] */ short *rhs);


void __RPC_STUB _Form_get_StartUpPosition_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring] */ HRESULT __stdcall _Form_Missing49_Proxy( 
    _Form * This);


void __RPC_STUB _Form_Missing49_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][propget] */ HRESULT __stdcall _Form_get_OLEDropMode_Proxy( 
    _Form * This,
    /* [retval][out] */ short *rhs);


void __RPC_STUB _Form_get_OLEDropMode_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][propput] */ HRESULT __stdcall _Form_put_OLEDropMode_Proxy( 
    _Form * This,
    /* [in] */ short rhs);


void __RPC_STUB _Form_put_OLEDropMode_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][propget] */ HRESULT __stdcall _Form_get_Palette_Proxy( 
    _Form * This,
    /* [retval][out] */ long *rhs);


void __RPC_STUB _Form_get_Palette_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][propputref] */ HRESULT __stdcall _Form_putref_Palette_Proxy( 
    _Form * This,
    /* [in] */ long rhs);


void __RPC_STUB _Form_putref_Palette_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][propget] */ HRESULT __stdcall _Form_get_PaletteMode_Proxy( 
    _Form * This,
    /* [retval][out] */ short *rhs);


void __RPC_STUB _Form_get_PaletteMode_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][propput] */ HRESULT __stdcall _Form_put_PaletteMode_Proxy( 
    _Form * This,
    /* [in] */ short rhs);


void __RPC_STUB _Form_put_PaletteMode_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][propget] */ HRESULT __stdcall _Form_get_Moveable_Proxy( 
    _Form * This,
    /* [retval][out] */ VARIANT_BOOL *rhs);


void __RPC_STUB _Form_get_Moveable_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring] */ HRESULT __stdcall _Form_Missing50_Proxy( 
    _Form * This);


void __RPC_STUB _Form_Missing50_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring] */ HRESULT __stdcall _Form_Refresh_Proxy( 
    _Form * This);


void __RPC_STUB _Form_Refresh_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring] */ HRESULT __stdcall _Form_Move_Proxy( 
    _Form * This,
    /* [in] */ float Left,
    /* [optional][in] */ VARIANT Top,
    /* [optional][in] */ VARIANT Width,
    /* [optional][in] */ VARIANT Height);


void __RPC_STUB _Form_Move_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring] */ HRESULT __stdcall _Form_SetFocus_Proxy( 
    _Form * This);


void __RPC_STUB _Form_SetFocus_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring] */ HRESULT __stdcall _Form_ZOrder_Proxy( 
    _Form * This,
    /* [optional][in] */ VARIANT Position);


void __RPC_STUB _Form_ZOrder_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring] */ HRESULT __stdcall _Form_Show_Proxy( 
    _Form * This,
    /* [in] */ VARIANT modal,
    /* [in] */ VARIANT OwnerForm);


void __RPC_STUB _Form_Show_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring] */ HRESULT __stdcall _Form_Hide_Proxy( 
    _Form * This);


void __RPC_STUB _Form_Hide_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring] */ HRESULT __stdcall _Form_PrintForm_Proxy( 
    _Form * This);


void __RPC_STUB _Form_PrintForm_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring] */ HRESULT __stdcall _Form_PopupMenu_Proxy( 
    _Form * This,
    /* [in] */ IDispatch *Menu,
    /* [optional][in] */ VARIANT Flags,
    /* [optional][in] */ VARIANT x,
    /* [optional][in] */ VARIANT y,
    /* [optional][in] */ VARIANT BoldCommand);


void __RPC_STUB _Form_PopupMenu_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring] */ HRESULT __stdcall _Form_Circle_Proxy( 
    _Form * This,
    /* [in] */ float x,
    /* [in] */ float y,
    /* [in] */ float Radius,
    /* [optional][in] */ VARIANT Color,
    /* [optional][in] */ VARIANT StartAngle,
    /* [optional][in] */ VARIANT EndAngle,
    /* [optional][in] */ VARIANT Aspect);


void __RPC_STUB _Form_Circle_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring] */ HRESULT __stdcall _Form_Cls_Proxy( 
    _Form * This);


void __RPC_STUB _Form_Cls_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring] */ HRESULT __stdcall _Form_Line_Proxy( 
    _Form * This,
    /* [in] */ float x1,
    /* [in] */ float y1,
    /* [in] */ float x2,
    /* [in] */ float y2,
    /* [optional][in] */ VARIANT Color,
    /* [optional][in] */ VARIANT BorF);


void __RPC_STUB _Form_Line_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring] */ HRESULT __stdcall _Form_PaintPicture_Proxy( 
    _Form * This,
    /* [in] */ IDispatch *Picture,
    /* [in] */ float x1,
    /* [in] */ float y1,
    /* [optional][in] */ VARIANT width1,
    /* [optional][in] */ VARIANT height1,
    /* [optional][in] */ VARIANT x2,
    /* [optional][in] */ VARIANT y2,
    /* [optional][in] */ VARIANT width2,
    /* [optional][in] */ VARIANT height2,
    /* [optional][in] */ VARIANT Opcode);


void __RPC_STUB _Form_PaintPicture_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring] */ HRESULT __stdcall _Form_Point_Proxy( 
    _Form * This,
    /* [in] */ float x,
    /* [in] */ float y,
    /* [retval][out] */ long *rhs);


void __RPC_STUB _Form_Point_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring] */ HRESULT __stdcall _Form_PSet_Proxy( 
    _Form * This,
    /* [in] */ float x,
    /* [in] */ float y,
    /* [optional][in] */ VARIANT Color);


void __RPC_STUB _Form_PSet_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring] */ HRESULT __stdcall _Form_Scale_Proxy( 
    _Form * This,
    /* [optional][in] */ VARIANT x1,
    /* [optional][in] */ VARIANT y1,
    /* [optional][in] */ VARIANT x2,
    /* [optional][in] */ VARIANT y2);


void __RPC_STUB _Form_Scale_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring] */ HRESULT __stdcall _Form_ScaleX_Proxy( 
    _Form * This,
    /* [in] */ float Value,
    /* [optional][in] */ VARIANT fromScale,
    /* [optional][in] */ VARIANT toScale,
    /* [retval][out] */ float *rhs);


void __RPC_STUB _Form_ScaleX_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring] */ HRESULT __stdcall _Form_ScaleY_Proxy( 
    _Form * This,
    /* [in] */ float Value,
    /* [optional][in] */ VARIANT fromScale,
    /* [optional][in] */ VARIANT toScale,
    /* [retval][out] */ float *rhs);


void __RPC_STUB _Form_ScaleY_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring] */ HRESULT __stdcall _Form_TextWidth_Proxy( 
    _Form * This,
    /* [in] */ BSTR str,
    /* [retval][out] */ float *rhs);


void __RPC_STUB _Form_TextWidth_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring] */ HRESULT __stdcall _Form_TextHeight_Proxy( 
    _Form * This,
    /* [in] */ BSTR str,
    /* [retval][out] */ float *rhs);


void __RPC_STUB _Form_TextHeight_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring] */ HRESULT __stdcall _Form_ValidateControls_Proxy( 
    _Form * This,
    /* [retval][out] */ VARIANT_BOOL *rhs);


void __RPC_STUB _Form_ValidateControls_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring] */ HRESULT __stdcall _Form_WhatsThisMode_Proxy( 
    _Form * This);


void __RPC_STUB _Form_WhatsThisMode_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring] */ HRESULT __stdcall _Form_OLEDrag_Proxy( 
    _Form * This);


void __RPC_STUB _Form_OLEDrag_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* ___Form_INTERFACE_DEFINED__ */


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


