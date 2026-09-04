#include "vba_internal.h"
#include "FormWrapper.hpp"
#include "EventDispatch.hpp"
#include "FormWindow.hpp"
#include "FormProperties.hpp"

// MIDL-generated header for the real _Form interface (Form.idl) -- included plainly,
// the same way App.cpp includes App.h, so _Form comes in as a genuine C++ abstract
// class (pure virtual methods) rather than the raw "struct of function pointers"
// CINTERFACE produces. See _FormImpl below for why that's enough even though this
// project still ultimately needs a raw vtable blob at runtime.
#include "Form.h"

vbFormWrapper::vbFormWrapper(
	vba_VBVTable			*pWrapperVtable,
	ObjectInfoWithOptional	*pObjInfo
)
	: vbClassWrapper(pWrapperVtable, pObjInfo)
{
}

vbFormWrapper::~vbFormWrapper()
{
	// A form whose owning object is going away without ever being explicitly Unloaded
	// would otherwise leave a real, still-visible window pointing at freed memory via
	// GWLP_USERDATA -- tear it down here instead.
	DestroyFormWindow();
}

bool vbFormWrapper::EnsureWindowCreated()
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	if (m_hwnd)
	{
		return true;
	}

	FormTemplateProps templateProps;
	bool bTemplateFound = ParseFormTemplate(m_pObjInfo, &templateProps);

	wchar_t wszCaption[256];
	MultiByteToWideChar(CP_ACP, 0, templateProps.caption, -1, wszCaption, 256);

	DEBUG_WIDE_OBJ(
		"bTemplateFound %d, caption \"%hs\", left %d, top %d, width %d, height %d, backColor %.8x, borderStyle %d, enabled %d, controlBox %d, maxButton %d, minButton %d, windowState %d",
		(int)bTemplateFound,
		templateProps.caption,
		templateProps.clientLeft,
		templateProps.clientTop,
		templateProps.clientWidth,
		templateProps.clientHeight,
		templateProps.backColor,
		(int)templateProps.borderStyle,
		(int)templateProps.enabled,
		(int)templateProps.controlBox,
		(int)templateProps.maxButton,
		(int)templateProps.minButton,
		(int)templateProps.windowState
	);

	m_windowState = templateProps.windowState;

	FormWindowCreateParams createParams = {};
	createParams.caption = wszCaption;
	createParams.pOwnerVBVTable = m_pVBVTable;
	createParams.clientLeftTwips = templateProps.clientLeft;
	createParams.clientTopTwips = templateProps.clientTop;
	createParams.clientWidthTwips = templateProps.clientWidth;
	createParams.clientHeightTwips = templateProps.clientHeight;
	createParams.backColor = templateProps.backColor;
	createParams.borderStyle = templateProps.borderStyle;
	createParams.enabled = templateProps.enabled;
	createParams.controlBox = templateProps.controlBox;
	createParams.maxButton = templateProps.maxButton;
	createParams.minButton = templateProps.minButton;

	m_hwnd = CreateFormWindow(&m_hwnd, createParams);

	DEBUG_WIDE_OBJ(
		"m_hwnd %.8x",
		(unsigned int)m_hwnd
	);

	if (m_hwnd)
	{
		// Real VB6 fires Form_Load once, right after the form is created (via the
		// Load statement or an implicit first Show) and before it's actually shown.
		TryFireLoadOnce();
	}

	return m_hwnd != nullptr;
}

void vbFormWrapper::DestroyFormWindow()
{
	if (m_hwnd)
	{
		// FormWndProc clears m_hwnd itself (via the slot pointer handed to
		// CreateFormWindow) once WM_DESTROY actually runs.
		DestroyWindow(m_hwnd);
	}
}

HRESULT vbFormWrapper::Show(VARIANTARG windowStyle, VARIANTARG /*ownerForm*/)
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	if (!EnsureWindowCreated())
	{
		return E_FAIL;
	}

	int nCmdShow = SW_SHOWNORMAL;
	if (m_windowState == 1) nCmdShow = SW_SHOWMINIMIZED;
	else if (m_windowState == 2) nCmdShow = SW_SHOWMAXIMIZED;

	ShowWindow(m_hwnd, nCmdShow);
	UpdateWindow(m_hwnd);

	// vbModal (1) is Show's documented default when the WindowMode argument is
	// omitted (confirmed live: FormTest's compiled call passes a VT_ERROR/
	// DISP_E_PARAMNOTFOUND-style Variant for an omitted OwnerForm, but always
	// supplies a real VT_I4/VT_I2 value for WindowMode itself in this codepath).
	bool bModal = true;
	if (windowStyle.vt == VT_I4)
	{
		bModal = (windowStyle.lVal != 0);
	}
	else if (windowStyle.vt == VT_I2)
	{
		bModal = (windowStyle.iVal != 0);
	}

	DEBUG_WIDE_OBJ(
		"m_hwnd %.8x, bModal %d",
		(unsigned int)m_hwnd,
		(int)bModal
	);

	if (bModal)
	{
		RunModalMessageLoop(&m_hwnd);
	}

	return S_OK;
}

/**
 * @brief			Finds the sole implemented intrinsic-event thunk whose target sub's
 *					stdcall argument size (see GetStdcallArgBytes) matches
 *					wantedArgBytes, or nullptr if none/more than one match.
 *
 *					Background: a Form-derived class's intrinsic event handlers reuse
 *					the same "sub esp+4,imm; jmp Handler" thunk shape WithEvents
 *					already uses, packed tightly starting at a fixed offset in its
 *					self-sink block, one per implemented handler, in SOURCE
 *					DECLARATION ORDER (confirmed live by compiling calibration test
 *					forms) -- there is no compiled name/ordinal table saying which
 *					slot is which event (checked: opt fields, PublicObjectDescriptor.
 *					lpMethodNames, the control-record header bytes; the real msvbvm60
 *					mechanism traces into a per-object-type C++ virtual dispatch
 *					inside its window procedure that wasn't fully resolved). Different
 *					intrinsic events mostly take different argument shapes though
 *					(Load/Click/Activate/... take none; QueryUnload takes two ByRef
 *					Integers = 8 bytes; Unload takes one = 4 bytes; etc.), so
 *					classifying each implemented thunk by its target's actual stdcall
 *					argument byte count disambiguates them whenever the currently-
 *					implemented set doesn't have two events sharing the same byte
 *					count (e.g. Load+Click together would collide at 0 bytes) --
 *					in which case this deliberately returns nullptr rather than guess.
 */
void * vbFormWrapper::FindIntrinsicThunkByArgBytes(int wantedArgBytes) const
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	if (!m_pObjInfo || m_pObjInfo->opt.wEventCount == 0 || !m_pObjInfo->opt.lpEvents)
	{
		return nullptr;
	}

	void *pMatch = nullptr;
	int matchCount = 0;

	__try
	{
		void **pRawThunkArray = (void**)m_pObjInfo->opt.lpEvents;
		for (WORD i = 0; i < m_pObjInfo->opt.wEventCount; i++)
		{
			void *pRawThunkMid = pRawThunkArray[i];
			if (!pRawThunkMid)
			{
				continue;
			}

			// opt.lpEvents[i] points 8 bytes into the thunk (past its
			// "sub esp+4,imm32" instruction, landing on the "jmp Handler") for this
			// self-event shape -- confirmed live across every calibration form.
			void *pThunkStart = (void*)((BYTE*)pRawThunkMid - 8);
			void *pTarget = ResolveHandlerThunkTarget(pThunkStart);
			int rawArgBytes = GetStdcallArgBytes(pTarget);

			// Every handler's retn count includes an implicit 4-byte "this" pointer
			// on top of its VB-visible arguments -- confirmed live: Form_Load (no
			// VB arguments) actually ends "retn 4", and Form_QueryUnload (two ByRef
			// Integers) ends "retn 0Ch" (12), not the raw 0/8 this function compares
			// against.
			int visibleArgBytes = (rawArgBytes >= 4) ? (rawArgBytes - 4) : -1;

			DEBUG_WIDE_OBJ(
				"thunk[%d] %.8x -> target %.8x, rawArgBytes %d, visibleArgBytes %d, wanted %d",
				i,
				(unsigned int)pThunkStart,
				(unsigned int)pTarget,
				rawArgBytes,
				visibleArgBytes,
				wantedArgBytes
			);

			if (visibleArgBytes == wantedArgBytes)
			{
				pMatch = pThunkStart;
				matchCount++;
			}
		}
	}
	__except (EXCEPTION_EXECUTE_HANDLER)
	{
		return nullptr;
	}

	return (matchCount == 1) ? pMatch : nullptr;
}

bool vbFormWrapper::TryFireQueryUnload(short *pCancel)
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	if (pCancel)
	{
		*pCancel = 0;
	}

	// QueryUnload(Cancel As Integer, UnloadMode As Integer), both ByRef -> 8 bytes.
	void *pThunkStart = FindIntrinsicThunkByArgBytes(8);
	if (!pThunkStart)
	{
		return false;
	}

	short unloadMode = 1; // vbFormControlMenu -- closed via the window's own system menu/X button

	__try
	{
		VARIANTARG args[2] = {};
		args[0].vt = VT_I4;
		args[0].lVal = (LONG)(INT_PTR)pCancel;
		args[1].vt = VT_I4;
		args[1].lVal = (LONG)(INT_PTR)&unloadMode;

		DEBUG_WIDE_OBJ(
			"pThunkStart %.8x",
			(unsigned int)pThunkStart
		);

		InvokeHandlerThunk(pThunkStart, m_pVBVTable, args, 2);
	}
	__except (EXCEPTION_EXECUTE_HANDLER)
	{
		DEBUG_WIDE_OBJ(
			"exception invoking QueryUnload -- classification guess was wrong for this thunk"
		);
		return false;
	}

	return true;
}

void vbFormWrapper::TryFireLoadOnce()
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	if (m_bLoadFired)
	{
		return;
	}
	m_bLoadFired = true;

	// Form_Load() takes no arguments -> 0 bytes.
	void *pThunkStart = FindIntrinsicThunkByArgBytes(0);
	if (!pThunkStart)
	{
		return;
	}

	__try
	{
		DEBUG_WIDE_OBJ(
			"pThunkStart %.8x",
			(unsigned int)pThunkStart
		);

		InvokeHandlerThunk(pThunkStart, m_pVBVTable, nullptr, 0);
	}
	__except (EXCEPTION_EXECUTE_HANDLER)
	{
		DEBUG_WIDE_OBJ(
			"exception invoking Load -- classification guess was wrong for this thunk"
		);
	}
}

/**
 * Real, compiler-implemented `_Form` interface (Form.idl) -- the same pattern
 * App.cpp's `_AppImpl : public _App` and VBGlobal.cpp's VBGlobal implementation
 * already use for their own MIDL interfaces: a genuine C++ class with a genuine,
 * compiler-generated vtable, not a hand-maintained raw struct/CINTERFACE cast.
 *
 * Most members are inert stubs (matching how App.cpp's own MissingNN placeholders
 * are handled) -- this project doesn't model most of `_Form`'s surface yet (see
 * Form.idl's own header comment). What actually matters is structural: this class
 * must implement literally every member `_Form` declares to be instantiable at all,
 * which means its compiler-generated vtable is guaranteed to have every real member
 * (Show, Caption, ...) at exactly its real, compiler-assigned slot -- if a future
 * edit to Form.idl ever changed that, this class would simply stop compiling (a pure
 * virtual left unimplemented, or `override` failing to match a removed one) rather
 * than silently drifting, which is the whole point of doing it this way instead of a
 * hand-maintained offset constant.
 *
 * A single instance of this class only ever exists as a vtable *template*: see
 * GetFormInterfaceVtableTail below, which is the only thing that ever touches it.
 * Nothing calls its IUnknown/IDispatch members or QueryInterfaces for `_Form` on it
 * for real -- this project's own vba_BASIC_CLASS_IUnknownBridge covers those slots in
 * every wrapped object instead, Form or not (see ClassWrapper.hpp) -- so those stay
 * harmless stubs too. Every REAL member (Show, get_Caption/put_Caption, ...) is the
 * exception: __vbaNew (ObjectManipulation.cpp) copies this class's whole vtable (from
 * Missing1 on) into a Form-derived object's real, synthesized per-instance vtable, so
 * those overrides' addresses end up embedded there and genuinely execute when
 * compiled VB6 code calls through them. Each such override resolves its owning
 * vbFormWrapper via OwnerWrapper() below and does the real work using whatever
 * low-level primitive that class exposes (e.g. GetHwnd()) -- this class is the one
 * and only place `_Form`'s actual COM property/method behavior is implemented;
 * vbFormWrapper itself stays free of any COM-interface-shaped surface, owning only
 * the window's lifecycle (creation/destruction/intrinsic event dispatch).
 */
class _FormImpl : public _Form
{
public:
	// IUnknown/IDispatch: never actually reached (see the class comment above) but
	// required for this class to be concrete.
	HRESULT __stdcall QueryInterface(REFIID riid, void **ppvObject) override
	{
		return E_NOTIMPL;
	}

	ULONG __stdcall AddRef() override
	{
		return 1;
	}

	ULONG __stdcall Release() override
	{
		return 1;
	}

	HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) override
	{
		return E_NOTIMPL;
	}

	HRESULT __stdcall GetTypeInfo(UINT itinfo, LCID lcid, ITypeInfo **pptinfo) override
	{
		return E_NOTIMPL;
	}

	HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgdispid) override
	{
		return E_NOTIMPL;
	}

	HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) override
	{
		return E_NOTIMPL;
	}

private:
	/**
	 * Resolves "this" to the owning vbFormWrapper -- see the class comment above:
	 * every real (non-stub) member below is only ever reached via a copied-vtable
	 * slot call whose real "this" is a vba_VBVTable*, not a _FormImpl*, exactly like
	 * BASIC_CLASS_Form_Show used to take that explicitly as its first parameter.
	 * Returns nullptr if that invariant somehow doesn't hold (never expected).
	 */
	vbFormWrapper * OwnerWrapper()
	{
		vba_VBVTable *pFakeVtable = reinterpret_cast<vba_VBVTable*>(this);
		if (!pFakeVtable || !pFakeVtable->pWrapper)
		{
			return nullptr;
		}

		return dynamic_cast<vbFormWrapper*>(pFakeVtable->pWrapper);
	}

public:
	/* [helpstring] */ HRESULT __stdcall Missing1( void) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring] */ HRESULT __stdcall Missing2( void) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring] */ HRESULT __stdcall Missing3( void) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring] */ HRESULT __stdcall Missing4( void) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring] */ HRESULT __stdcall Missing5( void) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring] */ HRESULT __stdcall Missing6( void) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring] */ HRESULT __stdcall Missing7( void) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring] */ HRESULT __stdcall Missing8( void) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring] */ HRESULT __stdcall Missing9( void) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring] */ HRESULT __stdcall Missing10( void) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring] */ HRESULT __stdcall Missing11( void) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propget] */ HRESULT __stdcall get_Name( BSTR *rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring] */ HRESULT __stdcall Missing12( void) override
	{
		return E_NOTIMPL;
	}

	/* Real implementation -- the window's own title bar text. Referencing a Form
	   property implicitly loads it in real VB6, hence the EnsureWindowCreated() call
	   before touching the (otherwise possibly-not-yet-created) window handle. */
	HRESULT __stdcall get_Caption(BSTR *rhs) override
	{
		if (!rhs)
		{
			return E_POINTER;
		}

		vbFormWrapper *pFormWrapper = OwnerWrapper();
		if (!pFormWrapper)
		{
			return E_UNEXPECTED;
		}

		pFormWrapper->EnsureWindowCreated();

		wchar_t caption[512] = { 0 };
		HWND hwnd = pFormWrapper->GetHwnd();
		if (hwnd)
		{
			GetWindowTextW(hwnd, caption, ARRAYSIZE(caption));
		}

		*rhs = SysAllocString(caption);
		return *rhs ? S_OK : E_OUTOFMEMORY;
	}

	HRESULT __stdcall put_Caption(BSTR rhs) override
	{
		vbFormWrapper *pFormWrapper = OwnerWrapper();
		if (!pFormWrapper)
		{
			return E_UNEXPECTED;
		}

		pFormWrapper->EnsureWindowCreated();

		HWND hwnd = pFormWrapper->GetHwnd();
		if (hwnd)
		{
			SetWindowTextW(hwnd, rhs ? rhs : L"");
		}

		return S_OK;
	}

	/* [helpstring][propget] */ HRESULT __stdcall get_hWnd( long *rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring] */ HRESULT __stdcall Missing13( void) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propget] */ HRESULT __stdcall get_BackColor( long *rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propput] */ HRESULT __stdcall put_BackColor( long rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propget] */ HRESULT __stdcall get_ForeColor( long *rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propput] */ HRESULT __stdcall put_ForeColor( long rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propget] */ HRESULT __stdcall get_Left( long *rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propput] */ HRESULT __stdcall put_Left( long rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propget] */ HRESULT __stdcall get_Top( long *rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propput] */ HRESULT __stdcall put_Top( long rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propget] */ HRESULT __stdcall get_Width( long *rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propput] */ HRESULT __stdcall put_Width( long rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propget] */ HRESULT __stdcall get_Height( long *rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propput] */ HRESULT __stdcall put_Height( long rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propget] */ HRESULT __stdcall get_Enabled( VARIANT_BOOL *rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propput] */ HRESULT __stdcall put_Enabled( VARIANT_BOOL rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propget] */ HRESULT __stdcall get_WindowState( long *rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propput] */ HRESULT __stdcall put_WindowState( long rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propget] */ HRESULT __stdcall get_MousePointer( long *rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propput] */ HRESULT __stdcall put_MousePointer( long rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propget] */ HRESULT __stdcall get_FontName( BSTR *rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propput] */ HRESULT __stdcall put_FontName( BSTR rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propget] */ HRESULT __stdcall get_FontSize( float *rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propput] */ HRESULT __stdcall put_FontSize( float rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propget] */ HRESULT __stdcall get_FontBold( VARIANT_BOOL *rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propput] */ HRESULT __stdcall put_FontBold( VARIANT_BOOL rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propget] */ HRESULT __stdcall get_FontItalic( VARIANT_BOOL *rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propput] */ HRESULT __stdcall put_FontItalic( VARIANT_BOOL rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propget] */ HRESULT __stdcall get_FontStrikethru( VARIANT_BOOL *rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propput] */ HRESULT __stdcall put_FontStrikethru( VARIANT_BOOL rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propget] */ HRESULT __stdcall get_FontUnderline( VARIANT_BOOL *rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propput] */ HRESULT __stdcall put_FontUnderline( VARIANT_BOOL rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propget] */ HRESULT __stdcall get_hDC( long *rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring] */ HRESULT __stdcall Missing14( void) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propget] */ HRESULT __stdcall get_CurrentX( float *rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propput] */ HRESULT __stdcall put_CurrentX( float rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propget] */ HRESULT __stdcall get_CurrentY( float *rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propput] */ HRESULT __stdcall put_CurrentY( float rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propget] */ HRESULT __stdcall get_ScaleLeft( float *rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propput] */ HRESULT __stdcall put_ScaleLeft( float rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propget] */ HRESULT __stdcall get_ScaleTop( float *rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propput] */ HRESULT __stdcall put_ScaleTop( float rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propget] */ HRESULT __stdcall get_ScaleWidth( float *rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propput] */ HRESULT __stdcall put_ScaleWidth( float rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propget] */ HRESULT __stdcall get_ScaleHeight( float *rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propput] */ HRESULT __stdcall put_ScaleHeight( float rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propget] */ HRESULT __stdcall get_ScaleMode( short *rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propput] */ HRESULT __stdcall put_ScaleMode( short rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propget] */ HRESULT __stdcall get_FontTransparent( VARIANT_BOOL *rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propput] */ HRESULT __stdcall put_FontTransparent( VARIANT_BOOL rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propget] */ HRESULT __stdcall get_DrawStyle( short *rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propput] */ HRESULT __stdcall put_DrawStyle( short rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propget] */ HRESULT __stdcall get_DrawWidth( short *rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propput] */ HRESULT __stdcall put_DrawWidth( short rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propget] */ HRESULT __stdcall get_FillStyle( short *rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propput] */ HRESULT __stdcall put_FillStyle( short rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propget] */ HRESULT __stdcall get_FillColor( long *rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propput] */ HRESULT __stdcall put_FillColor( long rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propget] */ HRESULT __stdcall get_DrawMode( short *rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propput] */ HRESULT __stdcall put_DrawMode( short rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propget] */ HRESULT __stdcall get_AutoRedraw( VARIANT_BOOL *rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propput] */ HRESULT __stdcall put_AutoRedraw( VARIANT_BOOL rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propget] */ HRESULT __stdcall get_Picture( long *rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propput] */ HRESULT __stdcall put_Picture( long rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propget] */ HRESULT __stdcall get_BorderStyle( short *rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propput] */ HRESULT __stdcall put_BorderStyle( short rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propget] */ HRESULT __stdcall get_Icon( long *rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propput] */ HRESULT __stdcall put_Icon( long rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propget] */ HRESULT __stdcall get_LinkTopic( BSTR *rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propput] */ HRESULT __stdcall put_LinkTopic( BSTR rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propget] */ HRESULT __stdcall get_LinkMode( short *rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propput] */ HRESULT __stdcall put_LinkMode( short rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propget] */ HRESULT __stdcall get_MaxButton( VARIANT_BOOL *rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring] */ HRESULT __stdcall Missing15( void) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propget] */ HRESULT __stdcall get_MinButton( VARIANT_BOOL *rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring] */ HRESULT __stdcall Missing16( void) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propget] */ HRESULT __stdcall get_ControlBox( VARIANT_BOOL *rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring] */ HRESULT __stdcall Missing17( void) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propget] */ HRESULT __stdcall get_Image( long *rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring] */ HRESULT __stdcall Missing18( void) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propget] */ HRESULT __stdcall get_HasDC( VARIANT_BOOL *rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring] */ HRESULT __stdcall Missing19( void) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring] */ HRESULT __stdcall Missing20( void) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring] */ HRESULT __stdcall Missing21( void) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring] */ HRESULT __stdcall Missing22( void) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring] */ HRESULT __stdcall Missing23( void) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring] */ HRESULT __stdcall Missing24( void) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring] */ HRESULT __stdcall Missing25( void) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propget] */ HRESULT __stdcall get_Visible( VARIANT_BOOL *rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propput] */ HRESULT __stdcall put_Visible( VARIANT_BOOL rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propget] */ HRESULT __stdcall get_Tag( BSTR *rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propput] */ HRESULT __stdcall put_Tag( BSTR rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propget] */ HRESULT __stdcall get_MDIChild( VARIANT_BOOL *rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring] */ HRESULT __stdcall Missing26( void) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propget] */ HRESULT __stdcall get_KeyPreview( VARIANT_BOOL *rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propput] */ HRESULT __stdcall put_KeyPreview( VARIANT_BOOL rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propget] */ HRESULT __stdcall get_ClipControls( VARIANT_BOOL *rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propput] */ HRESULT __stdcall put_ClipControls( VARIANT_BOOL rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propget] */ HRESULT __stdcall get_HelpContextID( long *rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propput] */ HRESULT __stdcall put_HelpContextID( long rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propget] */ HRESULT __stdcall get_ActiveControl( long *rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring] */ HRESULT __stdcall Missing27( void) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring] */ HRESULT __stdcall Missing28( void) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring] */ HRESULT __stdcall Missing29( void) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring] */ HRESULT __stdcall Missing30( void) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring] */ HRESULT __stdcall Missing31( void) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring] */ HRESULT __stdcall Missing32( void) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring] */ HRESULT __stdcall Missing33( void) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring] */ HRESULT __stdcall Missing34( void) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring] */ HRESULT __stdcall Missing35( void) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propget] */ HRESULT __stdcall get_Count( short *rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring] */ HRESULT __stdcall Missing36( void) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propget] */ HRESULT __stdcall get_Controls( long *rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring] */ HRESULT __stdcall Missing37( void) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propget] */ HRESULT __stdcall get_MouseIcon( long *rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propput] */ HRESULT __stdcall put_MouseIcon( long rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring] */ HRESULT __stdcall Missing38( void) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring] */ HRESULT __stdcall Missing39( void) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring] */ HRESULT __stdcall Missing40( void) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring] */ HRESULT __stdcall Missing41( void) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring] */ HRESULT __stdcall Missing42( void) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring] */ HRESULT __stdcall Missing43( void) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring] */ HRESULT __stdcall Missing44( void) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring] */ HRESULT __stdcall Missing45( void) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propget] */ HRESULT __stdcall get_Font( long *rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propputref] */ HRESULT __stdcall putref_Font( long rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propget] */ HRESULT __stdcall get_Appearance( short *rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propput] */ HRESULT __stdcall put_Appearance( short rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propget] */ HRESULT __stdcall get_WhatsThisButton( VARIANT_BOOL *rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring] */ HRESULT __stdcall Missing46( void) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propget] */ HRESULT __stdcall get_WhatsThisHelp( VARIANT_BOOL *rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring] */ HRESULT __stdcall Missing47( void) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propget] */ HRESULT __stdcall get_ShowInTaskbar( VARIANT_BOOL *rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring] */ HRESULT __stdcall Missing48( void) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propget] */ HRESULT __stdcall get_RightToLeft( VARIANT_BOOL *rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propput] */ HRESULT __stdcall put_RightToLeft( VARIANT_BOOL rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propget] */ HRESULT __stdcall get_StartUpPosition( short *rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring] */ HRESULT __stdcall Missing49( void) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propget] */ HRESULT __stdcall get_OLEDropMode( short *rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propput] */ HRESULT __stdcall put_OLEDropMode( short rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propget] */ HRESULT __stdcall get_Palette( long *rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propputref] */ HRESULT __stdcall putref_Palette( long rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propget] */ HRESULT __stdcall get_PaletteMode( short *rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propput] */ HRESULT __stdcall put_PaletteMode( short rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring][propget] */ HRESULT __stdcall get_Moveable( VARIANT_BOOL *rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring] */ HRESULT __stdcall Missing50( void) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring] */ HRESULT __stdcall Refresh( void) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring] */ HRESULT __stdcall Move( float Left, VARIANT Top, VARIANT Width, VARIANT Height) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring] */ HRESULT __stdcall SetFocus( void) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring] */ HRESULT __stdcall ZOrder( VARIANT Position) override
	{
		return E_NOTIMPL;
	}

	/* Real implementation. Unlike most stubs above, this one genuinely runs: __vbaNew
	   (ObjectManipulation.cpp) copies this whole class's compiler-generated vtable
	   (from Missing1 on) into a Form-derived object's synthesized per-instance
	   vtable, so this exact function pointer ends up sitting at the real, compiler-
	   verified Show slot -- compiled VB6 code's "f.Show vbModal" calls straight
	   through to here. See OwnerWrapper()'s comment for why "this" isn't really a
	   _FormImpl* at that point. */
	HRESULT __stdcall Show(VARIANT modal, VARIANT OwnerForm) override
	{
		vbFormWrapper *pFormWrapper = OwnerWrapper();
		if (!pFormWrapper)
		{
			return E_UNEXPECTED;
		}

		return pFormWrapper->Show(modal, OwnerForm);
	}

	/* [helpstring] */ HRESULT __stdcall Hide( void) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring] */ HRESULT __stdcall PrintForm( void) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring] */ HRESULT __stdcall PopupMenu( IDispatch *Menu, VARIANT Flags, VARIANT x, VARIANT y, VARIANT BoldCommand) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring] */ HRESULT __stdcall Circle( float x, float y, float Radius, VARIANT Color, VARIANT StartAngle, VARIANT EndAngle, VARIANT Aspect) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring] */ HRESULT __stdcall Cls( void) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring] */ HRESULT __stdcall Line( float x1, float y1, float x2, float y2, VARIANT Color, VARIANT BorF) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring] */ HRESULT __stdcall PaintPicture( IDispatch *Picture, float x1, float y1, VARIANT width1, VARIANT height1, VARIANT x2, VARIANT y2, VARIANT width2, VARIANT height2, VARIANT Opcode) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring] */ HRESULT __stdcall Point( float x, float y, long *rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring] */ HRESULT __stdcall PSet( float x, float y, VARIANT Color) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring] */ HRESULT __stdcall Scale( VARIANT x1, VARIANT y1, VARIANT x2, VARIANT y2) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring] */ HRESULT __stdcall ScaleX( float Value, VARIANT fromScale, VARIANT toScale, float *rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring] */ HRESULT __stdcall ScaleY( float Value, VARIANT fromScale, VARIANT toScale, float *rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring] */ HRESULT __stdcall TextWidth( BSTR str, float *rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring] */ HRESULT __stdcall TextHeight( BSTR str, float *rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring] */ HRESULT __stdcall ValidateControls( VARIANT_BOOL *rhs) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring] */ HRESULT __stdcall WhatsThisMode( void) override
	{
		return E_NOTIMPL;
	}

	/* [helpstring] */ HRESULT __stdcall OLEDrag( void) override
	{
		return E_NOTIMPL;
	}

};

void * const * GetFormInterfaceVtableTail(size_t *pSlotCount)
{
	// A function-local static: constructed once (thread-safe "magic static"
	// initialization), lives for the program's duration -- exactly what's needed
	// since the pointer returned below must stay valid for every Form __vbaNew ever
	// constructs afterward.
	static const _FormImpl s_formImplTemplate;

	void * const *pRealVtable = *reinterpret_cast<void * const * const *>(&s_formImplTemplate);

	// _Form's vtable begins with IDispatch's own 7 slots (QueryInterface/AddRef/
	// Release/GetTypeInfoCount/GetTypeInfo/GetIDsOfNames/Invoke) -- this project's
	// own synthesized vtables already have their own 7-slot bridge there instead
	// (vba_BASIC_CLASS_IUnknownBridge, ClassWrapper.hpp), so only the slots after
	// those are wanted here.
	const size_t kIDispatchSlotCount = 7;

	// Matches the number of members _Form declares after IDispatch (Missing1 through
	// OLEDrag, per Form.idl) -- kept as a plain constant since C++ doesn't expose a
	// vtable's slot count directly; a mismatch here from a future Form.idl edit would
	// need _FormImpl above to gain/lose a matching override to keep compiling, which
	// is what actually keeps this constant honest, not the number itself.
	const size_t kFormTailSlotCount = 183;

	*pSlotCount = kFormTailSlotCount;
	return pRealVtable + kIDispatchSlotCount;
}
