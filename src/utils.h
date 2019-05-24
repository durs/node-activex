//-------------------------------------------------------------------------------------------------------
// Project: NodeActiveX
// Author: Yuri Dursin
// Description: Common utilities for translation COM - NodeJS
//-------------------------------------------------------------------------------------------------------

#pragma once

//-------------------------------------------------------------------------------------------------------

#ifdef _DEBUG
#define NODE_DEBUG
#endif

#if V8_MAJOR_VERSION >= 7
#if V8_MINOR_VERSION > 0	// node v11.13
	#define NODE_BOOL_ISOLATE
#endif
#endif

#ifdef NODE_BOOL_ISOLATE
#define NODE_BOOL(isolate, v) v->BooleanValue(isolate)
#else
#define NODE_BOOL(isolate, v) v->BooleanValue(isolate->GetCurrentContext()).FromMaybe(false)
#endif

#ifdef NODE_DEBUG
	#define NODE_DEBUG_PREFIX "### "
	#define NODE_DEBUG_MSG(msg) { printf(NODE_DEBUG_PREFIX"%s", msg); std::cout << std::endl; }
	#define NODE_DEBUG_FMT(msg, arg) { std::cout << NODE_DEBUG_PREFIX; printf(msg, arg); std::cout << std::endl; }
	#define NODE_DEBUG_FMT2(msg, arg, arg2) { std::cout << NODE_DEBUG_PREFIX; printf(msg, arg, arg2); std::cout << std::endl; }
#else
	#define NODE_DEBUG_MSG(msg)
	#define NODE_DEBUG_FMT(msg, arg)
	#define NODE_DEBUG_FMT2(msg, arg, arg2)
#endif

inline Local<String> v8str(Isolate *isolate, const char *text) {
	Local<String> str;
	if (!text || !String::NewFromUtf8(isolate, text, NewStringType::kNormal).ToLocal(&str)) {
		str = String::Empty(isolate);
	}
	return str;
}

inline Local<String> v8str(Isolate *isolate, const uint16_t *text) {
	Local<String> str;
	if (!text || !String::NewFromTwoByte(isolate, (const uint16_t*)text, NewStringType::kNormal).ToLocal(&str)) {
		str = String::Empty(isolate);
	}
	return str;
}

inline Local<String> v8str(Isolate *isolate, const wchar_t *text) {
	return v8str(isolate, (const uint16_t *)text);
}

//-------------------------------------------------------------------------------------------------------
#ifndef USE_ATL

class CComVariant : public VARIANT {
public:
    inline CComVariant() { 
		memset((VARIANT*)this, 0, sizeof(VARIANT));
	}
    inline CComVariant(const CComVariant &src) { 
		memset((VARIANT*)this, 0, sizeof(VARIANT));
		VariantCopyInd(this, &src);
	}
    inline CComVariant(const VARIANT &src) { 
		memset((VARIANT*)this, 0, sizeof(VARIANT));
		VariantCopyInd(this, &src);
	}
    inline CComVariant(LONG v) { 
		memset((VARIANT*)this, 0, sizeof(VARIANT));
		vt = VT_I4;
		lVal = v; 
	}
	inline CComVariant(LPOLESTR v) {
		memset((VARIANT*)this, 0, sizeof(VARIANT));
		vt = VT_BSTR;
		bstrVal = SysAllocString(v);
	}
	inline ~CComVariant() {
		Clear(); 
	}
    inline void Clear() {
        if (vt != VT_EMPTY)
            VariantClear(this);
    }
    inline void Detach(VARIANT *dst) {
        *dst = *this;
        vt = VT_EMPTY;
    }
	inline HRESULT CopyTo(VARIANT *dst) {
		return VariantCopy(dst, this);
	}

	inline HRESULT ChangeType(VARTYPE vtNew, const VARIANT* pSrc = NULL) {
		return VariantChangeType(this, pSrc ? pSrc : this, 0, vtNew);
	}

	inline ULONG ArrayLength() {
		if ((vt & VT_ARRAY) == 0) return 0;
		SAFEARRAY *varr = (vt & VT_BYREF) != 0 ? *pparray : parray;
		return varr ? varr->rgsabound[0].cElements : 0;
	}

	inline HRESULT ArrayGet(LONG index, CComVariant &var) {
		if ((vt & VT_ARRAY) == 0) return E_NOTIMPL;
		SAFEARRAY *varr = (vt & VT_BYREF) != 0 ? *pparray : parray;
		if (!varr) return E_FAIL;
		index += varr->rgsabound[0].lLbound;
		VARTYPE vart = vt & VT_TYPEMASK;
		HRESULT hr = SafeArrayGetElement(varr, &index, (vart == VT_VARIANT) ? (void*)&var : (void*)&var.byref);
		if (SUCCEEDED(hr) && vart != VT_VARIANT) var.vt = vart;
		return hr;
	}
	template<typename T>
	inline T* ArrayGet(ULONG index = 0) {
		return ((T*)parray->pvData) + index;
	}
	inline HRESULT ArrayCreate(VARTYPE avt, ULONG cnt) {
		Clear();
		parray = SafeArrayCreateVector(avt, 0, cnt);
		if (!parray) return E_UNEXPECTED;
		vt = VT_ARRAY | avt;
		return S_OK;
	}
	inline HRESULT ArrayResize(ULONG cnt) {
		SAFEARRAYBOUND bnds = { cnt, 0 };
		return SafeArrayRedim(parray, &bnds);
	}
};

class CComBSTR {
public:
    BSTR p;
    inline CComBSTR() : p(0) {}
    inline CComBSTR(const CComBSTR &src) : p(0) {}
    inline ~CComBSTR() { Free(); }
    inline void Attach(BSTR _p) { Free(); p = _p; }
    inline BSTR Detach() { BSTR pp = p; p = 0; return pp; }
    inline void Free() { if (p) { SysFreeString(p); p = 0; } }

    inline operator BSTR () const { return p; }
    inline BSTR* operator&() { return &p; }
    inline bool operator!() const { return (p == 0); }
    inline bool operator!=(BSTR _p) const { return !operator==(_p); }
    inline bool operator==(BSTR _p) const { return p == _p; }
    inline BSTR operator = (BSTR _p) {
        if (p != _p) Attach(_p ? SysAllocString(_p) : 0);
        return p;
    }
};

class CComException: public EXCEPINFO {
public:
	inline CComException() {
		memset((EXCEPINFO*)this, 0, sizeof(EXCEPINFO));
	}
	inline ~CComException() {
		Clear(true);
	}
	inline void Clear(bool internal = false) {
		if (bstrSource) SysFreeString(bstrSource);
		if (bstrDescription) SysFreeString(bstrDescription);
		if (bstrHelpFile) SysFreeString(bstrHelpFile);
		if (!internal) memset((EXCEPINFO*)this, 0, sizeof(EXCEPINFO));
	}
};

template <typename T = IUnknown>
class CComPtr {
public:
    T *p;
    inline CComPtr() : p(0) {}
    inline CComPtr(T *_p) : p(0) { Attach(_p); }
    inline CComPtr(const CComPtr<T> &ptr) : p(0) { if (ptr.p) Attach(ptr.p); }
    inline ~CComPtr() { Release(); }

    inline void Attach(T *_p) { Release(); p = _p; if (p) p->AddRef(); }
    inline T *Detach() { T *pp = p; p = 0; return pp; }
    inline void Release() { if (p) { p->Release(); p = 0; } }

    inline operator T*() const { return p; }
    inline T* operator->() const { return p; }
    inline T& operator*() const { return *p; }
    inline T** operator&() { return &p; }
    inline bool operator!() const { return (p == 0); }
    inline bool operator!=(T* _p) const { return !operator==(_p); }
    inline bool operator==(T* _p) const { return p == _p; }
    inline T* operator = (T* _p) {
        if (p != _p) Attach(_p);
        return p;
    }

    inline HRESULT CoCreateInstance(REFCLSID rclsid, LPUNKNOWN pUnkOuter = NULL, DWORD dwClsContext = CLSCTX_ALL) {
        Release();
        return ::CoCreateInstance(rclsid, pUnkOuter, dwClsContext, __uuidof(T), (void**)&p);
    }

    inline HRESULT CoCreateInstance(LPCOLESTR szProgID, LPUNKNOWN pUnkOuter = NULL, DWORD dwClsContext = CLSCTX_ALL) {
        Release();
        CLSID clsid;
        HRESULT hr = CLSIDFromProgID(szProgID, &clsid);
        if FAILED(hr) return hr;
        return ::CoCreateInstance(clsid, pUnkOuter, dwClsContext, __uuidof(T), (void**)&p);
    }
};

#endif
//-------------------------------------------------------------------------------------------------------

Local<String> GetWin32ErroroMessage(Isolate *isolate, HRESULT hrcode, LPCOLESTR msg, LPCOLESTR msg2 = 0, LPCOLESTR desc = 0);

inline Local<Value> Win32Error(Isolate *isolate, HRESULT hrcode, LPCOLESTR id = 0, LPCOLESTR msg = 0) {
	auto err = Exception::Error(GetWin32ErroroMessage(isolate, hrcode, id, msg));
	auto obj = Local<Object>::Cast(err);
	obj->Set(isolate->GetCurrentContext(), v8str(isolate, "errno"), Integer::New(isolate, hrcode));
	return err;
}

inline Local<Value> DispError(Isolate *isolate, HRESULT hrcode, LPCOLESTR id = 0, LPCOLESTR msg = 0, EXCEPINFO *except = 0) {
	Local<Context> ctx = isolate->GetCurrentContext();
	CComBSTR desc;
    CComPtr<IErrorInfo> errinfo;
    HRESULT hr = GetErrorInfo(0, &errinfo);
    if (hr == S_OK) errinfo->GetDescription(&desc);
	auto err = Exception::Error(GetWin32ErroroMessage(isolate, hrcode, id, msg, desc));
	auto obj = Local<Object>::Cast(err);
	obj->Set(ctx, v8str(isolate, "errno"), Integer::New(isolate, hrcode));
	if (except) {
		if (except->scode != 0) obj->Set(ctx, v8str(isolate, "code"), Integer::New(isolate, except->scode));
		else if (except->wCode != 0) obj->Set(ctx, v8str(isolate, "code"), Integer::New(isolate, except->wCode));
		if (except->bstrSource != 0) obj->Set(ctx, v8str(isolate, "source"), v8str(isolate, except->bstrSource));
		if (except->bstrDescription != 0) obj->Set(ctx, v8str(isolate, "description"), v8str(isolate, except->bstrDescription));
	}
	return err;
}

inline Local<Value> DispErrorNull(Isolate *isolate) {
    return Exception::TypeError(v8str(isolate, "DispNull"));
}

inline Local<Value> DispErrorInvalid(Isolate *isolate) {
    return Exception::TypeError(v8str(isolate, "DispInvalid"));
}

inline Local<Value> TypeError(Isolate *isolate, const char *msg) {
    return Exception::TypeError(v8str(isolate, msg));
}

inline Local<Value> InvalidArgumentsError(Isolate *isolate) {
    return Exception::TypeError(v8str(isolate, "Invalid arguments"));
}

inline Local<Value> Error(Isolate *isolate, const char *msg) {
    return Exception::Error(v8str(isolate, msg));
}

//-------------------------------------------------------------------------------------------------------

inline HRESULT DispFind(IDispatch *disp, LPOLESTR name, DISPID *dispid) {
	LPOLESTR names[] = { name };
	return disp->GetIDsOfNames(GUID_NULL, names, 1, 0, dispid);
}

inline HRESULT DispInvoke(IDispatch *disp, DISPID dispid, UINT argcnt = 0, VARIANT *args = 0, VARIANT *ret = 0, WORD  flags = DISPATCH_METHOD, EXCEPINFO *except = 0) {
	DISPPARAMS params = { args, 0, argcnt, 0 };
	DISPID dispidNamed = DISPID_PROPERTYPUT;
	if (flags == DISPATCH_PROPERTYPUT) { // It`s a magic
		params.cNamedArgs = 1;
		params.rgdispidNamedArgs = &dispidNamed;
	}
	return disp->Invoke(dispid, IID_NULL, 0, flags, &params, ret, except, 0);
}

inline HRESULT DispInvoke(IDispatch *disp, LPOLESTR name, UINT argcnt = 0, VARIANT *args = 0, VARIANT *ret = 0, WORD  flags = DISPATCH_METHOD, DISPID *dispid = 0, EXCEPINFO *except = 0) {
	LPOLESTR names[] = { name };
    DISPID dispids[] = { 0 };
	HRESULT hrcode = disp->GetIDsOfNames(GUID_NULL, names, 1, 0, dispids);
	if SUCCEEDED(hrcode) hrcode = DispInvoke(disp, dispids[0], argcnt, args, ret, flags, except);
	if (dispid) *dispid = dispids[0];
	return hrcode;
}

//-------------------------------------------------------------------------------------------------------

template<typename INTTYPE>
inline INTTYPE Variant2Int(const VARIANT &v, const INTTYPE def) {
    VARTYPE vt = (v.vt & VT_TYPEMASK);
	bool by_ref = (v.vt & VT_BYREF) != 0;
    switch (vt) {
    case VT_EMPTY:
    case VT_NULL:
        return def;
    case VT_I1:
    case VT_I2:
    case VT_I4:
    case VT_INT:
        return (INTTYPE)(by_ref ? *v.plVal : v.lVal);
    case VT_UI1:
    case VT_UI2:
    case VT_UI4:
    case VT_UINT:
        return (INTTYPE)(by_ref ? *v.pulVal : v.ulVal);
	case VT_CY:
		return (INTTYPE)((by_ref ? v.pcyVal : &v.cyVal)->int64 / 10000);
	case VT_R4:
        return (INTTYPE)(by_ref ? *v.pfltVal : v.fltVal);
    case VT_R8:
        return (INTTYPE)(by_ref ? *v.pdblVal : v.dblVal);
    case VT_DATE:
        return (INTTYPE)(by_ref ? *v.pdate : v.date);
	case VT_DECIMAL: {
		LONG64 int64val;
		return SUCCEEDED(VarI8FromDec(by_ref ? v.pdecVal : &v.decVal, &int64val)) ? (INTTYPE)int64val : def; 
	}
	case VT_BOOL:
        return (v.boolVal == VARIANT_TRUE) ? 1 : 0;
	case VT_VARIANT:
		if (v.pvarVal) return Variant2Int<INTTYPE>(*v.pvarVal, def);
	}
    VARIANT dst;
    return SUCCEEDED(VariantChangeType(&dst, &v, 0, VT_INT)) ? (INTTYPE)dst.intVal : def;
}

Local<Value> Variant2Array(Isolate *isolate, const VARIANT &v);
Local<Value> Variant2Value(Isolate *isolate, const VARIANT &v, bool allow_disp = false);
Local<Value> Variant2String(Isolate *isolate, const VARIANT &v);
void Value2Variant(Isolate *isolate, Local<Value> &val, VARIANT &var, VARTYPE vt = VT_EMPTY);
bool Value2Unknown(Isolate *isolate, Local<Value> &val, IUnknown **unk);
bool VariantUnkGet(VARIANT *v, IUnknown **unk);
bool VariantDispGet(VARIANT *v, IDispatch **disp);
bool UnknownDispGet(IUnknown *unk, IDispatch **disp);

//-------------------------------------------------------------------------------------------------------

inline bool v8val2bool(Isolate *isolate, const Local<Value> &v, bool def) {
	Local<Context> ctx = isolate->GetCurrentContext();
    if (v.IsEmpty()) return def;
    if (v->IsBoolean()) return NODE_BOOL(isolate, v);
    if (v->IsInt32()) return v->Int32Value(ctx).FromMaybe(def ? 1 : 0) != 0;
    if (v->IsUint32()) return v->Uint32Value(ctx).FromMaybe(def ? 1 : 0) != 0;
    return def;
}

//-------------------------------------------------------------------------------------------------------

class VarArguments {
public:
	std::vector<CComVariant> items;
	VarArguments() {}
	VarArguments(Isolate *isolate, Local<Value> value) {
		items.resize(1);
		Value2Variant(isolate, value, items[0]);
	}
	VarArguments(Isolate *isolate, const FunctionCallbackInfo<Value> &args) {
		int argcnt = args.Length();
		items.resize(argcnt);
		for (int i = 0; i < argcnt; i ++)
			Value2Variant(isolate, args[argcnt - i - 1], items[i]);
	}
    inline bool IsDefault() {
        if (items.size() != 1) return false;
        auto &arg = items[0];
        if (arg.vt != VT_BSTR || arg.bstrVal == nullptr) return false;
        return wcscmp(arg.bstrVal, L"default") == 0;
    }
};

class NodeArguments {
public:
	std::vector<Local<Value>> items;
	NodeArguments(Isolate *isolate, DISPPARAMS *pDispParams, bool allow_disp, bool reverse_arguments = true) {
		UINT argcnt = pDispParams->cArgs;
		items.resize(argcnt);
		for (UINT i = 0; i < argcnt; i++) {
			items[i] = Variant2Value(isolate, pDispParams->rgvarg[reverse_arguments ? argcnt - i - 1 : i], allow_disp);
		}
	}
};

//-------------------------------------------------------------------------------------------------------

template<typename IBASE = IUnknown>
class UnknownImpl : public IBASE {
public:
	inline UnknownImpl() : refcnt(0) {}
	virtual ~UnknownImpl() {}

	// IUnknown interface
	virtual HRESULT __stdcall QueryInterface(REFIID qiid, void **ppvObject) {
		if ((qiid == IID_IUnknown) || (qiid == __uuidof(IBASE))) {
			*ppvObject = this;
			AddRef();
			return S_OK;
		}
		return E_NOINTERFACE;
	}

	virtual ULONG __stdcall AddRef() {
		return InterlockedIncrement(&refcnt);
	}

	virtual ULONG __stdcall Release() {
		if (InterlockedDecrement(&refcnt) != 0) return refcnt;
		delete this;
		return 0;
	}

protected:
	LONG refcnt;

};

class DispArrayImpl : public UnknownImpl<IDispatch> {
public:
	CComVariant var;
	DispArrayImpl(const VARIANT &v): var(v) {}

	// IDispatch interface
	virtual HRESULT STDMETHODCALLTYPE GetTypeInfoCount(UINT *pctinfo) { *pctinfo = 0; return S_OK; }
	virtual HRESULT STDMETHODCALLTYPE GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) { return E_NOTIMPL; }
	virtual HRESULT STDMETHODCALLTYPE GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId);
	virtual HRESULT STDMETHODCALLTYPE Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr);
};

class DispEnumImpl : public UnknownImpl<IDispatch> {
public:
    CComPtr<IEnumVARIANT> ptr;
    DispEnumImpl() {}
    DispEnumImpl(IEnumVARIANT *p) : ptr(p) {}

    // IDispatch interface
    virtual HRESULT STDMETHODCALLTYPE GetTypeInfoCount(UINT *pctinfo) { *pctinfo = 0; return S_OK; }
    virtual HRESULT STDMETHODCALLTYPE GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) { return E_NOTIMPL; }
    virtual HRESULT STDMETHODCALLTYPE GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId);
    virtual HRESULT STDMETHODCALLTYPE Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr);
};


// {9DCE8520-2EFE-48C0-A0DC-951B291872C0}
extern const GUID CLSID_DispObjectImpl;

class DispObjectImpl : public UnknownImpl<IDispatch> {
public:
	Persistent<Object> obj;

	struct name_t { 
		DISPID dispid;
		std::wstring name;
		inline name_t(DISPID id, const std::wstring &nm): dispid(id), name(nm) {}
	};
	typedef std::shared_ptr<name_t> name_ptr;
	typedef std::map<std::wstring, name_ptr> names_t;
	typedef std::map<DISPID, name_ptr> index_t;
	DISPID dispid_next;
	names_t names;
	index_t index;
	bool reverse_arguments;

	inline DispObjectImpl(const Local<Object> &_obj, bool revargs = true) : obj(Isolate::GetCurrent(), _obj), dispid_next(1), reverse_arguments(revargs){}
	virtual ~DispObjectImpl() { obj.Reset(); }

	// IUnknown interface
	virtual HRESULT __stdcall QueryInterface(REFIID qiid, void **ppvObject) {
		if (qiid == CLSID_DispObjectImpl) { *ppvObject = this; return S_OK; }
		return UnknownImpl<IDispatch>::QueryInterface(qiid, ppvObject);
	}

	// IDispatch interface
	virtual HRESULT STDMETHODCALLTYPE GetTypeInfoCount(UINT *pctinfo) { *pctinfo = 0; return S_OK; }
	virtual HRESULT STDMETHODCALLTYPE GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) { return E_NOTIMPL; }
	virtual HRESULT STDMETHODCALLTYPE GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId);
	virtual HRESULT STDMETHODCALLTYPE Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr);
};

double FromOleDate(double);
double ToOleDate(double);

//-------------------------------------------------------------------------------------------------------
