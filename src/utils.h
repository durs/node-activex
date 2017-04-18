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

//-------------------------------------------------------------------------------------------------------
#ifndef USE_ATL

class CComVariant : public VARIANT {
public:
    inline CComVariant() { vt = VT_EMPTY; }
    inline CComVariant(const CComVariant &src) { VariantCopyInd(this, &src); }
    inline CComVariant(const VARIANT &src) { VariantCopyInd(this, &src); }
    inline CComVariant(LONG v) { vt = VT_I4; lVal = v; }
    inline ~CComVariant() { if (vt != VT_EMPTY) VariantClear(this); }
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

template <typename T = IUnknown>
class CComPtr {
public:
    T *p;
    inline CComPtr() : p(0) {}
    inline CComPtr(T *_p) : p(0) { Attach(_p); }
    inline CComPtr(const CComPtr<T> &ppt) : p(0) { if (ptr.p) Attach(ptr.p); }
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

inline Local<String> GetWin32ErroroMessage(Isolate *isolate, HRESULT hrcode, LPOLESTR msg, LPOLESTR desc = 0)
{
	uint16_t buf[1024], *bufptr = buf;
	size_t buflen = (sizeof(buf) / sizeof(uint16_t)) - 1;
	if (msg) {
		size_t msglen = wcslen(msg);
		if (msglen > buflen) msglen = buflen;
		if (msglen > 0) memcpy(bufptr, msg, msglen * sizeof(uint16_t));
		buflen -= msglen;
		bufptr += msglen;
		if (buflen > 1) { 
			bufptr[0] = ':'; 
			bufptr[1] = ' '; 
			buflen += 2; 
			bufptr += 2; 
		}
	}
	if (buflen > 0) {
        DWORD len = desc ? wcslen(desc) : 0;
        if (len > 0) {
            if (len >= buflen) len = buflen - 1;
            memcpy(bufptr, desc, len * sizeof(OLECHAR));
        }
        else {
            len = FormatMessageW(FORMAT_MESSAGE_FROM_SYSTEM, 0, hrcode, MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), (LPOLESTR)bufptr, buflen, 0);
            if (len == 0) len = swprintf_s((LPOLESTR)bufptr, buflen, L"Error 0x%08X", hrcode);
        }
		buflen -= len;
		bufptr += len;
	}
	bufptr[0] = 0;
	return String::NewFromTwoByte(isolate, buf);
}

inline Local<Value> Win32Error(Isolate *isolate, HRESULT hrcode, LPOLESTR msg = 0)
{
    return Exception::Error(GetWin32ErroroMessage(isolate, hrcode, msg));
}

inline Local<Value> DispError(Isolate *isolate, HRESULT hrcode, LPOLESTR msg = 0)
{
    CComBSTR desc;
    CComPtr<IErrorInfo> errinfo;
    HRESULT hr = GetErrorInfo(0, &errinfo);
    if (hr == S_OK) errinfo->GetDescription(&desc);
	return Exception::Error(GetWin32ErroroMessage(isolate, hrcode, msg, desc));
}

inline Local<Value> TypeError(Isolate *isolate, const char *msg)
{
    return Exception::TypeError(String::NewFromUtf8(isolate, msg));
}

inline Local<Value> Error(Isolate *isolate, const char *msg)
{
    return Exception::Error(String::NewFromUtf8(isolate, msg));
}

//-------------------------------------------------------------------------------------------------------

inline HRESULT DispFind(IDispatch *disp, LPOLESTR name, DISPID *dispid)
{
	LPOLESTR names[] = { name };
	return disp->GetIDsOfNames(GUID_NULL, names, 1, 0, dispid);
}

inline HRESULT DispInvoke(IDispatch *disp, DISPID dispid, UINT argcnt = 0, VARIANT *args = 0, VARIANT *ret = 0, WORD  flags = DISPATCH_METHOD)
{
	DISPPARAMS params = { args, 0, argcnt, 0 };
	return disp->Invoke(dispid, IID_NULL, 0, flags, &params, ret, 0, 0);
}

inline HRESULT DispInvoke(IDispatch *disp, LPOLESTR name, UINT argcnt = 0, VARIANT *args = 0, VARIANT *ret = 0, WORD  flags = DISPATCH_METHOD, DISPID *dispid = 0)
{
	LPOLESTR names[] = { name };
    DISPID dispids[] = { 0 };
	HRESULT hrcode = disp->GetIDsOfNames(GUID_NULL, names, 1, 0, dispids);
	if SUCCEEDED(hrcode) hrcode = DispInvoke(disp, dispids[0], argcnt, args, ret, flags);
	if (dispid) *dispid = dispids[0];
	return hrcode;
}

//-------------------------------------------------------------------------------------------------------

inline Local<Value> Variant2Value(Isolate *isolate, const CComVariant &v)
{
	VARTYPE vt = (v.vt & VT_TYPEMASK);
	switch (vt)
	{
	case VT_EMPTY:
		return Undefined(isolate);
	case VT_NULL:
		return Null(isolate);
	case VT_I1:
	case VT_I2:
	case VT_I4:
	case VT_INT:
		return Int32::New(isolate, v.lVal);
	case VT_UI1:
	case VT_UI2:
	case VT_UI4:
	case VT_UINT:
		return Uint32::New(isolate, v.ulVal);

	case VT_R4:
		return Number::New(isolate, v.fltVal);

	case VT_R8:
		return Number::New(isolate, v.dblVal);

	case VT_DATE:
		return Date::New(isolate, v.date);

	case VT_BOOL:
		return Boolean::New(isolate, v.boolVal == VARIANT_TRUE);

	case VT_BSTR:
		return String::NewFromTwoByte(isolate, (uint16_t*)(((v.vt & VT_BYREF) != 0) ? *v.pbstrVal : v.bstrVal));
	}
	return Undefined(isolate);
}

inline void Value2Variant(Handle<Value> &val, VARIANT &var)
{
	if (val->IsUint32())
	{
		var.vt = VT_UI4;
		var.ulVal = val->Uint32Value();
	}
	else if (val->IsInt32())
	{
		var.vt = VT_I4;
		var.lVal = val->Int32Value();
	}
	else if (val->IsNumber())
	{
		var.vt = VT_R8;
		var.dblVal = val->NumberValue();
	}
	else if (val->IsDate())
	{
		var.vt = VT_DATE;
		var.date = val->NumberValue();
	}
	else if (val->IsBoolean())
	{
		var.vt = VT_BOOL;
		var.boolVal = val->BooleanValue() ? VARIANT_TRUE : VARIANT_FALSE;
	}
	else if (val->IsUndefined())
	{
		var.vt = VT_EMPTY;
	}
	else if (val->IsNull())
	{
		var.vt = VT_NULL;
	}
	else
	{
		String::Value str(val);
		var.vt = VT_BSTR;
		var.bstrVal = (str.length() > 0) ? SysAllocString((LPOLESTR)*str) : 0;
	}
}

inline bool VariantDispGet(VARIANT *v, IDispatch **disp)
{
	if ((v->vt & VT_TYPEMASK) == VT_DISPATCH)
	{
		*disp = ((v->vt & VT_BYREF) != 0) ? *v->ppdispVal : v->pdispVal;
		if (*disp) (*disp)->AddRef();
		return true;
	}
	if ((v->vt & VT_TYPEMASK) == VT_UNKNOWN)
	{
		IUnknown *unk = ((v->vt & VT_BYREF) != 0) ? *v->ppunkVal : v->punkVal;
		if (!unk || FAILED(unk->QueryInterface(__uuidof(IDispatch), (void**)disp))) *disp = 0;
		return true;
	}
	return false;
}

//-------------------------------------------------------------------------------------------------------

inline bool v8val2bool(const Local<Value> &v, bool def) {
    if (v.IsEmpty()) return def;
    if (v->IsBoolean()) return v->BooleanValue();
    if (v->IsInt32()) return v->Int32Value() != 0;
    if (v->IsUint32()) return v->Uint32Value() != 0;
    return def;
}

//-------------------------------------------------------------------------------------------------------

class VarArgumets
{
public:
	std::vector<CComVariant> items;
	VarArgumets(Local<Value> value)
	{
		items.resize(1);
		Value2Variant(value, items[0]);
	}
	VarArgumets(const FunctionCallbackInfo<Value> &args)
	{
		int argcnt = args.Length();
		items.resize(argcnt);
		for (int i = 0; i < argcnt; i ++)
			Value2Variant(args[argcnt - i - 1], items[i]);
	}
};

//-------------------------------------------------------------------------------------------------------
