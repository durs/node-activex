//-------------------------------------------------------------------------------------------------------
// Project: NodeActiveX
// Author: Yuri Dursin
// Last Modification: 2011-11-20
// Description: Common utilities for translation COM - NodeJS
//-------------------------------------------------------------------------------------------------------

#pragma once

//-------------------------------------------------------------------------------------------------------

#ifdef _DEBUG
#define NODE_DEBUG
#endif

#ifdef NODE_DEBUG
#define NODE_DEBUG_PREFIX "n# "
#define NODE_DEBUG_MSG(msg) { printf(NODE_DEBUG_PREFIX"%s", msg); cout << endl; }
#define NODE_DEBUG_FMT(msg, arg) { cout << NODE_DEBUG_PREFIX; printf(msg, arg); cout << endl; }
#define NODE_DEBUG_FMT2(msg, arg, arg2) { cout << NODE_DEBUG_PREFIX; printf(msg, arg, arg2); cout << endl; }
#else
#define NODE_DEBUG_MSG(msg)
#define NODE_DEBUG_FMT(msg, arg)
#define NODE_DEBUG_FMT2(msg, arg, arg2)
#endif

//-------------------------------------------------------------------------------------------------------

inline Handle<Value> GetWin32ErroroMessage(HRESULT hrcode, LPOLESTR msg = 0)
{
	uint16_t buf[1024], *bufptr = buf;
	size_t buflen = (sizeof(buf) / sizeof(uint16_t)) - 1;
	if (msg)
	{
		size_t msglen = wcslen(msg);
		if (msglen > buflen) msglen = buflen;
		if (msglen > 0) memcpy(bufptr, msg, msglen * sizeof(uint16_t));
		buflen -= msglen;
		bufptr += msglen;
		if (buflen > 1) 
		{ 
			bufptr[0] = ':'; 
			bufptr[1] = ' '; 
			buflen += 2; 
			bufptr += 2; 
		}
	}
	if (buflen > 0)
	{
		DWORD len = FormatMessageW(FORMAT_MESSAGE_FROM_SYSTEM, 0, hrcode, MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), (LPOLESTR)bufptr, buflen, 0);
		if (len == 0) len = swprintf_s((LPOLESTR)bufptr, buflen, L"Error 0x%08X", hrcode);
		buflen -= len;
		bufptr += len;
	}
	bufptr[0] = 0;
	return String::New(buf);
}

inline Handle<Value> DispError(HRESULT hrcode, LPOLESTR msg = 0)
{
	return ThrowException(GetWin32ErroroMessage(hrcode, msg));
}

//-------------------------------------------------------------------------------------------------------

inline HRESULT DispFind(IDispatch *disp, LPOLESTR name, DISPID *dispid)
{
	LPOLESTR names[] = { name };
	return disp->GetIDsOfNames(GUID_NULL, names, 1, 0, dispid);
}

inline HRESULT DispInvoke(IDispatch *disp, LONG dispid, LONG argcnt = 0, VARIANT *args = 0, VARIANT *ret = 0, WORD  flags = DISPATCH_METHOD)
{
	DISPPARAMS params = { args, 0, argcnt, 0 };
	return disp->Invoke(dispid, IID_NULL, 0, flags, &params, ret, 0, 0);
}

inline HRESULT DispInvoke(IDispatch *disp, LPOLESTR name, LONG argcnt = 0, VARIANT *args = 0, VARIANT *ret = 0, WORD  flags = DISPATCH_METHOD, DISPID *dispid = 0)
{
	LPOLESTR names[] = { name };
	LONG dispids[] = { 0 };
	HRESULT hrcode = disp->GetIDsOfNames(GUID_NULL, names, 1, 0, dispids);
	if SUCCEEDED(hrcode) hrcode = DispInvoke(disp, dispids[0], argcnt, args, ret, flags);
	if (dispid) *dispid = dispids[0];
	return hrcode;
}

//-------------------------------------------------------------------------------------------------------

inline Handle<Value> Variant2Value(CComVariant &v)
{
	VARTYPE vt = (v.vt & VT_TYPEMASK);
	switch (vt)
	{
	case VT_EMPTY:
		return Undefined();
	case VT_NULL:
		return Null();
	case VT_I1:
	case VT_I2:
	case VT_I4:
	case VT_INT:
		return Int32::New(v.lVal);
	case VT_UI1:
	case VT_UI2:
	case VT_UI4:
	case VT_UINT:
		return Uint32::New(v.ulVal);

	case VT_R4:
		return Number::New(v.fltVal);

	case VT_R8:
		return Number::New(v.dblVal);

	case VT_DATE:
		return Date::New(v.date);

	case VT_BOOL:
		return Boolean::New(v.boolVal == VARIANT_TRUE);

	case VT_BSTR:
		return String::New((uint16_t*)(((v.vt & VT_BYREF) != 0) ? *v.pbstrVal : v.bstrVal));
	}
	return Undefined();
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

class VarArgumets
{
public:
	std::vector<CComVariant> items;
	VarArgumets(Local<Value> value)
	{
		items.resize(1);
		Value2Variant(value, items[0]);
	}
	VarArgumets(const Arguments &args)
	{
		int argcnt = args.Length();
		items.resize(argcnt);
		for (int i = 0; i < argcnt; i ++)
			Value2Variant(args[argcnt - i - 1], items[i]);
	}
};

//-------------------------------------------------------------------------------------------------------
