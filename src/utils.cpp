//-------------------------------------------------------------------------------------------------------
// Project: NodeActiveX
// Author: Yuri Dursin
// Description:  Common utilities implementations
//-------------------------------------------------------------------------------------------------------

#include "stdafx.h"
#include "disp.h"

const GUID CLSID_DispObjectImpl = { 0x9dce8520, 0x2efe, 0x48c0,{ 0xa0, 0xdc, 0x95, 0x1b, 0x29, 0x18, 0x72, 0xc0 } };

//-------------------------------------------------------------------------------------------------------

Local<String> GetWin32ErroroMessage(Isolate *isolate, HRESULT hrcode, LPCOLESTR msg, LPCOLESTR msg2, LPCOLESTR desc) {
	uint16_t buf[1024], *bufptr = buf;
	size_t len, buflen = (sizeof(buf) / sizeof(uint16_t)) - 1;
	if (msg) {
		len = wcslen(msg);
		if (len > buflen) len = buflen;
		if (len > 0) memcpy(bufptr, msg, len * sizeof(uint16_t));
		buflen -= len;
		bufptr += len;
		if (buflen > 1) {
			bufptr[0] = ':';
			bufptr[1] = ' ';
			buflen -= 2;
			bufptr += 2;
		}
	}
	if (msg2) {
		len = wcslen(msg2);
		if (len > buflen) len = buflen;
		if (len > 0) memcpy(bufptr, msg2, len * sizeof(uint16_t));
		buflen -= len;
		bufptr += len;
		if (buflen > 0) {
			bufptr[0] = ' ';
			buflen -= 1;
			bufptr += 1;
		}
	}
	if (buflen > 0) {
		len = desc ? wcslen(desc) : 0;
		if (len > 0) {
			if (len >= buflen) len = buflen - 1;
			memcpy(bufptr, desc, len * sizeof(OLECHAR));
		}
		else {
			len = FormatMessageW(FORMAT_MESSAGE_FROM_SYSTEM, 0, hrcode, MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), (LPOLESTR)bufptr, (DWORD)buflen, 0);
			if (len == 0) len = swprintf_s((LPOLESTR)bufptr, buflen, L"Error 0x%08X", hrcode);
		}
		buflen -= len;
		bufptr += len;
	}
	bufptr[0] = 0;
	return String::NewFromTwoByte(isolate, buf);
}

//-------------------------------------------------------------------------------------------------------

Local<Value> Variant2Array(Isolate *isolate, const VARIANT &v) {
	if ((v.vt & VT_ARRAY) == 0) return Null(isolate);
	SAFEARRAY *varr = (v.vt & VT_BYREF) != 0 ? *v.pparray : v.parray;
	if (!varr || varr->cDims != 1) return Null(isolate);
	VARTYPE vt = v.vt & VT_TYPEMASK;
	LONG cnt = (LONG)varr->rgsabound[0].cElements;
	Local<Array> arr = Array::New(isolate, cnt);
	for (LONG i = varr->rgsabound[0].lLbound; i < cnt; i++) {
		CComVariant vi;
		if SUCCEEDED(SafeArrayGetElement(varr, &i, (vt == VT_VARIANT) ? (void*)&vi : (void*)&vi.byref)) {
			if (vt != VT_VARIANT) vi.vt = vt;
			arr->Set((uint32_t)i, Variant2Value(isolate, vi, true));
		}
	}
	return arr;
}

Local<Value> Variant2Value(Isolate *isolate, const VARIANT &v, bool allow_disp) {
	if ((v.vt & VT_ARRAY) != 0) return Variant2Array(isolate, v);
	VARTYPE vt = (v.vt & VT_TYPEMASK);
	bool by_ref = (v.vt & VT_BYREF) != 0;
	switch (vt) {
	case VT_NULL:
		return Null(isolate);
	case VT_I1:
	case VT_I2:
	case VT_I4:
	case VT_INT:
		return Int32::New(isolate, by_ref ? *v.plVal : v.lVal);
	case VT_UI1:
	case VT_UI2:
	case VT_UI4:
	case VT_UINT:
		return Uint32::New(isolate, by_ref ? *v.pulVal : v.ulVal);
	case VT_R4:
		return Number::New(isolate, by_ref ? *v.pfltVal : v.fltVal);
	case VT_R8:
		return Number::New(isolate, by_ref ? *v.pdblVal : v.dblVal);
	case VT_DATE:
		return Date::New(isolate, by_ref ? *v.pdate : v.date);
	case VT_BOOL:
		return Boolean::New(isolate, (by_ref ? *v.pboolVal : v.boolVal) == VARIANT_TRUE);
	case VT_DISPATCH: {
		IDispatch *disp = (by_ref ? *v.ppdispVal : v.pdispVal);
		if (!disp) return Null(isolate);
		if (allow_disp) {
			DispObjectImpl *impl;
			if (disp->QueryInterface(CLSID_DispObjectImpl, (void**)&impl) == S_OK) {
				return impl->obj.Get(isolate);
			}
			return DispObject::NodeCreate(isolate, disp, L"Dispatch", option_auto);
		}
		return String::NewFromUtf8(isolate, "[Dispatch]");
	}
	case VT_UNKNOWN: {
		CComPtr<IDispatch> disp;
		if (allow_disp && UnknownDispGet(((v.vt & VT_BYREF) != 0) ? *v.ppunkVal : v.punkVal, &disp)) {
			return DispObject::NodeCreate(isolate, disp, L"Unknown", option_auto);
		}
		return String::NewFromUtf8(isolate, "[Unknown]");
	}
	case VT_BSTR: {
        BSTR bstr = by_ref ? (v.pbstrVal ? *v.pbstrVal : nullptr) : v.bstrVal;
        if (!bstr) return String::Empty(isolate);
		return String::NewFromTwoByte(isolate, (uint16_t*)bstr);
    }
	case VT_VARIANT: 
		if (v.pvarVal) return Variant2Value(isolate, *v.pvarVal, allow_disp);
	}
	return Undefined(isolate);
}

Local<Value> Variant2String(Isolate *isolate, const VARIANT &v) {
	char buf[256] = {};
	VARTYPE vt = (v.vt & VT_TYPEMASK);
	bool by_ref = (v.vt & VT_BYREF) != 0;
	switch (vt) {
	case VT_EMPTY:
		strcpy(buf, "EMPTY");
		break;
	case VT_NULL:
		strcpy(buf, "NULL");
		break;
	case VT_I1:
	case VT_I2:
	case VT_I4:
	case VT_INT:
		sprintf_s(buf, "%i", (int)(by_ref ? *v.plVal : v.lVal));
		break;
	case VT_UI1:
	case VT_UI2:
	case VT_UI4:
	case VT_UINT:
		sprintf_s(buf, "%u", (unsigned int)(by_ref ? *v.pulVal : v.ulVal));
		break;
	case VT_R4:
		sprintf_s(buf, "%f", (double)(by_ref ? *v.pfltVal : v.fltVal));
		break;
	case VT_R8:
		sprintf_s(buf, "%f", (double)(by_ref ? *v.pdblVal : v.dblVal));
		break;
	case VT_DATE:
		return Date::New(isolate, by_ref ? *v.pdate : v.date);
	case VT_BOOL:
		strcpy(buf, ((by_ref ? *v.pboolVal : v.boolVal) == VARIANT_TRUE) ? "true" : "false");
	case VT_DISPATCH:
		strcpy(buf, "[Dispatch]");
		break;
	case VT_UNKNOWN: 
		strcpy(buf, "[Unknown]");
		break;
	case VT_VARIANT:
		if (v.pvarVal) return Variant2String(isolate, *v.pvarVal);
		break;
	default:
		CComVariant tmp;
		if (SUCCEEDED(VariantChangeType(&tmp, &v, 0, VT_BSTR)) && tmp.vt == VT_BSTR && v.bstrVal != nullptr) {
			return String::NewFromTwoByte(isolate, (uint16_t*)v.bstrVal);
		}
	}
	return String::NewFromUtf8(isolate, buf, String::kNormalString);
}

void Value2Variant(Isolate *isolate, Local<Value> &val, VARIANT &var, VARTYPE vt) {
	if (val.IsEmpty() || val->IsUndefined()) {
		var.vt = VT_EMPTY;
	}
	else if (val->IsNull()) {
		var.vt = VT_NULL;
	}
	else if (val->IsInt32()) {
		var.vt = VT_I4;
		var.lVal = val->Int32Value();
	}
	else if (val->IsUint32()) {
		var.ulVal = val->Uint32Value();
		var.vt = (var.ulVal <= 0x7FFFFFFF) ? VT_I4 : VT_UI4;
	}
	else if (val->IsNumber()) {
		var.vt = VT_R8;
		var.dblVal = val->NumberValue();
	}
	else if (val->IsDate()) {
		var.vt = VT_DATE;
		var.date = val->NumberValue();
	}
	else if (val->IsBoolean()) {
		var.vt = VT_BOOL;
		var.boolVal = val->BooleanValue() ? VARIANT_TRUE : VARIANT_FALSE;
	}
	else if (val->IsArray() && (vt != VT_NULL)) {
		Local<Array> arr = v8::Local<Array>::Cast(val);
		uint32_t len = arr->Length();
		if (vt == VT_EMPTY) vt = VT_VARIANT;
		var.vt = VT_ARRAY | vt;
		var.parray = SafeArrayCreateVector(vt, 0, len);
		for (uint32_t i = 0; i < len; i++) {
			CComVariant v;
			Value2Variant(isolate, arr->Get(i), v, vt);
			void *pv;
			if (vt == VT_VARIANT) pv = (void*)&v;
			else if (vt == VT_DISPATCH || vt == VT_UNKNOWN || vt == VT_BSTR) pv = v.byref;
			else pv = (void*)&v.byref;
			SafeArrayPutElement(var.parray, (LONG*)&i, pv);
		}
		vt = VT_EMPTY;
	}
	else if (val->IsObject()) {
		Local<Object> obj = val->ToObject();
		if (!DispObject::GetValueOf(isolate, obj, var) && !VariantObject::GetValueOf(isolate, obj, var)) {
			var.vt = VT_DISPATCH;
			var.pdispVal = new DispObjectImpl(obj);
			var.pdispVal->AddRef();
		}
	}
	else {
		String::Value str(val);
		var.vt = VT_BSTR;
		var.bstrVal = (str.length() > 0) ? SysAllocString((LPOLESTR)*str) : 0;
	}
	if (vt != VT_EMPTY && vt != VT_NULL && vt != VT_VARIANT) {
		if FAILED(VariantChangeType(&var, &var, 0, vt))
			VariantClear(&var);
	}
}

bool UnknownDispGet(IUnknown *unk, IDispatch **disp) {
	if (!unk) return false;
	if SUCCEEDED(unk->QueryInterface(__uuidof(IDispatch), (void**)disp)) {
		return true;
	}
	CComPtr<IEnumVARIANT> enum_ptr;
	if SUCCEEDED(unk->QueryInterface(__uuidof(IEnumVARIANT), (void**)&enum_ptr)) {
		*disp = new DispEnumImpl(enum_ptr);
		(*disp)->AddRef();
		return true;
	}
	return false;
}

bool VariantDispGet(VARIANT *v, IDispatch **disp) {
	/*
	if ((v->vt & VT_ARRAY) != 0) {
		*disp = new DispArrayImpl(*v);
		(*disp)->AddRef();
		return true;
	}
	*/
	if ((v->vt & VT_TYPEMASK) == VT_DISPATCH) {
        *disp = ((v->vt & VT_BYREF) != 0) ? *v->ppdispVal : v->pdispVal;
        if (*disp) (*disp)->AddRef();
        return true;
    }
    if ((v->vt & VT_TYPEMASK) == VT_UNKNOWN) {
		return UnknownDispGet(((v->vt & VT_BYREF) != 0) ? *v->ppunkVal : v->punkVal, disp);
    }
    return false;
}

//-------------------------------------------------------------------------------------------------------
// DispArrayImpl implemetation

HRESULT STDMETHODCALLTYPE DispArrayImpl::GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) {
	if (cNames != 1 || !rgszNames[0]) return DISP_E_UNKNOWNNAME;
	LPOLESTR name = rgszNames[0];
	if (wcscmp(name, L"length") == 0) *rgDispId = 1;
	else return DISP_E_UNKNOWNNAME;
	return S_OK;
}

HRESULT STDMETHODCALLTYPE DispArrayImpl::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) {
	HRESULT hrcode = S_OK;
	UINT argcnt = pDispParams->cArgs;
	VARIANT *args = pDispParams->rgvarg;

	if ((var.vt & VT_ARRAY) == 0) return E_NOTIMPL;
	SAFEARRAY *arr = ((var.vt & VT_BYREF) != 0) ? *var.pparray : var.parray;

	switch (dispIdMember) {
	case 1: {
		if (pVarResult) {
			pVarResult->vt = VT_INT;
			pVarResult->intVal = (INT)(arr ? arr->rgsabound[0].cElements : 0);
		}
		return hrcode; }
	}
	return E_NOTIMPL;
}

//-------------------------------------------------------------------------------------------------------
// DispEnumImpl implemetation

HRESULT STDMETHODCALLTYPE DispEnumImpl::GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) {
    if (cNames != 1 || !rgszNames[0]) return DISP_E_UNKNOWNNAME;
    LPOLESTR name = rgszNames[0];
    if (wcscmp(name, L"Next") == 0) *rgDispId = 1;
    else if (wcscmp(name, L"Skip") == 0) *rgDispId = 2;
    else if (wcscmp(name, L"Reset") == 0) *rgDispId = 3;
    else if (wcscmp(name, L"Clone") == 0) *rgDispId = 4;
    else return DISP_E_UNKNOWNNAME;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE DispEnumImpl::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) {
    HRESULT hrcode = S_OK;
    UINT argcnt = pDispParams->cArgs;
    VARIANT *args = pDispParams->rgvarg;
    switch (dispIdMember) {
    case 1: {
        CComVariant arr;
        ULONG fetched, celt = (argcnt > 0) ? Variant2Int(args[argcnt - 1], (ULONG)1) : 1;
        if (!pVarResult || celt == 0) hrcode = E_INVALIDARG;
        if SUCCEEDED(hrcode) hrcode = arr.ArrayCreate(VT_VARIANT, celt);
        if SUCCEEDED(hrcode) hrcode = ptr->Next(celt, arr.ArrayGet<VARIANT>(0), &fetched);
        if SUCCEEDED(hrcode) {
            if (fetched == 0) pVarResult->vt = VT_EMPTY;
            else if (fetched == 1) {
                VARIANT *v = arr.ArrayGet<VARIANT>(0);
                *pVarResult = *v;
                v->vt = VT_EMPTY;
            }
            else {
                if (fetched < celt) hrcode = arr.ArrayResize(fetched);
                if SUCCEEDED(hrcode) arr.Detach(pVarResult);
            }
        }
        return hrcode; }
    case 2: {
        if (pVarResult) pVarResult->vt = VT_EMPTY;
        ULONG celt = (argcnt > 0) ? Variant2Int(args[argcnt - 1], (ULONG)1) : 1;
        return ptr->Skip(celt); 
        }
    case 3: {
        if (pVarResult) pVarResult->vt = VT_EMPTY;
        return ptr->Reset(); 
        }
    case 4: {
        if (!pVarResult) hrcode = E_INVALIDARG;
        std::auto_ptr<DispEnumImpl> disp;
        if SUCCEEDED(hrcode) hrcode = ptr->Clone(&disp->ptr);
        if SUCCEEDED(hrcode) {
            pVarResult->vt = VT_DISPATCH;
            pVarResult->pdispVal = disp.release();
            pVarResult->pdispVal->AddRef();
        }
        return hrcode; }
    }
    return E_NOTIMPL;
}

//-------------------------------------------------------------------------------------------------------
// DispObjectImpl implemetation

HRESULT STDMETHODCALLTYPE DispObjectImpl::GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) {
	if (cNames != 1 || !rgszNames[0]) return DISP_E_UNKNOWNNAME;
	std::wstring name(rgszNames[0]);
	name_ptr &ptr = names[name];
	if (!ptr) {
		ptr.reset(new name_t(dispid_next++, name));
		index.insert(index_t::value_type(ptr->dispid, ptr));
	}
	*rgDispId = ptr->dispid;
	return S_OK;
}

HRESULT STDMETHODCALLTYPE DispObjectImpl::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) {
	Isolate *isolate = Isolate::GetCurrent();
	Local<Object> self = obj.Get(isolate);
	Local<Value> name, val, ret;

	// Prepare name by member id
	if (dispIdMember != DISPID_VALUE) {
		index_t::const_iterator p = index.find(dispIdMember);
		if (p == index.end()) return DISP_E_MEMBERNOTFOUND;
		name_t &info = *p->second;
		name = String::NewFromTwoByte(isolate, (uint16_t*)info.name.c_str());
	}

	// Set property value
	if ((wFlags & DISPATCH_PROPERTYPUT) != 0) {
		UINT argcnt = pDispParams->cArgs;
		VARIANT *key = (argcnt > 1) ? &pDispParams->rgvarg[--argcnt] : nullptr;
		if (argcnt > 0) val = Variant2Value(isolate, pDispParams->rgvarg[--argcnt], true);
		else val = Undefined(isolate);
		bool rcode;

		// Set simple object property value
		if (!key) {
			if (name.IsEmpty()) return DISP_E_MEMBERNOTFOUND;
			rcode = self->Set(name, val);
		}

		// Set object/array item value
		else {
			Local<Object> target;
			if (name.IsEmpty()) target = self;
			else {
				Local<Value> obj = self->Get(name);
				if (!obj.IsEmpty()) target = Local<Object>::Cast(obj);
				if (target.IsEmpty()) return DISP_E_BADCALLEE;
			}

			LONG index = Variant2Int<LONG>(*key, -1);
			if (index >= 0) rcode = target->Set((uint32_t)index, val);
			else rcode = target->Set(Variant2Value(isolate, *key, false), val);
		}

		// Store result
		if (pVarResult) {
			pVarResult->vt = VT_BOOL;
			pVarResult->boolVal = rcode ? VARIANT_TRUE : VARIANT_FALSE;
		}
		return S_OK;
	}

	// Prepare property item
	if (name.IsEmpty()) val = self;
	else val = self->Get(name);

	// Call property as method
	if ((wFlags & DISPATCH_METHOD) != 0) {
		wFlags = 0;
		NodeArguments args(isolate, pDispParams, true);
		int argcnt = (int)args.items.size();
		Local<Value> *argptr = (argcnt > 0) ? &args.items[0] : nullptr;
		if (val->IsFunction()) {
			Local<Function> func = Local<Function>::Cast(val);
			if (func.IsEmpty()) return DISP_E_BADCALLEE;
			ret = func->Call(self, argcnt, argptr);
		}
		else if (val->IsObject()) {
			wFlags = DISPATCH_PROPERTYGET;
			//Local<Object> target = val->ToObject();
			//target->CallAsFunction(isolate->GetCurrentContext(), target, args.items.size(), &args.items[0]).ToLocal(&ret);
		}
		else {
			ret = val;
		}
	}

	// Get property value
	if ((wFlags & DISPATCH_PROPERTYGET) != 0) {
		if (pDispParams->cArgs == 1) {
			Local<Object> target;
			if (!val.IsEmpty()) target = Local<Object>::Cast(val);
			if (target.IsEmpty()) return DISP_E_BADCALLEE;
			VARIANT &key = pDispParams->rgvarg[0];
			LONG index = Variant2Int<LONG>(key, -1);
			if (index >= 0) ret = target->Get((uint32_t)index);
			else ret = target->Get(Variant2Value(isolate, key, false));
		}
		else {
			ret = val;
		}
	}

	// Store result
	if (pVarResult) {
		Value2Variant(isolate, ret, *pVarResult, VT_NULL);
	}
	return S_OK;
}

//-------------------------------------------------------------------------------------------------------
