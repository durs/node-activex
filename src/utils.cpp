//-------------------------------------------------------------------------------------------------------
// Project: NodeActiveX
// Author: Yuri Dursin
// Description:  Common utilities implementations
//-------------------------------------------------------------------------------------------------------

#include "stdafx.h"
#include "disp.h"

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
			len = FormatMessageW(FORMAT_MESSAGE_FROM_SYSTEM, 0, hrcode, MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), (LPOLESTR)bufptr, buflen, 0);
			if (len == 0) len = swprintf_s((LPOLESTR)bufptr, buflen, L"Error 0x%08X", hrcode);
		}
		buflen -= len;
		bufptr += len;
	}
	bufptr[0] = 0;
	return String::NewFromTwoByte(isolate, buf);
}

//-------------------------------------------------------------------------------------------------------

Local<Value> Variant2Value(Isolate *isolate, const VARIANT &v) {
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
	case VT_DISPATCH:
		return String::NewFromUtf8(isolate, "[Dispatch]");
	case VT_BSTR:
		return String::NewFromTwoByte(isolate, (uint16_t*)(by_ref ? *v.pbstrVal : v.bstrVal));
	case VT_VARIANT: 
		if (v.pvarVal) return Variant2Value(isolate, *v.pvarVal);
	}
	return Undefined(isolate);
}

void Value2Variant(Handle<Value> &val, VARIANT &var) {
	if (val.IsEmpty() || val->IsUndefined()) {
		var.vt = VT_EMPTY;
	}
	else if (val->IsNull()) {
		var.vt = VT_NULL;
	}
	else if (val->IsUint32()) {
		var.vt = VT_UI4;
		var.ulVal = val->Uint32Value();
	}
	else if (val->IsInt32()) {
		var.vt = VT_I4;
		var.lVal = val->Int32Value();
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
	else if (val->IsObject()) {
		var.vt = VT_DISPATCH;
		var.pdispVal = new DispObjectImpl(val->ToObject());
		var.pdispVal->AddRef();
	}
	else {
		String::Value str(val);
		var.vt = VT_BSTR;
		var.bstrVal = (str.length() > 0) ? SysAllocString((LPOLESTR)*str) : 0;
	}
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
		if (argcnt > 0) val = Variant2Value(isolate, pDispParams->rgvarg[--argcnt]);
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

			LONG index = Variant2nt<LONG>(*key, -1);
			if (index >= 0) rcode = target->Set((uint32_t)index, val);
			else rcode = target->Set(Variant2Value(isolate, *key), val);
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
		NodeArguments args(isolate, pDispParams);
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
			LONG index = Variant2nt<LONG>(key, -1);
			if (index >= 0) ret = target->Get((uint32_t)index);
			else ret = target->Get(Variant2Value(isolate, key));
		}
		else {
			ret = val;
		}
	}

	// Store result
	if (pVarResult) {
		Value2Variant(ret, *pVarResult);
	}
	return S_OK;
}

//-------------------------------------------------------------------------------------------------------
