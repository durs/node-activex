//-------------------------------------------------------------------------------------------------------
// Project: NodeActiveX
// Author: Yuri Dursin
// Last Modification: 2017-04-16
// Description: DispObject class declarations. This class incapsulates COM IDispatch interface to Node JS Object
//-------------------------------------------------------------------------------------------------------

#pragma once

#include "utils.h"

class DispInfo {
public:
	std::weak_ptr<DispInfo> parent;
	CComPtr<IDispatch> ptr;
	bool prepared;

    struct func_t { DISPID dispid; INVOKEKIND kind; };
	typedef std::shared_ptr<func_t> func_ptr;
	typedef std::map<DISPID, func_ptr> func_by_dispid_t;
	//typedef std::map<std::wstring, func_ptr> func_by_name_t;
	func_by_dispid_t funcs_by_dispid;
	//func_by_name_t funcs_by_name;

	inline DispInfo() : prepared(false) {}
	inline DispInfo(IDispatch *disp) : prepared(false) { Prepare(disp); }

	bool Prepare(IDispatch *disp) {
		UINT i, cnt;
		if SUCCEEDED(disp->GetTypeInfoCount(&cnt)) {
			for (i = 0; i < cnt; i++) {
				CComPtr<ITypeInfo> info;
				if (disp->GetTypeInfo(i, 0, &info) != S_OK) continue;
				PrepareType(info);
			}
		}
		ptr = disp;
		prepared = (funcs_by_dispid.size() > 3 /*QueryInterface, AddRef, Release */);
		return prepared;;
	}

	bool PrepareType(ITypeInfo *info) {
		UINT n = 0;
		while (PrepareFunc(info, n)) n++;
		/*
		VARDESC *vdesc;
		if (info->GetVarDesc(dispid - 1, &vdesc) == S_OK) {
			info->ReleaseVarDesc(vdesc);
		}
		*/
		return n > 0;
	}

	bool PrepareFunc(ITypeInfo *info, UINT n) {
		FUNCDESC *desc;
		if (info->GetFuncDesc(n, &desc) != S_OK) return false;
		CComBSTR name;
		UINT cnt_ret = 1;
		if (info->GetNames(desc->memid, &name, 1, &cnt_ret) == S_OK && cnt_ret > 0 && (bool)name) {
			func_ptr f(new func_t);
            f->dispid = desc->memid;
            f->kind = desc->invkind;
			funcs_by_dispid[desc->memid] = f;
			//funcs_by_name[std::wstring(name)] = f;
		}
		info->ReleaseFuncDesc(desc);
		return true;
	}

	inline bool IsProperty(const DISPID dispid) {
		if (!prepared) return false;
		func_by_dispid_t::const_iterator it = funcs_by_dispid.find(dispid);
		if (it == funcs_by_dispid.end()) return false;
		return (it->second->kind & (INVOKE_PROPERTYGET | INVOKE_FUNC)) == INVOKE_PROPERTYGET;
	}

	/*
	inline bool IsProperty(const std::wstring &name) {
		if (!prepared) return false;
		func_by_name_t::const_iterator it = funcs_by_name.find(name);
		if (it == funcs_by_name.end()) return false;
		return (it->second->kind & (INVOKE_PROPERTYGET | INVOKE_FUNC)) == INVOKE_PROPERTYGET;
	}
	*/

	HRESULT FindProperty(LPOLESTR name, DISPID *dispid) {
		return DispFind(ptr, name, dispid);
	}

	HRESULT GetProperty(DISPID dispid, LONG index, VARIANT *value) {
		CComVariant arg(index);
		LONG argcnt = (index >= 0) ? 1 : 0;
		HRESULT hrcode = DispInvoke(ptr, dispid, argcnt, &arg, value, DISPATCH_PROPERTYGET);
		if FAILED(hrcode) value->vt = VT_EMPTY;
		return hrcode;
	}

	HRESULT SetProperty(DISPID dispid, LONG argcnt, VARIANT *args, VARIANT *value) {
		HRESULT hrcode = DispInvoke(ptr, dispid, argcnt, args, value, DISPATCH_PROPERTYPUT);
		if FAILED(hrcode) value->vt = VT_EMPTY;
		return hrcode;
	}

    HRESULT ExecuteMethod(DISPID dispid, LONG argcnt, VARIANT *args, VARIANT *value) {
        HRESULT hrcode = DispInvoke(ptr, dispid, argcnt, args, value, DISPATCH_METHOD);
        return hrcode;
    }
};

typedef std::shared_ptr<DispInfo> DispInfoPtr;


class DispObject: public ObjectWrap
{
public:
	DispObject(const DispInfoPtr &ptr, LPCOLESTR name, DISPID id = DISPID_UNKNOWN, LONG indx = -1);
	~DispObject();

	static void NodeInit(Handle<Object> target);

private:
	static Local<Object> NodeCreate(Isolate *isolate, const Local<Object> &parent, const DispInfoPtr &ptr, LPCOLESTR name, DISPID id = DISPID_UNKNOWN, LONG indx = -1);

	static void NodeCreate(const FunctionCallbackInfo<Value> &args);
	static void NodeValueOf(const FunctionCallbackInfo<Value> &args);
	static void NodeToString(const FunctionCallbackInfo<Value> &args);
	static void NodeGet(Local<String> name, const PropertyCallbackInfo<Value> &args);
	static void NodeSet(Local<String> name, Local<Value> value, const PropertyCallbackInfo<Value> &args);
	static void NodeGetByIndex(uint32_t index, const PropertyCallbackInfo<Value> &args);
	static void NodeSetByIndex(uint32_t index, Local<Value> value, const PropertyCallbackInfo<Value> &args);
	static void NodeCall(const FunctionCallbackInfo<Value> &args);

protected:
	bool get(LPOLESTR name, uint32_t index, const PropertyCallbackInfo<Value> &args);
	bool get(uint32_t index, const PropertyCallbackInfo<Value> &args);
	bool set(Isolate *isolate, LPOLESTR name, Local<Value> value);
	void call(Isolate *isolate, const FunctionCallbackInfo<Value> &args);

	HRESULT valueOf(Isolate *isolate, Local<Value> &value);
	void toString(const FunctionCallbackInfo<Value> &args);

private:
    static Persistent<ObjectTemplate> inst_template;
    static Persistent<Function> constructor;

	enum state_t { state_none = 0, state_prepared = 1, state_owned = 2 };
	int state;
	inline bool is_prepared() { return (state & state_prepared) != 0; }
	inline bool is_owned() { return (state & state_owned) != 0; }

	DispInfoPtr disp;
	std::wstring name;
	DISPID dispid;
	LONG index;

	HRESULT Prepare(VARIANT *value = 0);
};
