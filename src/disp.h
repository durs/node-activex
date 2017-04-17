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
	std::shared_ptr<DispInfo> parent;
	CComPtr<IDispatch> ptr;
	bool prepared;

    struct func_t { MEMBERID mid; INVOKEKIND kind; };
	typedef std::shared_ptr<func_t> func_ptr;
	typedef std::map<std::wstring, func_ptr> func_map;
	func_map funcs;

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
		prepared = (funcs.size() > 3 /*QueryInterface, AddRef, Release */);
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
            f->mid = desc->memid;
            f->kind = desc->invkind;
			funcs[std::wstring(name)] = f;
		}
		info->ReleaseFuncDesc(desc);
		return true;
	}

	inline bool IsFunction(const std::wstring &name) {
		if (!prepared) return true;
		func_map::const_iterator it = funcs.find(name);
		if (it == funcs.end()) return false;
		return (it->second->kind & INVOKE_FUNC) != 0;
	}

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
	//DispObject(IDispatch *ptr, LPCOLESTR name, DISPID id = DISPID_UNKNOWN, LONG indx = -1);
	DispObject(const DispInfoPtr ptr, LPCOLESTR name, DISPID id = DISPID_UNKNOWN, LONG indx = -1);
	~DispObject();

	static void NodeInit(Handle<Object> target);
	static Local<Object> NodeCreate(Isolate *isolate, const DispInfoPtr &ptr, LPCOLESTR name, DISPID id = DISPID_UNKNOWN, LONG indx = -1);

private:
	static void NodeCreate(const FunctionCallbackInfo<Value> &args);
	static void NodeValueOf(const FunctionCallbackInfo<Value> &args);
	static void NodeToString(const FunctionCallbackInfo<Value> &args);
	static void NodeGet(Local<String> name, const PropertyCallbackInfo<Value> &args);
	static void NodeSet(Local<String> name, Local<Value> value, const PropertyCallbackInfo<Value> &args);
	static void NodeGetByIndex(uint32_t index, const PropertyCallbackInfo<Value> &args);
	static void NodeSetByIndex(uint32_t index, Local<Value> value, const PropertyCallbackInfo<Value> &args);
	static void NodeCall(const FunctionCallbackInfo<Value> &args);

protected:
	virtual Local<Value> find(Isolate *isolate, LPOLESTR name);
	virtual Local<Value> find(Isolate *isolate, uint32_t index);
	virtual Local<Value> valueOf(Isolate *isolate);
	virtual Local<Value> toString(Isolate *isolate);
	virtual bool set(Isolate *isolate, LPOLESTR name, Local<Value> value);
	virtual void call(Isolate *isolate, const FunctionCallbackInfo<Value> &args);

private:
    static Persistent<ObjectTemplate> inst_template;
    static Persistent<Function> constructor;

	enum State { None = 0, Prepared = 1, };
	int state;
	inline bool isPrepared() { return (state & Prepared) != 0; }
	//inline bool isOwned() { return (bool)owner_disp; }


	DispInfoPtr disp;
	//DispInfoPtr owner_disp;

	//CComPtr<IDispatch> disp;
	//CComPtr<IDispatch> owner_disp;
	std::wstring name;
	DISPID dispid;
	LONG index;

	HRESULT Prepare(VARIANT *value = 0);
};
