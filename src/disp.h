//-------------------------------------------------------------------------------------------------------
// Project: NodeActiveX
// Author: Yuri Dursin
// Description: DispObject class declarations. This class incapsulates COM IDispatch interface to Node JS Object
//-------------------------------------------------------------------------------------------------------

#pragma once

#include "utils.h"

enum options_t { 
    option_none = 0, 
    option_async = 0x01, 
    option_type = 0x02,
	option_activate = 0x04,
	option_prepared = 0x10,
    option_owned = 0x20,
	option_auto = (option_async | option_type),
	option_mask = 0x0F
};

class DispInfo {
public:
	std::weak_ptr<DispInfo> parent;
	CComPtr<IDispatch> ptr;
    std::wstring name;
	int options;

    struct func_t { DISPID dispid; int kind; };
	typedef std::shared_ptr<func_t> func_ptr;
	typedef std::map<DISPID, func_ptr> func_by_dispid_t;
	func_by_dispid_t funcs_by_dispid;

    inline DispInfo(IDispatch *disp, const std::wstring &nm, int opt, std::shared_ptr<DispInfo> *parnt = nullptr)
        : ptr(disp), options(opt), name(nm)
    { 
        if (parnt) parent = *parnt;
        if ((options & option_type) != 0)
            Prepare(disp);
    }

    void Prepare(IDispatch *disp) {
        Enumerate([this](ITypeInfo *info, FUNCDESC *desc) {
			func_ptr &ptr = this->funcs_by_dispid[desc->memid];
			if (!ptr) {
				ptr.reset(new func_t);
				ptr->dispid = desc->memid;
				ptr->kind = desc->invkind;
			}
			else {
				ptr->kind |= desc->invkind;
			}
        });
        bool prepared = funcs_by_dispid.size() > 3; // QueryInterface, AddRef, Release
        if (prepared) options |= option_prepared;
	}

    template<typename T>
    bool Enumerate(T process) {
        UINT i, cnt;
        if (!ptr || FAILED(ptr->GetTypeInfoCount(&cnt))) cnt = 0;
        else for (i = 0; i < cnt; i++) {
            CComPtr<ITypeInfo> info;
            if (ptr->GetTypeInfo(i, 0, &info) != S_OK) continue;
            PrepareType<T>(info, process);
        }
        return cnt > 0;
    }

    template<typename T>
    bool PrepareType(ITypeInfo *info, T process) {
		UINT n = 0;
		while (PrepareFunc<T>(info, n, process)) n++;
		/*
		VARDESC *vdesc;
		if (info->GetVarDesc(dispid - 1, &vdesc) == S_OK) {
			info->ReleaseVarDesc(vdesc);
		}
		*/
		return n > 0;
	}

    template<typename T>
	bool PrepareFunc(ITypeInfo *info, UINT n, T process) {
		FUNCDESC *desc;
		if (info->GetFuncDesc(n, &desc) != S_OK) return false;
        process(info, desc);
		info->ReleaseFuncDesc(desc);
		return true;
	}
    inline bool GetItemName(ITypeInfo *info, DISPID dispid, BSTR *name) {
        UINT cnt_ret;
        return info->GetNames(dispid, name, 1, &cnt_ret) == S_OK && cnt_ret > 0;
    }

	inline bool IsProperty(const DISPID dispid) {
		if ((options & option_prepared) == 0) return false;
		func_by_dispid_t::const_iterator it = funcs_by_dispid.find(dispid);
		if (it == funcs_by_dispid.end()) return false;
		return (it->second->kind & (INVOKE_PROPERTYGET | INVOKE_FUNC)) == INVOKE_PROPERTYGET;
	}

	HRESULT FindProperty(LPOLESTR name, DISPID *dispid) {
		return DispFind(ptr, name, dispid);
	}

	HRESULT GetProperty(DISPID dispid, LONG index, VARIANT *value) {
		CComVariant arg(index);
		LONG argcnt = (index >= 0) ? 1 : 0;
		HRESULT hrcode = DispInvoke(ptr, dispid, argcnt, &arg, value, DISPATCH_PROPERTYGET);
        if FAILED(hrcode) {
            value->vt = VT_EMPTY;
            if (dispid == DISPID_VALUE) hrcode = S_OK;
        }
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
	DispObject(const DispInfoPtr &ptr, const std::wstring &name, DISPID id = DISPID_UNKNOWN, LONG indx = -1);
	~DispObject();

	static Persistent<ObjectTemplate> inst_template;
	static Persistent<FunctionTemplate> clazz_template;
	static void NodeInit(const Local<Object> &target);
	static bool HasInstance(Isolate *isolate, const Local<Value> &obj) {
		Local<FunctionTemplate> clazz = clazz_template.Get(isolate);
		return !clazz.IsEmpty() && clazz->HasInstance(obj);
	}
	static bool GetValueOf(Isolate *isolate, const Local<Object> &obj, VARIANT &value) {
		Local<FunctionTemplate> clazz = clazz_template.Get(isolate);
		if (clazz.IsEmpty() || !clazz->HasInstance(obj)) return false;
		DispObject *self = Unwrap<DispObject>(obj);
		return self && SUCCEEDED(self->valueOf(isolate, value));
	}
	static Local<Object> NodeCreate(Isolate *isolate, IDispatch *disp, const std::wstring &name, int opt) {
		Local<Object> parent;
		DispInfoPtr ptr(new DispInfo(disp, name, opt));
		return DispObject::NodeCreate(isolate, parent, ptr, name);
	}

private:
	static Local<Object> NodeCreate(Isolate *isolate, const Local<Object> &parent, const DispInfoPtr &ptr, const std::wstring &name, DISPID id = DISPID_UNKNOWN, LONG indx = -1);

	static void NodeCreate(const FunctionCallbackInfo<Value> &args);
	static void NodeValueOf(const FunctionCallbackInfo<Value> &args);
	static void NodeToString(const FunctionCallbackInfo<Value> &args);
	static void NodeRelease(const FunctionCallbackInfo<Value> &args);
	static void NodeGet(Local<String> name, const PropertyCallbackInfo<Value> &args);
	static void NodeSet(Local<String> name, Local<Value> value, const PropertyCallbackInfo<Value> &args);
	static void NodeGetByIndex(uint32_t index, const PropertyCallbackInfo<Value> &args);
	static void NodeSetByIndex(uint32_t index, Local<Value> value, const PropertyCallbackInfo<Value> &args);
	static void NodeCall(const FunctionCallbackInfo<Value> &args);

protected:
	bool release();
	bool get(LPOLESTR tag, LONG index, const PropertyCallbackInfo<Value> &args);
	bool set(LPOLESTR tag, LONG index, const Local<Value> &value, const PropertyCallbackInfo<Value> &args);
	void call(Isolate *isolate, const FunctionCallbackInfo<Value> &args);

	HRESULT valueOf(Isolate *isolate, VARIANT &value);
	HRESULT valueOf(Isolate *isolate, const Local<Object> &self, Local<Value> &value);
	void toString(const FunctionCallbackInfo<Value> &args);
    Local<Value> getIdentity(Isolate *isolate);
    Local<Value> getTypeInfo(Isolate *isolate);

private:
	int options;
	inline bool is_null() { return !disp; }
	inline bool is_prepared() { return (options & option_prepared) != 0; }
	inline bool is_object() { return dispid == DISPID_VALUE && index < 0; }
	inline bool is_owned() { return (options & option_owned) != 0; }

	DispInfoPtr disp;
	std::wstring name;
	DISPID dispid;
	LONG index;

	HRESULT prepare();
};
