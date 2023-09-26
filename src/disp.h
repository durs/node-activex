//-------------------------------------------------------------------------------------------------------
// Project: NodeActiveX
// Author: Yuri Dursin
// Description: DispObject class declarations. This class incapsulates COM IDispatch interface to Node JS Object
//-------------------------------------------------------------------------------------------------------

#pragma once

#include "utils.h"

enum options_t { 
    option_none = 0, 
    option_async = 0x0001, 
    option_type = 0x0002,
	option_activate = 0x0004,
	option_prepared = 0x0100,
    option_owned = 0x0200,
	option_property = 0x0400,
	option_function_simple = 0x0800,
	option_mask = 0x00FF,
	option_auto = (option_async | option_type)
};

inline bool TypeInfoGetName(ITypeInfo *info, DISPID dispid, BSTR *name) {
	HRESULT hrcode = info->GetDocumentation(dispid, name, NULL, NULL, NULL);
	if SUCCEEDED(hrcode) return true;
	UINT cnt_ret;
	return info->GetNames(dispid, name, 1, &cnt_ret) == S_OK && cnt_ret > 0;
}

template<typename T>
bool TypeInfoPrepareFunc(ITypeInfo *info, UINT n, T process) {
	FUNCDESC *desc;
	if (info->GetFuncDesc(n, &desc) != S_OK) return false;
	process(info, desc, nullptr);
	info->ReleaseFuncDesc(desc);
	return true;
}

template<typename T>
bool TypeInfoPrepareVar(ITypeInfo *info, UINT n, T process) {
	VARDESC *desc;
	if (info->GetVarDesc(n, &desc) != S_OK) return false;
	process(info, nullptr, desc);
	info->ReleaseVarDesc(desc);
	return true;
}

template<typename T>
void TypeInfoPrepare(ITypeInfo *info, int mode, T process) {
	UINT cFuncs = 0, cVars = 0;
	TYPEATTR *pattr = NULL;
	if (info->GetTypeAttr(&pattr) == S_OK) {
		cFuncs = pattr->cFuncs;
		cVars = pattr->cVars;
		info->ReleaseTypeAttr(pattr);
	}
	if ((mode & 1) != 0) {
		for (UINT n = 0; n < cFuncs; n++) {
			TypeInfoPrepareFunc<T>(info, n, process);
		}
	}
	if ((mode & 2) != 0) {
		for (UINT n = 0; n < cVars; n++) {
			TypeInfoPrepareVar<T>(info, n, process);
		}
	}
}

template<typename T>
bool TypeLibEnumerate(ITypeLib *typelib, int mode, T process) {
	UINT i, cnt = typelib ? typelib->GetTypeInfoCount() : 0;
	for (i = 0; i < cnt; i++) {
		CComPtr<ITypeInfo> info;
		if (typelib->GetTypeInfo(i, &info) != S_OK) continue;
		TypeInfoPrepare<T>(info, mode, process);
	}
	return cnt > 0;
}

class DispInfo;

template<typename T>
bool TypeInfoEnumerate(DispInfo* dispInfo, int mode, T process) {
	IDispatch* disp = dispInfo->ptr;
	UINT i, cnt;
	CComPtr<ITypeLib> prevtypelib;
	if (!disp || FAILED(disp->GetTypeInfoCount(&cnt))) cnt = 0;
	else for (i = 0; i < cnt; i++) {
		CComPtr<ITypeInfo> info;
		if (disp->GetTypeInfo(i, 0, &info) != S_OK) continue;

		// Query type library
		UINT typeindex;
		CComPtr<ITypeLib> typelib;
		if (info->GetContainingTypeLib(&typelib, &typeindex) == S_OK) {

			// Check if typelib is managed
			CComPtr<ITypeLib2> typeLib2;

			if (SUCCEEDED(typelib->QueryInterface(IID_ITypeLib2, (void**)&typeLib2)) )
			{
				// {90883F05-3D28-11D2-8F17-00A0C9A6186D}
				const GUID GUID_ExportedFromComPlus = { 0x90883F05, 0x3D28, 0x11D2, { 0x8F, 0x17, 0x00, 0xA0, 0xC9, 0xA6, 0x18, 0x6D } };

				CComVariant cv;
				if (SUCCEEDED(typeLib2->GetCustData(GUID_ExportedFromComPlus, &cv)))
				{
					dispInfo->bManaged = true;
				}
			}

			// Enumerate all types in library types
			// May be very slow! need a special method
			/*
			if (typelib != prevtypelib) {
				TypeLibEnumerate<T>(typelib, mode, process);
				prevtypelib.Attach(typelib.Detach());
			}
			*/

			CComPtr<ITypeInfo> tinfo;
			if (typelib->GetTypeInfo(typeindex, &tinfo) == S_OK) {
				TypeInfoPrepare<T>(tinfo, mode, process);
			}
		}

		// Process types
		else {
			TypeInfoPrepare<T>(info, mode, process);
		}
	}
	return cnt > 0;
}

class DispInfo {
public:
	std::weak_ptr<DispInfo> parent;
	CComPtr<IDispatch> ptr;
    std::wstring name;
	int options;
	bool bManaged;

	struct type_t { 
		DISPID dispid; 
		int kind; 
		int argcnt_get; 
		inline type_t(DISPID dispid_, int kind_) : dispid(dispid_), kind(kind_), argcnt_get(0) {}
		inline bool is_property() const { return ((kind & INVOKE_FUNC) == 0); }
		inline bool is_property_simple() const { return (((kind & (INVOKE_PROPERTYGET | INVOKE_FUNC))) == INVOKE_PROPERTYGET) && (argcnt_get == 0); }
		inline bool is_function_simple() const { return (((kind & (INVOKE_PROPERTYGET | INVOKE_FUNC))) == INVOKE_FUNC) && (argcnt_get == 0); }
		inline bool is_property_advanced() const { return kind == (INVOKE_PROPERTYGET | INVOKE_PROPERTYPUT) && (argcnt_get == 1); }
	};
	typedef std::shared_ptr<type_t> type_ptr;
	typedef std::map<DISPID, type_ptr> types_by_dispid_t;
	types_by_dispid_t types_by_dispid;

    inline DispInfo(IDispatch *disp, const std::wstring &nm, int opt, std::shared_ptr<DispInfo> *parnt = nullptr)
        : ptr(disp), options(opt), name(nm), bManaged(false)
    { 
        if (parnt) parent = *parnt;
        if ((options & option_type) != 0)
            Prepare(disp);
    }

    void Prepare(IDispatch *disp) {
        Enumerate(1, [this](ITypeInfo *info, FUNCDESC *func, VARDESC *var) {
			type_ptr &ptr = this->types_by_dispid[func->memid];
			if (!ptr) ptr.reset(new type_t(func->memid, func->invkind));
			else ptr->kind |= func->invkind;
			if ((func->invkind & INVOKE_PROPERTYGET) != 0) {
				if (func->cParams > ptr->argcnt_get)
					ptr->argcnt_get = func->cParams;
			}
        });
        bool prepared = types_by_dispid.size() > 3; // QueryInterface, AddRef, Release
        if (prepared) options |= option_prepared;
	}

    template<typename T>
    bool Enumerate(int mode, T process) {
		return TypeInfoEnumerate(this, mode, process);
    }

	inline bool GetTypeInfo(const DISPID dispid, type_ptr &info) {
		if ((options & option_prepared) == 0) return false;
		types_by_dispid_t::const_iterator it = types_by_dispid.find(dispid);
		if (it == types_by_dispid.end()) return false;
		info = it->second;
		return true;
	}

	HRESULT FindProperty(LPOLESTR name, DISPID *dispid) {
		return DispFind(ptr, name, dispid);
	}

	HRESULT GetProperty(DISPID dispid, LONG argcnt, VARIANT *args, VARIANT *value, EXCEPINFO *except = 0) {
		HRESULT hrcode = DispInvoke(ptr, dispid, argcnt, args, value, DISPATCH_PROPERTYGET, except);
		return hrcode;
	}

	HRESULT GetProperty(DISPID dispid, LONG index, VARIANT *value, EXCEPINFO *except = 0) {
		CComVariant arg(index);
		LONG argcnt = (index >= 0) ? 1 : 0;
		return DispInvoke(ptr, dispid, argcnt, &arg, value, DISPATCH_PROPERTYGET, except);
	}

	HRESULT SetProperty(DISPID dispid, LONG argcnt, VARIANT *args, VARIANT *value, EXCEPINFO *except = 0) {
		HRESULT hrcode = DispInvoke(ptr, dispid, argcnt, args, value, DISPATCH_PROPERTYPUT, except);
		if FAILED(hrcode) value->vt = VT_EMPTY;
		return hrcode;
	}

    HRESULT ExecuteMethod(DISPID dispid, LONG argcnt, VARIANT *args, VARIANT *value, EXCEPINFO *except = 0) {
        HRESULT hrcode = DispInvoke(ptr, dispid, argcnt, args, value, DISPATCH_METHOD, except);
        return hrcode;
    }
};

typedef std::shared_ptr<DispInfo> DispInfoPtr;

class DispObject: public NodeObject
{
public:
	DispObject(const DispInfoPtr &ptr, const std::wstring &name, DISPID id = DISPID_UNKNOWN, LONG indx = -1, int opt = 0);
	~DispObject();

	static Persistent<ObjectTemplate> inst_template;
	static Persistent<FunctionTemplate> clazz_template;
    static NodeMethods clazz_methods;

	static void NodeInit(const Local<Object> &target, Isolate* isolate, Local<Context> &ctx);
	static bool HasInstance(Isolate *isolate, const Local<Value> &obj) {
		Local<FunctionTemplate> clazz = clazz_template.Get(isolate);
		return !clazz.IsEmpty() && clazz->HasInstance(obj);
	}
    static DispObject *GetPtr(Isolate *isolate, const Local<Object> &obj) {
        Local<FunctionTemplate> clazz = clazz_template.Get(isolate);
        if (clazz.IsEmpty() || !clazz->HasInstance(obj)) return nullptr;
        return Unwrap<DispObject>(obj);
    }
    static IDispatch *GetDispPtr(Isolate *isolate, const Local<Object> &obj) {
        DispObject *self = GetPtr(isolate, obj);
        return (self && self->disp) ? self->disp->ptr : nullptr;
    }
    static bool GetValueOf(Isolate *isolate, const Local<Object> &obj, VARIANT &value) {
		DispObject *self = GetPtr(isolate, obj);
		return self && SUCCEEDED(self->valueOf(isolate, value, false));
	}
	static Local<Object> NodeCreate(Isolate *isolate, IDispatch *disp, const std::wstring &name, int opt) {
		Local<Object> parent;
		DispInfoPtr ptr(new DispInfo(disp, name, opt));
		return DispObject::NodeCreate(isolate, parent, ptr, name);
	}

private:
	static Local<Object> NodeCreate(Isolate *isolate, const Local<Object> &parent, const DispInfoPtr &ptr, const std::wstring &name, DISPID id = DISPID_UNKNOWN, LONG indx = -1, int opt = 0);

	static void NodeCreate(const FunctionCallbackInfo<Value> &args);
	static void NodeValueOf(const FunctionCallbackInfo<Value> &args);
	static void NodeToString(const FunctionCallbackInfo<Value> &args);
	static void NodeRelease(const FunctionCallbackInfo<Value> &args);
	static void NodeCast(const FunctionCallbackInfo<Value> &args);
    static void NodeGet(Local<Name> name, const PropertyCallbackInfo<Value> &args);
	static void NodeSet(Local<Name> name, Local<Value> value, const PropertyCallbackInfo<Value> &args);
	static void NodeGetByIndex(uint32_t index, const PropertyCallbackInfo<Value> &args);
	static void NodeSetByIndex(uint32_t index, Local<Value> value, const PropertyCallbackInfo<Value> &args);
	static void NodeCall(const FunctionCallbackInfo<Value> &args);

    static void NodeConnectionPoints(const FunctionCallbackInfo<Value> &args);
    static void PeakAndDispatchMessages(const FunctionCallbackInfo<Value> &args);

protected:
	bool release();
	bool get(LPOLESTR tag, LONG index, const PropertyCallbackInfo<Value> &args);
	bool set(LPOLESTR tag, LONG index, const Local<Value> &value, const PropertyCallbackInfo<Value> &args);
	void call(Isolate *isolate, const FunctionCallbackInfo<Value> &args);

	HRESULT valueOf(Isolate *isolate, VARIANT &value, bool simple);
	HRESULT valueOf(Isolate *isolate, const Local<Object> &self, Local<Value> &value);
	void toString(const FunctionCallbackInfo<Value> &args);
    Local<Value> getIdentity(Isolate *isolate);

private:
	int options;
	inline bool is_null() { return !disp; }
	inline bool is_prepared() { return (options & option_prepared) != 0; }
	inline bool is_object() { return dispid == DISPID_VALUE /*&& index < 0*/; }
	inline bool is_owned() { return (options & option_owned) != 0; }

	Persistent<Value> items, methods, vars;
	void initTypeInfo(Isolate *isolate);

	DispInfoPtr disp;
	std::wstring name;
	DISPID dispid;
	LONG index;

	HRESULT prepare();
};


class VariantObject : public NodeObject
{
public:
	VariantObject() {};
	VariantObject(const VARIANT &v) : value(v) {};
	VariantObject(const FunctionCallbackInfo<Value> &args);

	static Persistent<ObjectTemplate> inst_template;
	static Persistent<FunctionTemplate> clazz_template;
    static NodeMethods clazz_methods;

	static void NodeInit(const Local<Object> &target, Isolate* isolate, Local<Context> &ctx);
	static bool HasInstance(Isolate *isolate, const Local<Value> &obj) {
		Local<FunctionTemplate> clazz = clazz_template.Get(isolate);
		return !clazz.IsEmpty() && clazz->HasInstance(obj);
	}
	static VariantObject *GetInstanceOf(Isolate *isolate, const Local<Object> &obj) {
		Local<FunctionTemplate> clazz = clazz_template.Get(isolate);
		if (clazz.IsEmpty() || !clazz->HasInstance(obj)) return false;
		return Unwrap<VariantObject>(obj);
	}
	static bool GetValueOf(Isolate *isolate, const Local<Object> &obj, VARIANT &value) {
		Local<FunctionTemplate> clazz = clazz_template.Get(isolate);
		if (clazz.IsEmpty() || !clazz->HasInstance(obj)) return false;
		VariantObject *self = Unwrap<VariantObject>(obj);
		return self && SUCCEEDED(self->value.CopyTo(&value));
	}

	static Local<Object> NodeCreateInstance(const FunctionCallbackInfo<Value> &args);
	static void NodeCreate(const FunctionCallbackInfo<Value> &args);
	static Local<Object> NodeCreate(Isolate* isolate, const VARIANT& var);

	static void NodeClear(const FunctionCallbackInfo<Value> &args);
	static void NodeAssign(const FunctionCallbackInfo<Value> &args);
	static void NodeCast(const FunctionCallbackInfo<Value> &args);
	static void NodeValueOf(const FunctionCallbackInfo<Value> &args);
	static void NodeToString(const FunctionCallbackInfo<Value> &args);
	static void NodeGet(Local<Name> name, const PropertyCallbackInfo<Value> &args);
	static void NodeSet(Local<Name> name, Local<Value> value, const PropertyCallbackInfo<Value> &args);
	static void NodeGetByIndex(uint32_t index, const PropertyCallbackInfo<Value> &args);
	static void NodeSetByIndex(uint32_t index, Local<Value> value, const PropertyCallbackInfo<Value> &args);



private:
	CComVariant value, pvalue;
	bool assign(Isolate *isolate, Local<Value> &val, Local<Value> &type);
};

class ConnectionPointObject : public NodeObject
{
public:
    ConnectionPointObject(IConnectionPoint *p, IDispatch* d);
    ConnectionPointObject(const FunctionCallbackInfo<Value> &args) {}
    static Persistent<ObjectTemplate> inst_template;
    static Persistent<FunctionTemplate> clazz_template;
    static Local<Object> NodeCreateInstance(Isolate *isolate, IConnectionPoint *p, IDispatch* d);
    static void NodeInit(const Local<Object> &target, Isolate* isolate, Local<Context> &ctx);
    static void NodeCreate(const FunctionCallbackInfo<Value> &args);
    static void NodeAdvise(const FunctionCallbackInfo<Value> &args);
	static void NodeUnadvise(const FunctionCallbackInfo<Value> &args);
	static void NodeConnectionPointMethods(const FunctionCallbackInfo<Value> &args);

private:
    bool InitIndex();

    CComPtr<IConnectionPoint> ptr;
    CComPtr<IDispatch> disp;
    DispObjectImpl::index_t index;

	std::unordered_set<DWORD> cookies;

};
