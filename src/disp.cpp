//-------------------------------------------------------------------------------------------------------
// Project: NodeActiveX
// Author: Yuri Dursin
// Description: DispObject class implementations
//-------------------------------------------------------------------------------------------------------

#include "stdafx.h"
#include "disp.h"

Persistent<ObjectTemplate> DispObject::inst_template;
Persistent<FunctionTemplate> DispObject::clazz_template;
NodeMethods DispObject::clazz_methods;

Persistent<ObjectTemplate> VariantObject::inst_template;
Persistent<FunctionTemplate> VariantObject::clazz_template;
NodeMethods VariantObject::clazz_methods;

Persistent<ObjectTemplate> ConnectionPointObject::inst_template;
Persistent<FunctionTemplate> ConnectionPointObject::clazz_template;

//-------------------------------------------------------------------------------------------------------
// DispObject implemetation

DispObject::DispObject(const DispInfoPtr &ptr, const std::wstring &nm, DISPID id, LONG indx, int opt)
	: disp(ptr), options((ptr->options & option_mask) | opt), name(nm), dispid(id), index(indx)
{	
	if (dispid == DISPID_UNKNOWN) {
		dispid = DISPID_VALUE;
        options |= option_prepared;
	}
	else options |= option_owned;
    NODE_DEBUG_FMT("DispObject '%S' constructor", name.c_str());
}

DispObject::~DispObject() {
	items.Reset();
	methods.Reset();
	vars.Reset();
	NODE_DEBUG_FMT("DispObject '%S' destructor", name.c_str());
}

HRESULT DispObject::prepare() {
	CComVariant value;
	HRESULT hrcode = disp ? disp->GetProperty(dispid, index, &value) : E_UNEXPECTED;

	// Init dispatch interface
	options |= option_prepared;
	CComPtr<IDispatch> ptr;
	if (VariantDispGet(&value, &ptr)) {
		disp.reset(new DispInfo(ptr, name, options, &disp));
		dispid = DISPID_VALUE;
	}
	else if ((value.vt & VT_ARRAY) != 0) {
		
	}
	return hrcode;
}

bool DispObject::release() {
    if (!disp) return false;
    NODE_DEBUG_FMT("DispObject '%S' release", name.c_str());
    disp.reset();            
	items.Reset();
	methods.Reset();
	vars.Reset();
	return true;
}

bool DispObject::get(LPOLESTR tag, LONG index, const PropertyCallbackInfo<Value> &args) {
	Isolate *isolate = args.GetIsolate();
	if (!is_prepared()) prepare();
    if (!disp) {
        isolate->ThrowException(DispErrorNull(isolate));
        return false;
    }

	// Search dispid
    HRESULT hrcode;
    DISPID propid;
	bool prop_by_key = false;
    bool this_prop = false;
    if (!tag || !*tag) {
        tag = (LPOLESTR)name.c_str();
        propid = dispid;
        this_prop = true;
    }
	else {
        hrcode = disp->FindProperty(tag, &propid);
        if (SUCCEEDED(hrcode) && propid == DISPID_UNKNOWN) hrcode = E_INVALIDARG;
        if FAILED(hrcode) {
			prop_by_key = (options & option_property) != 0;
			if (!prop_by_key) {
				//isolate->ThrowException(DispError(isolate, hrcode, L"DispPropertyFind", tag));
				args.GetReturnValue().SetUndefined();
				return false;
			}
			propid = dispid;
        }
    }

	// Check type info
	int opt = 0;
	bool is_property_simple = false;
	if (prop_by_key) {
		is_property_simple = true;
		opt |= option_property;
	}
	else {
		DispInfo::type_ptr disp_info;
		if (disp->GetTypeInfo(propid, disp_info)) {
			if (disp_info->is_function_simple()) opt |= option_function_simple;
			else {
				if (disp_info->is_property()) opt |= option_property;
				is_property_simple = disp_info->is_property_simple();
				
				if (disp->bManaged && tag && *tag && wcscmp(tag, L"ToString")==0)
				{
					// .NET ToString is reported as a property while we normally use it via .ToString() - i.e. as a method. 
					is_property_simple = false;
				}

			}
		}
		else if ( disp->bManaged && tag && *tag && wcscmp(tag, L"length") == 0) {
			DISPID lenprop;
			if ( SUCCEEDED(disp->FindProperty(L"length", &lenprop)) )
			{

				// If we have 'IReflect' and '.length' - assume it is .NET JS Array or JS Object
				is_property_simple = true;
			}
		}
		else if (disp->bManaged && tag && *tag && index>=0 ) {
			// jsarray[x]
			is_property_simple = true;
		}
	}

    // Return as property value
	if (is_property_simple) {
		CComException except;
		CComVariant value;
		VarArguments vargs;
		if (prop_by_key) vargs.items.push_back(CComVariant(tag));
		if (index >= 0) vargs.items.push_back(CComVariant(index));
		LONG argcnt = (LONG)vargs.items.size();
		VARIANT *pargs = (argcnt > 0) ? &vargs.items.front() : 0;
		hrcode = disp->GetProperty(propid, argcnt, pargs, &value, &except);
		if (FAILED(hrcode) && dispid != DISPID_VALUE){
			isolate->ThrowException(DispError(isolate, hrcode, L"DispPropertyGet", tag, &except));
			return false;
		}
		CComPtr<IDispatch> ptr;
		if (VariantDispGet(&value, &ptr)) {
			DispInfoPtr disp_result(new DispInfo(ptr, tag, options, &disp));
			Local<Object> result = DispObject::NodeCreate(isolate, args.This(), disp_result, tag, DISPID_UNKNOWN, -1, opt);
			args.GetReturnValue().Set(result);
		}
		else {
			args.GetReturnValue().Set(Variant2Value(isolate, value));
		}
	}

	// Return as dispatch object 
	else {
		Local<Object> result = DispObject::NodeCreate(isolate, args.This(), disp, tag, propid, index, opt);
		args.GetReturnValue().Set(result);
	}
	return true;
}

bool DispObject::set(LPOLESTR tag, LONG index, const Local<Value> &value, const PropertyCallbackInfo<Value> &args) {
	Isolate *isolate = args.GetIsolate();
	if (!is_prepared()) prepare();
    if (!disp) {
        isolate->ThrowException(DispErrorNull(isolate));
        return false;
    }
	
	// Search dispid
	HRESULT hrcode;
	DISPID propid;
    if (!tag || !*tag) {
		tag = (LPOLESTR)name.c_str();
		propid = dispid;
	}
	else {
		hrcode = disp->FindProperty(tag, &propid);
		if (SUCCEEDED(hrcode) && propid == DISPID_UNKNOWN) hrcode = E_INVALIDARG;
		if FAILED(hrcode) {
			isolate->ThrowException(DispError(isolate, hrcode, L"DispPropertyFind", tag));
			return false;
		}
	}

	// Set value using dispatch
	CComException except;
    CComVariant ret;
	VarArguments vargs(isolate, value);
	if (index >= 0) vargs.items.push_back(CComVariant(index));
	LONG argcnt = (LONG)vargs.items.size();
    VARIANT *pargs = (argcnt > 0) ? &vargs.items.front() : 0;
	hrcode = disp->SetProperty(propid, argcnt, pargs, &ret, &except);
	if FAILED(hrcode) {
		isolate->ThrowException(DispError(isolate, hrcode, L"DispPropertyPut", tag, &except));
        return false;
    }

	// Send result
	CComPtr<IDispatch> ptr;
	if (VariantDispGet(&ret, &ptr)) {
		std::wstring rtag;
		rtag.reserve(32);
		rtag += L"@";
		rtag += tag;
		DispInfoPtr disp_result(new DispInfo(ptr, tag, options, &disp));
		Local<Object> result = DispObject::NodeCreate(isolate, args.This(), disp_result, rtag);
		args.GetReturnValue().Set(result);
	}
	else {
		args.GetReturnValue().Set(Variant2Value(isolate, ret));
	}
    return true;
}

void DispObject::call(Isolate *isolate, const FunctionCallbackInfo<Value> &args) {
    if (!disp) {
        isolate->ThrowException(DispErrorNull(isolate));
        return;
    }

	CComException except;
	CComVariant ret;
	VarArguments vargs(isolate, args);
	LONG argcnt = (LONG)vargs.items.size();
	VARIANT *pargs = (argcnt > 0) ? &vargs.items.front() : 0;
	HRESULT hrcode;

    if (vargs.IsDefault()) {
        hrcode = valueOf(isolate, ret, true);
    }
    else if ((options & option_property) == 0) {
        hrcode = disp->ExecuteMethod(dispid, argcnt, pargs, &ret, &except);
    }
    else {
        DispInfo::type_ptr disp_info;
		disp->GetTypeInfo(dispid, disp_info);

		if(disp_info->is_property_advanced() && argcnt > 1) {
			hrcode = disp->SetProperty(dispid, argcnt, pargs, &ret, &except);
		}
		else {
			hrcode = disp->GetProperty(dispid, argcnt, pargs, &ret, &except);
		}
    }
    if FAILED(hrcode) {
        isolate->ThrowException(DispError(isolate, hrcode, L"DispInvoke", name.c_str(), &except));
        return;
    }

	// Prepare result
    Local<Value> result;
	CComPtr<IDispatch> ptr;
	if (VariantDispGet(&ret, &ptr)) {
        std::wstring tag;
        tag.reserve(32);
        tag += L"@";
        tag += name;
		DispInfoPtr disp_result(new DispInfo(ptr, tag, options, &disp));
		result = DispObject::NodeCreate(isolate, args.This(), disp_result, tag);
	}
	else {
		result = Variant2Value(isolate, ret, true);
	}
    args.GetReturnValue().Set(result);
}

HRESULT DispObject::valueOf(Isolate *isolate, VARIANT &value, bool simple) {
	if (!is_prepared()) prepare();
	HRESULT hrcode;
	if (!disp) hrcode = E_UNEXPECTED;

	// simple function without arguments
	else if ((options & option_function_simple) != 0) {
		hrcode = disp->ExecuteMethod(dispid, 0, 0, &value);
	}

	// property or array element
	else if (dispid != DISPID_VALUE || index >= 0 || simple) {
		hrcode = disp->GetProperty(dispid, index, &value);
	}

	// self dispatch object
	else /*if (is_object())*/ {
		value.vt = VT_DISPATCH;
		value.pdispVal = (IDispatch*)disp->ptr;
		if (value.pdispVal) value.pdispVal->AddRef();
		hrcode = S_OK;
	}
	return hrcode;
}

HRESULT DispObject::valueOf(Isolate *isolate, const Local<Object> &self, Local<Value> &value) {
	if (!is_prepared()) prepare();
	HRESULT hrcode;
	if (!disp) hrcode = E_UNEXPECTED;
	else {
		CComVariant val;

		// simple function without arguments

		if ((options & option_function_simple) != 0) {
			hrcode = disp->ExecuteMethod(dispid, 0, 0, &val);
		}

		// self value, property or array element
		else {
			hrcode = disp->GetProperty(dispid, index, &val);
			// Try to get some primitive value
			if FAILED(hrcode) {
				hrcode = disp->ExecuteMethod(dispid, 0, 0, &val);
			}
		}

		// convert result to v8 value
		if SUCCEEDED(hrcode) {
			value = Variant2Value(isolate, val);
		}

		// or return self as object
		else  {
			value = self;
			hrcode = S_OK;
		}
	}
	return hrcode;
}

void DispObject::toString(const FunctionCallbackInfo<Value> &args) {
	Isolate *isolate = args.GetIsolate();
	CComVariant val;
	HRESULT hrcode = valueOf(isolate, val, true);
	if FAILED(hrcode) {
		isolate->ThrowException(Win32Error(isolate, hrcode, L"DispToString"));
		return;
	}
	args.GetReturnValue().Set(Variant2String(isolate, val));
}

Local<Value> DispObject::getIdentity(Isolate *isolate) {
    //wchar_t buf[64];
    std::wstring id;
    id.reserve(128);
    id += name;
    DispInfoPtr ptr = disp;
    if (ptr && ptr->name == id)
        ptr = ptr->parent.lock();
    while (ptr) {
        id.insert(0, L".");
        id.insert(0, ptr->name);
        /*
        if (ptr->index >= 0) {
            swprintf_s(buf, L"[%ld]", index);
            id += buf;
        }
        */
        ptr = ptr->parent.lock();
    }
	return v8str(isolate, id.c_str());
}

void DispObject::initTypeInfo(Isolate *isolate) {
    if ((options & option_type) == 0 || !disp) {
        return;
    }
    uint32_t index = 0;
	Local<v8::Object> _items = v8::Array::New(isolate);
	Local<v8::Object> _methods = v8::Object::New(isolate);
	Local<v8::Object> _vars = v8::Object::New(isolate);
	disp->Enumerate(3, [isolate, &_items, &_vars, &_methods, &index](ITypeInfo *info, FUNCDESC *func, VARDESC *var) {
		Local<Context> ctx = isolate->GetCurrentContext();
        CComBSTR name;
		MEMBERID memid = func != nullptr ? func->memid : var->memid;
        TypeInfoGetName(info, memid, &name);
        Local<Object> item(Object::New(isolate));
		Local<String> vname;
		if (name) {
			vname = v8str(isolate, (BSTR)name);
			item->Set(ctx, v8str(isolate, "name"), vname);
		}
		item->Set(ctx, v8str(isolate, "dispid"), Int32::New(isolate, memid));
		if (func != nullptr) {
			item->Set(ctx, v8str(isolate, "invkind"), Int32::New(isolate, (int32_t)func->invkind));
			item->Set(ctx, v8str(isolate, "flags"), Int32::New(isolate, (int32_t)func->wFuncFlags));
			item->Set(ctx, v8str(isolate, "argcnt"), Int32::New(isolate, (int32_t)func->cParams));
			_methods->Set(ctx, vname, item);
		} else {
			item->Set(ctx, v8str(isolate, "varkind"), Int32::New(isolate, (int32_t)var->varkind));
			item->Set(ctx, v8str(isolate, "flags"), Int32::New(isolate, (int32_t)var->wVarFlags));
			if (var->varkind == VAR_CONST && var->lpvarValue != nullptr) {
				v8::Local<v8::Value> value = Variant2Value(isolate, *var->lpvarValue, false);
				item->Set(ctx, v8str(isolate, "value"), value);
			}
			_vars->Set(ctx, vname, item);
		}
		_items->Set(ctx, index++, item);
    });
	items.Reset(isolate, _items);
	methods.Reset(isolate, _methods);
	vars.Reset(isolate, _vars);
}

//-----------------------------------------------------------------------------------
// Static Node JS callbacks

void DispObject::NodeInit(const Local<Object> &target, Isolate* isolate, Local<Context> &ctx) {

    // Prepare constructor template
    Local<FunctionTemplate> clazz = FunctionTemplate::New(isolate, NodeCreate);
    clazz->SetClassName(v8str(isolate, "Dispatch"));

    clazz_methods.add(isolate, clazz, "toString", NodeToString);
    clazz_methods.add(isolate, clazz, "valueOf", NodeValueOf);

    Local<ObjectTemplate> &inst = clazz->InstanceTemplate();
    inst->SetInternalFieldCount(1);
    inst->SetHandler(NamedPropertyHandlerConfiguration(NodeGet, NodeSet));
    inst->SetHandler(IndexedPropertyHandlerConfiguration(NodeGetByIndex, NodeSetByIndex));
    inst->SetCallAsFunctionHandler(NodeCall);
	inst->SetNativeDataProperty(v8str(isolate, "__id"), NodeGet);
	inst->SetNativeDataProperty(v8str(isolate, "__value"), NodeGet);
	//inst->SetLazyDataProperty(v8str(isolate, "__type"), NodeGet, Local<Value>(), ReadOnly);
	//inst->SetLazyDataProperty(v8str(isolate, "__methods"), NodeGet, Local<Value>(), ReadOnly);
	//inst->SetLazyDataProperty(v8str(isolate, "__vars"), NodeGet, Local<Value>(), ReadOnly);
	inst->SetNativeDataProperty(v8str(isolate, "__type"), NodeGet);
	inst->SetNativeDataProperty(v8str(isolate, "__methods"), NodeGet);
	inst->SetNativeDataProperty(v8str(isolate, "__vars"), NodeGet);

    inst_template.Reset(isolate, inst);
	clazz_template.Reset(isolate, clazz);
    target->Set(ctx, v8str(isolate, "Object"), clazz->GetFunction(ctx).ToLocalChecked());
	target->Set(ctx, v8str(isolate, "cast"), FunctionTemplate::New(isolate, NodeCast)->GetFunction(ctx).ToLocalChecked());
	target->Set(ctx, v8str(isolate, "release"), FunctionTemplate::New(isolate, NodeRelease)->GetFunction(ctx).ToLocalChecked());

    target->Set(ctx, v8str(isolate, "getConnectionPoints"), FunctionTemplate::New(isolate, NodeConnectionPoints)->GetFunction(ctx).ToLocalChecked());
    target->Set(ctx, v8str(isolate, "peekAndDispatchMessages"), FunctionTemplate::New(isolate, PeakAndDispatchMessages)->GetFunction(ctx).ToLocalChecked());

    //Context::GetCurrent()->Global()->Set(v8str(isolate, "ActiveXObject"), t->GetFunction());
	NODE_DEBUG_MSG("DispObject initialized");
}

Local<Object> DispObject::NodeCreate(Isolate *isolate, const Local<Object> &parent, const DispInfoPtr &ptr, const std::wstring &name, DISPID id, LONG index, int opt) {
    Local<Object> self;
    if (!inst_template.IsEmpty()) {
		if (inst_template.Get(isolate)->NewInstance(isolate->GetCurrentContext()).ToLocal(&self)) {
			(new DispObject(ptr, name, id, index, opt))->Wrap(self);
			//Local<String> prop_id(v8str(isolate, "_identity"));
			//self->Set(prop_id, v8str(isolate, name.c_str()));
		}
	}
    return self;
}

void DispObject::NodeCreate(const FunctionCallbackInfo<Value> &args) {
    Isolate *isolate = args.GetIsolate();
	Local<Context> ctx = isolate->GetCurrentContext();
	bool isGetObject = false;
	bool isGetAccessibleObject = false;
    int argcnt = args.Length();
    if (argcnt < 1) {
        isolate->ThrowException(InvalidArgumentsError(isolate));
        return;
    }
    int options = (option_async | option_type);
    if (argcnt > 1) {
        Local<Value> val, argopt = args[1];
		bool isEmpty = argopt.IsEmpty();
		bool isObject = argopt->IsObject();
        if (!isEmpty && isObject) {
            auto opt = Local<Object>::Cast(argopt);
			if (opt->Get(ctx, v8str(isolate, "async")).ToLocal(&val)) {
				if (!v8val2bool(isolate, val, true)) options &= ~option_async;
            }
			if (opt->Get(ctx, v8str(isolate, "type")).ToLocal(&val)) {
				if (!v8val2bool(isolate, val, true)) options &= ~option_type;
            }
			if (opt->Get(ctx, v8str(isolate, "activate")).ToLocal(&val)) {
				if (v8val2bool(isolate, val, false)) options |= option_activate;
			}
			if (opt->Get(ctx, v8str(isolate, "getobject")).ToLocal(&val)) {
				if (v8val2bool(isolate, val, false)) isGetObject = true;
			}
			if (opt->Get(ctx, v8str(isolate, "getaccessibleobject")).ToLocal(&val)) {
				if (v8val2bool(isolate, val, false)) isGetAccessibleObject = true;
			}
		}
    }
   
    // Invoked as plain function
    if (!args.IsConstructCall()) {
		Local<FunctionTemplate> clazz = clazz_template.Get(isolate);
		if (clazz.IsEmpty()) {
			isolate->ThrowException(TypeError(isolate, "FunctionTemplateIsEmpty"));
			return;
		}
        const int argc = 1;
        Local<Value> argv[argc] = { args[0] };
		Local<Function> cons;
		Local<Context> context = isolate->GetCurrentContext();
		if (clazz->GetFunction(context).ToLocal(&cons)) {
			Local<Object> self;
			if (cons->NewInstance(context, argc, argv).ToLocal(&self)) {
				args.GetReturnValue().Set(self);
			}
		}
        return;
    }

	// Create dispatch object from ProgId
	HRESULT hrcode;
	std::wstring name;
	CComPtr<IDispatch> disp;
	if (args[0]->IsString()) {

		// Prepare arguments
        String::Value vname(isolate, args[0]);
		if (vname.length() <= 0) hrcode = E_INVALIDARG;
		else {
			name.assign((LPOLESTR)*vname, vname.length());

			CComPtr<IUnknown> unk;
			if (isGetObject)
			{
				hrcode = CoGetObject(name.c_str(), NULL, IID_IUnknown, (void**)&unk);
				if SUCCEEDED(hrcode) hrcode = unk->QueryInterface(&disp);
			} else {
				if (isGetAccessibleObject)
				{
					hrcode = GetAccessibleObject(name.c_str(), unk);
					if SUCCEEDED(hrcode) hrcode = unk->QueryInterface(&disp);
				} else {
					CLSID clsid;
					hrcode = CLSIDFromProgID(name.c_str(), &clsid);
					if SUCCEEDED(hrcode) {
						if ((options & option_activate) == 0) hrcode = E_FAIL;
						else {
							hrcode = GetActiveObject(clsid, NULL, &unk);
							if SUCCEEDED(hrcode) hrcode = unk->QueryInterface(&disp);
						}
						if FAILED(hrcode) {
							hrcode = disp.CoCreateInstance(clsid, nullptr, CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER);
						}
					}
				}
			}
		}
	}

	// Use supplied dispatch pointer
	else if (args[0]->IsUint8Array()) {
        size_t len = node::Buffer::Length(args[0]);
        void *data = node::Buffer::Data(args[0]);
        IDispatch *p = (len == sizeof(INT_PTR)) ? (IDispatch *) *(static_cast<INT_PTR*>(data)) : nullptr;
        if (!p) {
            isolate->ThrowException(InvalidArgumentsError(isolate));
            return;
        }
		disp.Attach(p);
		hrcode = S_OK;
	}

	// Create dispatch object from javascript object
	else if (args[0]->IsObject()) {
		name = L"#";
		disp = new DispObjectImpl(Local<Object>::Cast(args[0]));
		hrcode = S_OK;
	}

	// Other
	else {
		hrcode = E_INVALIDARG;
	}

	// Prepare result
	if FAILED(hrcode) {
		isolate->ThrowException(DispError(isolate, hrcode, L"CreateInstance", name.c_str()));
	}
	else {
		Local<Object> &self = args.This();
		DispInfoPtr ptr(new DispInfo(disp, name, options));
		(new DispObject(ptr, name))->Wrap(self);
		args.GetReturnValue().Set(self);
	}
}

void DispObject::NodeGet(Local<Name> name, const PropertyCallbackInfo<Value>& args) {
    Isolate *isolate = args.GetIsolate();
	DispObject *self = DispObject::Unwrap<DispObject>(args.This());
	if (!self) {
		isolate->ThrowException(DispErrorInvalid(isolate));
		return;
	}
	String::Value vname(isolate, name);
	LPOLESTR id = (vname.length() > 0) ? (LPOLESTR)*vname : L"valueOf";
    NODE_DEBUG_FMT2("DispObject '%S.%S' get", self->name.c_str(), id);
    if (_wcsicmp(id, L"__value") == 0) {
        Local<Value> result;
        HRESULT hrcode = self->valueOf(isolate, args.This(), result);
        if FAILED(hrcode) isolate->ThrowException(Win32Error(isolate, hrcode, L"DispValueOf"));
        else args.GetReturnValue().Set(result);
    }
    else if (_wcsicmp(id, L"__id") == 0) {
		args.GetReturnValue().Set(self->getIdentity(isolate));
	}
	else if (_wcsicmp(id, L"__type") == 0) {
		if (self->items.IsEmpty()) {
			self->initTypeInfo(isolate);
		}
		Local<Value> result = self->items.Get(isolate);
		args.GetReturnValue().Set(result);
	}
	else if (_wcsicmp(id, L"__methods") == 0) {
		if (self->methods.IsEmpty()) {
			self->initTypeInfo(isolate);
		}
		Local<Value> result = self->methods.Get(isolate);
		args.GetReturnValue().Set(result);
	}
	else if (_wcsicmp(id, L"__vars") == 0) {
		if (self->vars.IsEmpty()) {
			self->initTypeInfo(isolate);
		}
		Local<Value> result = self->vars.Get(isolate);
		args.GetReturnValue().Set(result);
	}
	else if (_wcsicmp(id, L"__proto__") == 0) {
		Local<Function> func;
		Local<FunctionTemplate> clazz = clazz_template.Get(isolate);
        Local<Context> ctx = isolate->GetCurrentContext();
        if (!clazz.IsEmpty() && clazz->GetFunction(ctx).ToLocal(&func)) {
            args.GetReturnValue().Set(func);
        }
        else {
            args.GetReturnValue().SetNull();
        }
	}
	else {
        Local<Function> func;
        if (clazz_methods.get(isolate, id, &func)) {
            args.GetReturnValue().Set(func);
        }

		else if (!self->get(id, -1, args)) {
			Local<Value> result;
			HRESULT hrcode = self->valueOf(isolate, args.This(), result);
			if FAILED(hrcode) isolate->ThrowException(Win32Error(isolate, hrcode, L"Unable to Get Value"));
			
            Local<Context> ctx = isolate->GetCurrentContext();
			MaybeLocal<Object> localObj = result->ToObject(ctx);
			if (localObj.IsEmpty()) {
				args.GetReturnValue().SetUndefined();
				return;
			}

			Local<Object> obj = localObj.ToLocalChecked();
			MaybeLocal<Value> realProp = obj->GetRealNamedPropertyInPrototypeChain(ctx, v8str(isolate, id));
			if (realProp.IsEmpty()) {
				// We may call non-existing property for an object to check its existence
				// So we should return undefined in this case
				args.GetReturnValue().SetUndefined();
			}
			else {
				Local<Value> ownProp = realProp.ToLocalChecked();
				if (ownProp->IsFunction()) {
					Local<Function> func = Local<Function>::Cast(ownProp);
					if (func.IsEmpty()) return;
					args.GetReturnValue().Set(func);
					return;
				}
			}
		}
	}
}

void DispObject::NodeGetByIndex(uint32_t index, const PropertyCallbackInfo<Value>& args) {
    Isolate *isolate = args.GetIsolate();
    DispObject *self = DispObject::Unwrap<DispObject>(args.This());
	if (!self) {
		isolate->ThrowException(DispErrorInvalid(isolate));
		return;
	}
    NODE_DEBUG_FMT2("DispObject '%S[%u]' get", self->name.c_str(), index);
    self->get(0, index, args);
}

void DispObject::NodeSet(Local<Name> name, Local<Value> value, const PropertyCallbackInfo<Value>& args) {
    Isolate *isolate = args.GetIsolate();
	DispObject *self = DispObject::Unwrap<DispObject>(args.This());
	if (!self) {
		isolate->ThrowException(DispErrorInvalid(isolate));
		return;
	}
	String::Value vname(isolate, name);
	LPOLESTR id = (vname.length() > 0) ? (LPOLESTR)*vname : L"";
	NODE_DEBUG_FMT2("DispObject '%S.%S' set", self->name.c_str(), id);
    self->set(id, -1, value, args);
}

void DispObject::NodeSetByIndex(uint32_t index, Local<Value> value, const PropertyCallbackInfo<Value>& args) {
    Isolate *isolate = args.GetIsolate();
    DispObject *self = DispObject::Unwrap<DispObject>(args.This());
	if (!self) {
		isolate->ThrowException(DispErrorInvalid(isolate));
		return;
	}
	NODE_DEBUG_FMT2("DispObject '%S[%u]' set", self->name.c_str(), index);
	self->set(0, index, value, args);
}

void DispObject::NodeCall(const FunctionCallbackInfo<Value> &args) {
    Isolate *isolate = args.GetIsolate();
    DispObject *self = DispObject::Unwrap<DispObject>(args.This());
	if (!self) {
		isolate->ThrowException(DispErrorInvalid(isolate));
		return;
	}
	NODE_DEBUG_FMT("DispObject '%S' call", self->name.c_str());
    self->call(isolate, args);
}

void DispObject::NodeValueOf(const FunctionCallbackInfo<Value>& args) {
	Isolate *isolate = args.GetIsolate();
	DispObject *self = DispObject::Unwrap<DispObject>(args.This());
	if (!self) {
		isolate->ThrowException(DispErrorInvalid(isolate));
		return;
	}
	Local<Value> result;
	HRESULT hrcode = self->valueOf(isolate, args.This(), result);
	if FAILED(hrcode) {
		isolate->ThrowException(Win32Error(isolate, hrcode, L"DispValueOf"));
		return;
	}
	args.GetReturnValue().Set(result);
}

void DispObject::NodeToString(const FunctionCallbackInfo<Value>& args) {
	Isolate *isolate = args.GetIsolate();
	DispObject *self = DispObject::Unwrap<DispObject>(args.This());
	if (!self) {
		isolate->ThrowException(DispErrorInvalid(isolate));
		return;
	}
	self->toString(args);
}

void DispObject::NodeRelease(const FunctionCallbackInfo<Value>& args) {
	Isolate *isolate = args.GetIsolate();
    int rcnt = 0, argcnt = args.Length();
    for (int argi = 0; argi < argcnt; argi++) {
        auto &obj = args[argi];
        if (obj->IsObject()) {
            auto disp_obj = Local<Object>::Cast(obj);
            DispObject *disp = DispObject::Unwrap<DispObject>(disp_obj);
            if (disp && disp->release())
                rcnt ++;
        }
    }
    args.GetReturnValue().Set(rcnt);
}

void DispObject::NodeCast(const FunctionCallbackInfo<Value>& args) {
	Local<Object> inst = VariantObject::NodeCreateInstance(args);
	args.GetReturnValue().Set(inst);
}

void DispObject::NodeConnectionPoints(const FunctionCallbackInfo<Value>& args) {
    Isolate *isolate = args.GetIsolate();
	Local<Context> ctx = isolate->GetCurrentContext();
    Local<Array> items = Array::New(isolate);
    CComPtr<IDispatch> ptr;
    CComPtr<IConnectionPointContainer> cp_cont;
    CComPtr<IEnumConnectionPoints> cp_enum;
    
    // prepare connecton points from arguments
    int argcnt = args.Length();
    if (argcnt >= 1) {
        if (Value2Unknown(isolate, args[0], (IUnknown**)&ptr)) {
            if SUCCEEDED(ptr->QueryInterface(&cp_cont)) {
                cp_cont->EnumConnectionPoints(&cp_enum);
            }
        }
    }

    // enumerate connection points
    if (cp_enum) {
        ULONG cnt_fetched;
        CComPtr<IConnectionPoint> cp_ptr;
        uint32_t cnt = 0;
        while (SUCCEEDED(cp_enum->Next(1, &cp_ptr, &cnt_fetched)) && cnt_fetched == 1) {
            items->Set(ctx, cnt++, ConnectionPointObject::NodeCreateInstance(isolate, cp_ptr, ptr));
            cp_ptr.Release();
        }
    }

    // return array of connection points
    args.GetReturnValue().Set(items);
}
void DispObject::PeakAndDispatchMessages(const FunctionCallbackInfo<Value>& args) {
    MSG msg;
    while(PeekMessage(&msg, 0, 0, 0, PM_REMOVE)) {
        DispatchMessage(&msg);
    }
}

//-------------------------------------------------------------------------------------------------------

class vtypes_t {
public:
    inline vtypes_t(std::initializer_list<std::pair<std::wstring, VARTYPE>> recs) {
        for (auto &rec : recs) {
            str2vt.emplace(rec.first, rec.second);
            vt2str.emplace(rec.second, rec.first);
        }
    }
    inline bool find(VARTYPE vt, std::wstring &name) {
        auto it = vt2str.find(vt);
        if (it == vt2str.end()) return false;
        name = it->second;
        return true;
    }
    inline VARTYPE find(const std::wstring &name) {
        auto it = str2vt.find(name);
        if (it == str2vt.end()) return VT_EMPTY;
        return it->second;
    }
private:
    std::map<std::wstring, VARTYPE> str2vt;
    std::map<VARTYPE, std::wstring> vt2str;
};

static vtypes_t vtypes({
	{ L"char", VT_I1 },
	{ L"uchar", VT_UI1 },
	{ L"byte", VT_UI1 },
	{ L"short", VT_I2 },
	{ L"ushort", VT_UI2 },
	{ L"int", VT_INT },
	{ L"uint", VT_UINT },
	{ L"long", VT_I8 },
	{ L"ulong", VT_UI8 },

	{ L"int8", VT_I1 },
	{ L"uint8", VT_UI1 },
	{ L"int16", VT_I2 },
	{ L"uint16", VT_UI2 },
	{ L"int32", VT_I4 },
	{ L"uint32", VT_UI4 },
	{ L"int64", VT_I8 },
	{ L"uint64", VT_UI8 },
	{ L"currency", VT_CY },

	{ L"float", VT_R4 },
	{ L"double", VT_R8 },
	{ L"date", VT_DATE },
	{ L"decimal", VT_DECIMAL },

	{ L"string", VT_BSTR },
	{ L"empty", VT_EMPTY },
	{ L"variant", VT_VARIANT },
	{ L"null", VT_NULL },
	{ L"byref", VT_BYREF }
});

bool VariantObject::assign(Isolate *isolate, Local<Value> &val, Local<Value> &type) {
	VARTYPE vt = VT_EMPTY;
	if (!type.IsEmpty()) {
		if (type->IsString()) {
			String::Value vtstr(isolate, type);
			const wchar_t *pvtstr = (const wchar_t *)*vtstr;
			int vtstr_len = vtstr.length();
			if (vtstr_len > 1) {
                if (pvtstr[0] == 'p') {
				    vt |= VT_BYREF;
				    vtstr_len--;
				    pvtstr++;
			    }
                else if (pvtstr[vtstr_len - 1] == '*') {
                    vt |= VT_BYREF;
                    vtstr_len--;
                }
                else if (pvtstr[vtstr_len - 2] == '[' || pvtstr[vtstr_len - 1] == ']') {
                    vt |= VT_ARRAY;
                    vtstr_len -= 2;
                }
            }
            if (vtstr_len > 0) {
				std::wstring type(pvtstr, vtstr_len);
                vt |= vtypes.find(type);
			}
		}
		else if (type->IsInt32()) {
			vt |= type->Int32Value(isolate->GetCurrentContext()).FromMaybe(0);
		}
	}

	if (val.IsEmpty()) {
		if FAILED(value.ChangeType(vt)) return false;
		if ((value.vt & VT_BYREF) == 0) pvalue.Clear();
		return true;
	}

	value.Clear();
	pvalue.Clear();
    if ((vt & VT_ARRAY) != 0) {
        Value2SafeArray(isolate, val, value, vt & ~VT_ARRAY);
    }
	else if ((vt & VT_BYREF) == 0) {
		Value2Variant(isolate, val, value, vt);
	}
	else {
		VARIANT *refvalue = nullptr;
		VARTYPE vt_noref = vt & ~VT_BYREF;
		VariantObject *ref = (!val.IsEmpty() && val->IsObject()) ? GetInstanceOf(isolate, Local<Object>::Cast(val)) : nullptr;
		if (ref) {
			if ((ref->value.vt & VT_BYREF) != 0) value = ref->value;
			else refvalue = &ref->value;
		}
		else {
			Value2Variant(isolate, val, pvalue, vt_noref);
			refvalue = &pvalue;
		}
		if (refvalue) {
			if (vt_noref == 0 || vt_noref == VT_VARIANT || refvalue->vt == VT_EMPTY) {
				value.vt = VT_VARIANT | VT_BYREF;
				value.pvarVal = refvalue;
			}
			else {
				value.vt = refvalue->vt | VT_BYREF;
				value.byref = &refvalue->intVal;
			}
		}
	}
	return true;
}

VariantObject::VariantObject(const FunctionCallbackInfo<Value> &args) {
	Local<Value> val, type;
	int argcnt = args.Length();
	if (argcnt > 0) val = args[0];
	if (argcnt > 1) type = args[1];
	assign(args.GetIsolate(), val, type);
}

void VariantObject::NodeInit(const Local<Object> &target, Isolate* isolate, Local<Context> &ctx) {

	// Prepare constructor template
	Local<FunctionTemplate> clazz = FunctionTemplate::New(isolate, NodeCreate);
	clazz->SetClassName(v8str(isolate, "Variant"));
    clazz_methods.add(isolate, clazz, "clear", NodeClear);
    clazz_methods.add(isolate, clazz, "assign", NodeAssign);
    clazz_methods.add(isolate, clazz, "cast", NodeCast);
    clazz_methods.add(isolate, clazz, "toString", NodeToString);
    clazz_methods.add(isolate, clazz, "valueOf", NodeValueOf);

	Local<ObjectTemplate> &inst = clazz->InstanceTemplate();
	inst->SetInternalFieldCount(1);
    inst->SetHandler(NamedPropertyHandlerConfiguration(NodeGet, NodeSet));
    inst->SetHandler(IndexedPropertyHandlerConfiguration(NodeGetByIndex, NodeSetByIndex));
    //inst->SetCallAsFunctionHandler(NodeCall);
	//inst->SetNativeDataProperty(v8str(isolate, "__id"), NodeGet);

	inst->SetNativeDataProperty(v8str(isolate, "__value"), NodeGet);
	//inst->SetLazyDataProperty(v8str(isolate, "__type"), NodeGet, Local<Value>(), ReadOnly);
	inst->SetNativeDataProperty(v8str(isolate, "__type"), NodeGet);

	inst_template.Reset(isolate, inst);
	clazz_template.Reset(isolate, clazz);
	Local<Function> func;
	if (clazz->GetFunction(ctx).ToLocal(&func)) {
		target->Set(ctx, v8str(isolate, "Variant"), func);
	}
	NODE_DEBUG_MSG("VariantObject initialized");
}

Local<Object> VariantObject::NodeCreateInstance(const FunctionCallbackInfo<Value> &args) {
	Local<Object> self;
	Isolate *isolate = args.GetIsolate();
	if (!inst_template.IsEmpty()) {
		if (inst_template.Get(isolate)->NewInstance(isolate->GetCurrentContext()).ToLocal(&self)) {
			(new VariantObject(args))->Wrap(self);
		}
	}
	return self;
}

void VariantObject::NodeCreate(const FunctionCallbackInfo<Value> &args) {
	Isolate *isolate = args.GetIsolate();
	Local<Object> &self = args.This();
	(new VariantObject(args))->Wrap(self);
	args.GetReturnValue().Set(self);
}

void VariantObject::NodeClear(const FunctionCallbackInfo<Value>& args) {
	Isolate *isolate = args.GetIsolate();
	VariantObject *self = VariantObject::Unwrap<VariantObject>(args.This());
	if (!self) {
		isolate->ThrowException(DispErrorInvalid(isolate));
		return;
	}
	self->value.Clear();
	self->pvalue.Clear();
}

void VariantObject::NodeAssign(const FunctionCallbackInfo<Value>& args) {
	Isolate *isolate = args.GetIsolate();
	VariantObject *self = VariantObject::Unwrap<VariantObject>(args.This());
	if (!self) {
		isolate->ThrowException(DispErrorInvalid(isolate));
		return;
	}
	Local<Value> val, type;
	int argcnt = args.Length();
	if (argcnt > 0) val = args[0];
	if (argcnt > 1) type = args[1];
	self->assign(isolate, val, type);
}

void VariantObject::NodeCast(const FunctionCallbackInfo<Value>& args) {
	Isolate *isolate = args.GetIsolate();
	VariantObject *self = VariantObject::Unwrap<VariantObject>(args.This());
	if (!self) {
		isolate->ThrowException(DispErrorInvalid(isolate));
		return;
	}
	Local<Value> val, type;
	int argcnt = args.Length();
	if (argcnt > 0) type = args[0];
	self->assign(isolate, val, type);
}

void VariantObject::NodeValueOf(const FunctionCallbackInfo<Value>& args) {
	Isolate *isolate = args.GetIsolate();
	VariantObject *self = VariantObject::Unwrap<VariantObject>(args.This());
	if (!self) {
		isolate->ThrowException(DispErrorInvalid(isolate));
		return;
	}
	// Last parameter false because valueOf should return primitive value
	Local<Value> result = Variant2Value(isolate, self->value, false);
	args.GetReturnValue().Set(result);
}

void VariantObject::NodeToString(const FunctionCallbackInfo<Value>& args) {
	Isolate *isolate = args.GetIsolate();
	VariantObject *self = VariantObject::Unwrap<VariantObject>(args.This());
	if (!self) {
		isolate->ThrowException(DispErrorInvalid(isolate));
		return;
	}
	Local<Value> result = Variant2String(isolate, self->value);
	args.GetReturnValue().Set(result);
}

void VariantObject::NodeGet(Local<Name> name, const PropertyCallbackInfo<Value>& args) {
	Isolate *isolate = args.GetIsolate();
	VariantObject *self = VariantObject::Unwrap<VariantObject>(args.This());
	if (!self) {
		isolate->ThrowException(DispErrorInvalid(isolate));
		return;
	}

	String::Value vname(isolate, name);
	LPOLESTR id = (vname.length() > 0) ? (LPOLESTR)*vname : L"valueOf";
	if (_wcsicmp(id, L"__value") == 0) {
		Local<Value> result = Variant2Value(isolate, self->value);
		args.GetReturnValue().Set(result);
	}
	else if (_wcsicmp(id, L"__type") == 0) {
		std::wstring type, name;
		if (self->value.vt & VT_BYREF) type += L"byref:";
		if (self->value.vt & VT_ARRAY) type = L"array:";
        if (vtypes.find(self->value.vt & VT_TYPEMASK, name)) type += name;
		else type += std::to_wstring(self->value.vt & VT_TYPEMASK);
		Local<String> text = v8str(isolate, type.c_str());
	}
	else if (_wcsicmp(id, L"__proto__") == 0) {
		Local<Function> func;
		Local<FunctionTemplate> clazz = clazz_template.Get(isolate);
        Local<Context> ctx = isolate->GetCurrentContext();
        if (!clazz.IsEmpty() && clazz_template.Get(isolate)->GetFunction(ctx).ToLocal(&func)) {
            args.GetReturnValue().Set(func);
        }
        else {
            args.GetReturnValue().SetNull();
        }
	}
    else if (_wcsicmp(id, L"length") == 0) {
        if ((self->value.vt & VT_ARRAY) != 0) {
            args.GetReturnValue().Set((uint32_t)self->value.ArrayLength());
        }
    }
    else {
        Local<Function> func;
        if (clazz_methods.get(isolate, id, &func)) {
            args.GetReturnValue().Set(func);
        }
    }
}

void VariantObject::NodeGetByIndex(uint32_t index, const PropertyCallbackInfo<Value>& args) {
	Isolate *isolate = args.GetIsolate();
	VariantObject *self = VariantObject::Unwrap<VariantObject>(args.This());
	if (!self) {
		isolate->ThrowException(DispErrorInvalid(isolate));
		return;
	}
	Local<Value> result;
	if ((self->value.vt & VT_ARRAY) == 0) {
		 result = Variant2Value(isolate, self->value);
	}
	else {
		CComVariant value;
		if SUCCEEDED(self->value.ArrayGet((LONG)index, value)) {
			result = Variant2Value(isolate, value);
		}
	}
	args.GetReturnValue().Set(result);
}

void VariantObject::NodeSet(Local<Name> name, Local<Value> val, const PropertyCallbackInfo<Value>& args) {
	Isolate *isolate = args.GetIsolate();
	VariantObject *self = VariantObject::Unwrap<VariantObject>(args.This());
	if (!self) {
		isolate->ThrowException(DispErrorInvalid(isolate));
		return;
	}
	isolate->ThrowException(DispError(isolate, E_NOTIMPL));
}

void VariantObject::NodeSetByIndex(uint32_t index, Local<Value> value, const PropertyCallbackInfo<Value>& args) {
	Isolate *isolate = args.GetIsolate();
	VariantObject *self = VariantObject::Unwrap<VariantObject>(args.This());
	if (!self) {
		isolate->ThrowException(DispErrorInvalid(isolate));
		return;
	}
	isolate->ThrowException(DispError(isolate, E_NOTIMPL));
}

Local<Object> VariantObject::NodeCreate(Isolate* isolate, const VARIANT& var) {
	Local<Object> self;
	if (!inst_template.IsEmpty()) {
		if (inst_template.Get(isolate)->NewInstance(isolate->GetCurrentContext()).ToLocal(&self)) {
			(new VariantObject(var))->Wrap(self);
		}
	}
	return self;
}

//-------------------------------------------------------------------------------------------------------

ConnectionPointObject::ConnectionPointObject(IConnectionPoint *p, IDispatch *d)
  : ptr(p), disp(d) {
    InitIndex();
}

bool ConnectionPointObject::InitIndex() {
  if (!ptr || !disp) {
    return false;
  }
  UINT typeindex = 0;
  CComPtr<ITypeInfo> typeinfo;
  if FAILED(disp->GetTypeInfo(typeindex, LOCALE_USER_DEFAULT, &typeinfo)) {
    return false;
  }

  CComPtr<ITypeLib> typelib;
  if FAILED(typeinfo->GetContainingTypeLib(&typelib, &typeindex)) {
    return false;
  }

  IID conniid;
  if FAILED(ptr->GetConnectionInterface(&conniid)) {
    return false;
  }

  CComPtr<ITypeInfo> conninfo;
  if FAILED(typelib->GetTypeInfoOfGuid(conniid, &conninfo)) {
    return false;
  }

  TYPEATTR *typeattr = nullptr;
  if FAILED(conninfo->GetTypeAttr(&typeattr)) {
    return false;
  }

  if (typeattr->typekind != TKIND_DISPATCH) {
    conninfo->ReleaseTypeAttr(typeattr);
    return false;
  }

  for (UINT fd = 0; fd < typeattr->cFuncs; ++fd) {
    FUNCDESC *funcdesc;
    if FAILED(conninfo->GetFuncDesc(fd, &funcdesc)) {
      continue;
    }
    if (!funcdesc) {
      break;
    }

    if (funcdesc->invkind != INVOKE_FUNC || funcdesc->funckind != FUNC_DISPATCH) {
      conninfo->ReleaseFuncDesc(funcdesc);
      continue;
    }

    // const size_t nameSize = 256;
    const size_t nameSize = 1; // only event function name required
    BSTR bstrNames[nameSize];
    UINT maxNames = nameSize;
    UINT maxNamesOut = 0;
    if SUCCEEDED(conninfo->GetNames(funcdesc->memid, reinterpret_cast<BSTR *>(&bstrNames), maxNames, &maxNamesOut)) {
      DISPID id = funcdesc->memid;
      std::wstring funcname(bstrNames[0]);
      index.insert(std::pair<DISPID, DispObjectImpl::name_ptr>(id, new DispObjectImpl::name_t(id, funcname)));

      for (size_t i = 0; i < maxNamesOut; i++) {
        SysFreeString(bstrNames[i]);
      }
    }

    conninfo->ReleaseFuncDesc(funcdesc);
  }

  conninfo->ReleaseTypeAttr(typeattr);

  return true;
}

Local<Object> ConnectionPointObject::NodeCreateInstance(Isolate *isolate, IConnectionPoint *p, IDispatch* d) {
    Local<Object> self;
    if (!inst_template.IsEmpty()) {
		if (inst_template.Get(isolate)->NewInstance(isolate->GetCurrentContext()).ToLocal(&self)) {
			(new ConnectionPointObject(p, d))->Wrap(self);
		}
    }
    return self;
}

void ConnectionPointObject::NodeInit(const Local<Object> &target, Isolate* isolate, Local<Context> &ctx) {

    // Prepare constructor template
    Local<FunctionTemplate> clazz = FunctionTemplate::New(isolate, NodeCreate);
	clazz->SetClassName(v8str(isolate, "ConnectionPoint"));

    NODE_SET_PROTOTYPE_METHOD(clazz, "advise", NodeAdvise);
	NODE_SET_PROTOTYPE_METHOD(clazz, "unadvise", NodeUnadvise);
	NODE_SET_PROTOTYPE_METHOD(clazz, "getMethods", NodeConnectionPointMethods);

    Local<ObjectTemplate> &inst = clazz->InstanceTemplate();
    inst->SetInternalFieldCount(1);

    inst_template.Reset(isolate, inst);
    clazz_template.Reset(isolate, clazz);
    //target->Set(v8str(isolate, "ConnectionPoint"), clazz->GetFunction());
    NODE_DEBUG_MSG("ConnectionPointObject initialized");
}

void ConnectionPointObject::NodeCreate(const FunctionCallbackInfo<Value> &args) {
    Isolate *isolate = args.GetIsolate();
    Local<Object> &self = args.This();
    (new ConnectionPointObject(args))->Wrap(self);
    args.GetReturnValue().Set(self);
}

void ConnectionPointObject::NodeAdvise(const FunctionCallbackInfo<Value> &args) {
    Isolate *isolate = args.GetIsolate();
    ConnectionPointObject *self = ConnectionPointObject::Unwrap<ConnectionPointObject>(args.This());
    if (!self || !self->ptr) {
        isolate->ThrowException(DispErrorInvalid(isolate));
        return;
    }
    CComPtr<IUnknown> unk;
    int argcnt = args.Length();
    if (argcnt > 0) {
        Local<Value> val = args[0];
        if (!Value2Unknown(isolate, val, &unk)) {
            Local<Object> obj;
            if (!val.IsEmpty() && val->IsObject() && val->ToObject(isolate->GetCurrentContext()).ToLocal(&obj)) {

                // .NET Connection Points require to implement specific interface
                // So we need to remember its IID for the case when Container does QueryInterface for it
                IID connif;
                self->ptr->GetConnectionInterface(&connif);
                DispObjectImpl *impl = new DispObjectImpl(obj, false, connif);
                // It requires reversed arguments
                impl->reverse_arguments = true;
                impl->index = self->index;
                if (self->index.size()) {
                    impl->dispid_next = self->index.rbegin()->first + 1;
                }
                unk.Attach(impl);
            }
        }
    }
    if (!unk) {
        isolate->ThrowException(InvalidArgumentsError(isolate));
        return;
    }
    DWORD dwCookie;
    HRESULT hrcode = self->ptr->Advise(unk, &dwCookie);
    if FAILED(hrcode) {
        isolate->ThrowException(DispError(isolate, hrcode));
        return;
    }
    self->cookies.insert(dwCookie);
    args.GetReturnValue().Set(v8::Integer::New(isolate, (uint32_t)dwCookie));
}

void ConnectionPointObject::NodeUnadvise(const FunctionCallbackInfo<Value> &args) {
    Isolate *isolate = args.GetIsolate();
    Local<Context> ctx = isolate->GetCurrentContext();
    ConnectionPointObject *self = ConnectionPointObject::Unwrap<ConnectionPointObject>(args.This());
    if (!self || !self->ptr) {
        isolate->ThrowException(DispErrorInvalid(isolate));
        return;
    }

    if (args.Length()==0 || !args[0]->IsUint32()) {
        isolate->ThrowException(InvalidArgumentsError(isolate));
        return;
    }
    DWORD dwCookie = (args[0]->Uint32Value(ctx)).FromMaybe(0);
    if (dwCookie == 0 || self->cookies.find(dwCookie) == self->cookies.end()) {
        isolate->ThrowException(InvalidArgumentsError(isolate));
    return;
    }
	self->cookies.erase(dwCookie);
    HRESULT hrcode = self->ptr->Unadvise(dwCookie);
    if FAILED(hrcode) {
        isolate->ThrowException(DispError(isolate, hrcode));
        return;
    }
}

void ConnectionPointObject::NodeConnectionPointMethods(const FunctionCallbackInfo<Value>& args) {
	Isolate* isolate = args.GetIsolate();
	Local<Context> ctx = isolate->GetCurrentContext();
	Local<Array> items = Array::New(isolate);

	ConnectionPointObject* self = ConnectionPointObject::Unwrap<ConnectionPointObject>(args.This());

	DispObjectImpl::index_t::iterator it;
	uint32_t cnt = 0;

	for (it = self->index.begin(); it != self->index.end(); it++)
	{
		items->Set(ctx, cnt++, v8str(isolate, it->second->name.c_str()));
	}

	args.GetReturnValue().Set(items);
}

//-------------------------------------------------------------------------------------------------------
