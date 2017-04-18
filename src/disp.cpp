//-------------------------------------------------------------------------------------------------------
// Project: NodeActiveX
// Author: Yuri Dursin
// Last Modification: 2017-04-16
// Description: DispObject class implementations
//-------------------------------------------------------------------------------------------------------

#include "stdafx.h"
#include "disp.h"

Persistent<ObjectTemplate> DispObject::inst_template;
Persistent<Function> DispObject::constructor;

//-------------------------------------------------------------------------------------------------------
// DispObject implemetation

DispObject::DispObject(const DispInfoPtr &ptr, LPCOLESTR nm, DISPID id, LONG indx)
	: state(None), name(nm), dispid(id), index(indx)
{	
	disp = ptr;
	if (dispid == DISPID_UNKNOWN) {
		dispid = DISPID_VALUE;
		state |= state_prepared;
	}
	else state |= state_owned;
	NODE_DEBUG_FMT("DispObject '%S' constructor", name.c_str());
}

DispObject::~DispObject() {
	NODE_DEBUG_FMT("DispObject '%S' destructor", name.c_str());
}

HRESULT DispObject::Prepare(VARIANT *value) {
	CComVariant val; 
	if (!value) value = &val;
	HRESULT hrcode = disp->GetProperty(dispid, index, value);

	// Init dispatch interface
	if (!is_prepared()) {
		state |= state_prepared;
		CComPtr<IDispatch> ptr;
		if (VariantDispGet(value, &ptr)) {
			DispInfoPtr disp_new(new DispInfo(ptr));
			disp_new->parent = disp;
			disp = disp_new;
			dispid = DISPID_VALUE;
		}
	}

	return hrcode;
}

bool DispObject::get(LPOLESTR name, uint32_t index, const PropertyCallbackInfo<Value> &args) {
	Isolate *isolate = args.GetIsolate();

	// Prepare 
	if (!is_prepared()) Prepare();

	// Search dispid
	DISPID dispid;
	HRESULT hrcode = disp->FindProperty(name, &dispid);
	if (SUCCEEDED(hrcode) && dispid == DISPID_UNKNOWN) hrcode = E_INVALIDARG;
	if FAILED(hrcode) {
		isolate->ThrowException(DispError(isolate, hrcode, L"DispPropertyFind"));
		return false;
	}

    // Return as property value
	if (disp->IsProperty(dispid)) {
		CComPtr<IDispatch> ptr;
		CComVariant value;
		hrcode = disp->GetProperty(dispid, -1, &value);
		if FAILED(hrcode) {
			isolate->ThrowException(DispError(isolate, hrcode, L"DispPropertyGet"));
			return false;
		}
		if (VariantDispGet(&value, &ptr)) {
			DispInfoPtr disp_result(new DispInfo(ptr));
			disp_result->parent = disp;
			Local<Object> result = DispObject::NodeCreate(isolate, args.This(), disp_result, name);
			args.GetReturnValue().Set(result);
		}
		else {
			args.GetReturnValue().Set(Variant2Value(isolate, value));
		}
	}

	// Return as dispatch object 
	else {
		Local<Object> result = DispObject::NodeCreate(isolate, args.This(), disp, name, dispid);
		args.GetReturnValue().Set(result);
	}
	return true;
}

bool DispObject::get(uint32_t index, const PropertyCallbackInfo<Value> &args) {
	Isolate *isolate = args.GetIsolate();

	// Prepare 
	if (!is_prepared()) Prepare();

	Local<Object> result = DispObject::NodeCreate(isolate, args.This(), disp, name.c_str(), dispid, index);
	args.GetReturnValue().Set(result);
	return true;
}

bool DispObject::set(Isolate *isolate, LPOLESTR name, Local<Value> value) {
	// Prepare disp
	if (!is_prepared()) Prepare();

	// Set value using dispatch
    CComVariant ret;
    VarArgumets vargs(value);
    size_t argcnt = vargs.items.size();
    VARIANT *pargs = (argcnt > 0) ? &vargs.items.front() : 0;
    HRESULT hrcode = disp->SetProperty(dispid, argcnt, pargs, &ret);
	//HRESULT hrcode = DispInvoke(disp, name, 0, 0, &ret, DISPATCH_PROPERTYPUT, &dispid);
	if FAILED(hrcode) {
		isolate->ThrowException(DispError(isolate, hrcode, L"DispPropertyPut"));
        return false;
    }
    return true; // DispObject::NodeCreate(isolate, disp, name, dispid);
}

void DispObject::call(Isolate *isolate, const FunctionCallbackInfo<Value> &args)
{
	CComVariant ret;
	VarArgumets vargs(args);
	size_t argcnt = vargs.items.size();
	VARIANT *pargs = (argcnt > 0) ? &vargs.items.front() : 0;
	HRESULT hrcode = disp->ExecuteMethod(dispid, argcnt, pargs, &ret);
    if FAILED(hrcode) {
        isolate->ThrowException(DispError(isolate, hrcode, L"DispInvoke"));
        return;
    }

	// Prepare result
    Local<Value> result;
	CComPtr<IDispatch> ptr;
	if (VariantDispGet(&ret, &ptr)) {
		DispInfoPtr disp_result(new DispInfo(ptr));
		disp_result->parent = disp;
		result = DispObject::NodeCreate(isolate, args.This(), disp_result, L"Result");
	}
	else {
		result = Variant2Value(isolate, ret);
	}
    args.GetReturnValue().Set(result);
}

HRESULT DispObject::valueOf(Isolate *isolate, Local<Value> &value) {
	CComVariant val;
	HRESULT hrcode = Prepare(&val);
	if SUCCEEDED(hrcode) value = Variant2Value(isolate, val);
	return hrcode;
}

void DispObject::toString(const FunctionCallbackInfo<Value> &args) {
	Isolate *isolate = args.GetIsolate();
	CComVariant val;
	HRESULT hrcode = Prepare(&val);
	if FAILED(hrcode) {
		isolate->ThrowException(Win32Error(isolate, hrcode, L"DispToString"));
		return;
	}
	args.GetReturnValue().Set(Variant2Value(isolate, val));
}

//-----------------------------------------------------------------------------------
// Static Node JS callbacks

void DispObject::NodeInit(Handle<Object> target) {
    Isolate *isolate = target->GetIsolate();

    // Prepare constructor template
    Local<String> prop_name(String::NewFromUtf8(isolate, "Object"));
    Local<String> clazz_name(String::NewFromUtf8(isolate, "Dispatch"));
    Local<FunctionTemplate> clazz = FunctionTemplate::New(isolate, NodeCreate);
    clazz->SetClassName(clazz_name);

	NODE_SET_PROTOTYPE_METHOD(clazz, "toString", NodeToString);
	NODE_SET_PROTOTYPE_METHOD(clazz, "valueOf", NodeValueOf);

    Local<ObjectTemplate> &inst = clazz->InstanceTemplate();
    inst->SetInternalFieldCount(1);
    inst->SetNamedPropertyHandler(NodeGet, NodeSet);
    inst->SetIndexedPropertyHandler(NodeGetByIndex, NodeSetByIndex);
    inst->SetCallAsFunctionHandler(NodeCall);
	inst->SetNativeDataProperty(String::NewFromUtf8(isolate, "__id"), NodeGet);
	inst->SetNativeDataProperty(String::NewFromUtf8(isolate, "__value"), NodeGet);

    inst_template.Reset(isolate, inst);
    constructor.Reset(isolate, clazz->GetFunction());
    target->Set(prop_name, clazz->GetFunction());

    //Context::GetCurrent()->Global()->Set(String::NewFromUtf8("ActiveXObject"), t->GetFunction());
	NODE_DEBUG_MSG("DispObject initialized");
}

Local<Object> DispObject::NodeCreate(Isolate *isolate, const Local<Object> &parent, const DispInfoPtr &ptr, LPCOLESTR name, DISPID id, LONG index) {
    Local<Object> self;
    if (!inst_template.IsEmpty()) {
        self = inst_template.Get(isolate)->NewInstance();
        (new DispObject(ptr, name, id, index))->Wrap(self);
		//Local<String> prop_id(String::NewFromUtf8(isolate, "_identity"));
		//self->Set(prop_id, String::NewFromTwoByte(isolate, (uint16_t*)name));
	}
    return self;
}

void DispObject::NodeCreate(const FunctionCallbackInfo<Value> &args) {
    Isolate *isolate = args.GetIsolate();
    if (args.Length() < 1) {
        isolate->ThrowException(TypeError(isolate, "innvalid arguments"));
        return;
    }
    
    // Invoked as plain function
    if (!args.IsConstructCall()) {
        const int argc = 1;
        Local<Value> argv[argc] = { args[0] };
        Local<Context> context = isolate->GetCurrentContext();
        Local<Function> cons = Local<Function>::New(isolate, constructor);
        Local<Object> self = cons->NewInstance(context, argc, argv).ToLocalChecked();
        args.GetReturnValue().Set(self);
        return;
    }

    // Prepare arguments
	Local<String> progid = args[0]->ToString();
	String::Value vname(progid);
	LPOLESTR name = (vname.length() > 0) ? (LPOLESTR)*vname : 0;

    // Initialize COM object
	CComPtr<IDispatch> disp;
	HRESULT hrcode = name ? disp.CoCreateInstance(name) : E_INVALIDARG;
	if FAILED(hrcode) isolate->ThrowException(DispError(isolate, hrcode, L"CoCreateInstance"));

	// Create object
    else {
        Local<Object> &self = args.This();
		DispInfoPtr ptr(new DispInfo(disp));
		(new DispObject(ptr, name))->Wrap(self);
        args.GetReturnValue().Set(self);
    }
}

void DispObject::NodeGet(Local<String> name, const PropertyCallbackInfo<Value>& args) {
    Isolate *isolate = args.GetIsolate();
	DispObject *self = DispObject::Unwrap<DispObject>(args.This());
	if (!self) {
		isolate->ThrowException(Error(isolate, "DispIsEmpty"));
		return;
	}
	
	String::Value vname(name);
	LPOLESTR id = (vname.length() > 0) ? (LPOLESTR)*vname : L"";
    NODE_DEBUG_FMT2("DispObject '%S.%S' get", self->name.c_str(), id);
	if (_wcsicmp(id, L"__value") == 0) {
		Local<Value> result;
		HRESULT hrcode = self->valueOf(isolate, result);
		if FAILED(hrcode) isolate->ThrowException(Win32Error(isolate, hrcode, L"DispValueOf"));
		else args.GetReturnValue().Set(result);
	}
	else if (_wcsicmp(id, L"__id") == 0) {
		args.GetReturnValue().Set(String::NewFromTwoByte(isolate, (uint16_t*)self->name.c_str()));
	}
	else if (_wcsicmp(id, L"valueOf") == 0) {
		args.GetReturnValue().Set(FunctionTemplate::New(isolate, NodeValueOf, args.This())->GetFunction());
	}
	else if (_wcsicmp(id, L"toString") == 0) {
		args.GetReturnValue().Set(FunctionTemplate::New(isolate, NodeToString, args.This())->GetFunction());
	}
	else {
		self->get(id, -1, args);
	}
}

void DispObject::NodeGetByIndex(uint32_t index, const PropertyCallbackInfo<Value>& args) {
    Isolate *isolate = args.GetIsolate();
    DispObject *self = DispObject::Unwrap<DispObject>(args.This());
	if (!self) {
		isolate->ThrowException(Error(isolate, "DispIsEmpty"));
		return;
	}
    NODE_DEBUG_FMT2("DispObject '%S[%u]' get", self->name.c_str(), index);
    self->get(index, args);
}

void DispObject::NodeSet(Local<String> name, Local<Value> value, const PropertyCallbackInfo<Value>& args) {
    Isolate *isolate = args.GetIsolate();
	DispObject *self = DispObject::Unwrap<DispObject>(args.This());
	if (!self) {
		isolate->ThrowException(Error(isolate, "DispIsEmpty"));
		return;
	}
	String::Value vname(name);
	LPOLESTR id = (vname.length() > 0) ? (LPOLESTR)*vname : L"";
	NODE_DEBUG_FMT2("DispObject '%S.%S' set", self->name.c_str(), id);
    if (self->set(isolate, id, value)) 
        args.GetReturnValue().Set(args.This());
}

void DispObject::NodeSetByIndex(uint32_t index, Local<Value> value, const PropertyCallbackInfo<Value>& args) {
    Isolate *isolate = args.GetIsolate();
    DispObject *self = DispObject::Unwrap<DispObject>(args.This());
	if (!self) {
		isolate->ThrowException(Error(isolate, "DispIsEmpty"));
		return;
	}
	NODE_DEBUG_FMT2("DispObject '%S[%u]' set", self->name.c_str(), index);
	isolate->ThrowException(Win32Error(isolate, E_NOTIMPL, __FUNCTIONW__));
}

void DispObject::NodeCall(const FunctionCallbackInfo<Value> &args) {
    Isolate *isolate = args.GetIsolate();
    DispObject *self = DispObject::Unwrap<DispObject>(args.This());
	if (!self) {
		isolate->ThrowException(Error(isolate, "DispIsEmpty"));
		return;
	}
	NODE_DEBUG_FMT("DispObject '%S' call", self->name.c_str());
    self->call(isolate, args);
}

void DispObject::NodeValueOf(const FunctionCallbackInfo<Value>& args) {
	Isolate *isolate = args.GetIsolate();
	DispObject *self = DispObject::Unwrap<DispObject>(args.This());
	if (!self) {
		isolate->ThrowException(Error(isolate, "DispIsEmpty"));
		return;
	}
	Local<Value> result;
	HRESULT hrcode = self->valueOf(isolate, result);
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
		isolate->ThrowException(Error(isolate, "DispIsEmpty"));
		return;
	}
	self->toString(args);
}

//-------------------------------------------------------------------------------------------------------
