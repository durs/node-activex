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

DispObject::DispObject(const DispInfoPtr ptr, LPCOLESTR nm, DISPID id, LONG indx)
	: state(None), name(nm), dispid(id), index(indx)
{	
	disp = ptr;
	if (disp) state |= Prepared;
	if (dispid == DISPID_UNKNOWN) dispid = DISPID_VALUE;
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
	if ((state & Prepared) == 0) {
		state |= Prepared;
		CComPtr<IDispatch> ptr;
		if (VariantDispGet(value, &ptr)) {
			disp.reset(new DispInfo(ptr));
			dispid = DISPID_VALUE;
		}
	}

	return hrcode;
}

Local<Value> DispObject::find(Isolate *isolate, LPOLESTR name) {
	// Prepare disp
	if (!isPrepared()) Prepare();

	// Search dispid
	DISPID dispid;
	if FAILED(disp->FindProperty(name, &dispid)) return Undefined(isolate);
	if (dispid == DISPID_UNKNOWN) return Undefined(isolate);

    // Try preapre as property
    //CComVariant ret;
    //if (DispInvoke(disp, dispid, 0, 0, &ret, DISPATCH_PROPERTYGET) == S_OK) {
    //    if (ret.vt != VT_EMPTY && ret.vt != VT_ERROR) { // RecordSet.MoveNext return VT_EMPTY
    //        if ((ret.vt & VT_TYPEMASK) != VT_DISPATCH) return Variant2Value(isolate, ret);
    //        pdisp = ((ret.vt & VT_BYREF) != 0) ? *ret.ppdispVal : ret.pdispVal;
    //        if (!pdisp) return Null(isolate);
    //        return DispObject::NodeCreate(isolate, pdisp, name);
    //    }
   // }

	return DispObject::NodeCreate(isolate, disp, name, dispid);
}

Local<Value> DispObject::find(Isolate *isolate, uint32_t index) {
	//if (!isOwned()) Undefined(isolate);
	return DispObject::NodeCreate(isolate, disp, name.c_str(), dispid, index);
}

Local<Value> DispObject::valueOf(Isolate *isolate) {
	CComVariant value;
	HRESULT hrcode = Prepare(&value);
	//if FAILED(hrcode) return DispError(isolate, hrcode, L"DispPropertyGet");
	if FAILED(hrcode) return Undefined(isolate); 
	return Variant2Value(isolate, value);
}

Local<Value> DispObject::toString(Isolate *isolate) {
	CComVariant value;
	HRESULT hrcode = Prepare(&value);
	if FAILED(hrcode) return String::NewFromTwoByte(isolate, (uint16_t*)name.c_str());
	return Variant2Value(isolate, value);
}

bool DispObject::set(Isolate *isolate, LPOLESTR name, Local<Value> value) {
	// Prepare disp
	if (!isPrepared()) Prepare();
	/*
    if (!disp) {
        isolate->ThrowException(Win32Error(isolate, E_NOTIMPL, __FUNCTIONW__));
        return false;
    }
	*/

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
    if (VariantDispGet(&ret, &ptr))
        result = DispObject::NodeCreate(isolate, std::make_shared<DispInfo>(ptr), L"Result");
    else 
        result = Variant2Value(isolate, ret);
    args.GetReturnValue().Set(result);
}

//-----------------------------------------------------------------------------------
// Static Node JS callbacks

void DispObject::NodeInit(Handle<Object> target)
{
    Isolate *isolate = target->GetIsolate();

    // Prepare constructor template
    Local<String> prop_name(String::NewFromUtf8(isolate, "Object"));
    Local<String> clazz_name(String::NewFromUtf8(isolate, "Dispatch"));
    Local<FunctionTemplate> clazz = FunctionTemplate::New(isolate, NodeCreate);
    clazz->SetClassName(clazz_name);
 
    Local<ObjectTemplate> &inst = clazz->InstanceTemplate();
    inst->SetInternalFieldCount(1);
    inst->SetNamedPropertyHandler(NodeGet, NodeSet);
    inst->SetIndexedPropertyHandler(NodeGetByIndex, NodeSetByIndex);
    inst->SetCallAsFunctionHandler(NodeCall);

    //inst->SetAccessCheckCallback
    //inst->SetAccessCheckCallbacks(Engine::GetNamedItem, 0, )
    //inst.MakeWeak(0, DestroyInstance);

    //NODE_SET_PROTOTYPE_METHOD(clazz, "toString", NodeToString);
    //NODE_SET_PROTOTYPE_METHOD(clazz, "valueOf", NodeValueOf);

    inst_template.Reset(isolate, inst);
    constructor.Reset(isolate, clazz->GetFunction());
    target->Set(prop_name, clazz->GetFunction());

    //Context::GetCurrent()->Global()->Set(String::NewFromUtf8("ActiveXObject"), t->GetFunction());
	NODE_DEBUG_MSG("DispObject initialized");
}

Local<Object> DispObject::NodeCreate(Isolate *isolate, const DispInfoPtr &ptr, LPCOLESTR name, DISPID id, LONG index)
{
    Local<Object> self;
    if (!inst_template.IsEmpty()) {
        self = inst_template.Get(isolate)->NewInstance();
        (new DispObject(ptr, name, id, index))->Wrap(self);
    }
    return self;
}

void DispObject::NodeCreate(const FunctionCallbackInfo<Value> &args)
{
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

void DispObject::NodeValueOf(const FunctionCallbackInfo<Value>& args)
{
    Isolate *isolate = args.GetIsolate();
	DispObject *self = DispObject::Unwrap<DispObject>(args.This());
    if (self) {
        Local<Value> &val = self->valueOf(isolate);
        if (!val.IsEmpty() && !val->IsUndefined()) {
            args.GetReturnValue().Set(val);
            return;
        }
    }
    args.GetReturnValue().Set(args.This());
}

void DispObject::NodeToString(const FunctionCallbackInfo<Value>& args)
{
    Isolate *isolate = args.GetIsolate();
	DispObject *self = DispObject::Unwrap<DispObject>(args.This());
	if (!self) args.GetReturnValue().Set(Undefined(isolate));
    else args.GetReturnValue().Set(self->toString(isolate));
}

void DispObject::NodeGet(Local<String> name, const PropertyCallbackInfo<Value>& args)
{
    Isolate *isolate = args.GetIsolate();
	String::Value vname(name);
	LPOLESTR id = (vname.length() > 0) ? (LPOLESTR)*vname : 0;
	DispObject *self = DispObject::Unwrap<DispObject>(args.This());
    if (!id || !self) return;
    NODE_DEBUG_FMT2("DispObject '%S.%S' get", self->name.c_str(), id);
    Local<Value> result;

    if (_wcsicmp(id, L"valueOf") == 0)
        result = FunctionTemplate::New(isolate, NodeValueOf, args.This())->GetFunction();
    else if (_wcsicmp(id, L"toString") == 0)
        result = FunctionTemplate::New(isolate, NodeToString, args.This())->GetFunction();
    else
        result = self->find(isolate, id);
    args.GetReturnValue().Set(result);
}

void DispObject::NodeGetByIndex(uint32_t index, const PropertyCallbackInfo<Value>& args)
{
    Isolate *isolate = args.GetIsolate();
    DispObject *self = DispObject::Unwrap<DispObject>(args.This());
    if (!self) return;
    NODE_DEBUG_FMT2("DispObject '%S[%u]' get", self->name.c_str(), index);
    Local<Value> result = self->find(isolate, index);
    args.GetReturnValue().Set(result);
}

void DispObject::NodeSet(Local<String> name, Local<Value> value, const PropertyCallbackInfo<Value>& args)
{
    Isolate *isolate = args.GetIsolate();
	String::Value vname(name);
	LPOLESTR id = (vname.length() > 0) ? (LPOLESTR)*vname : 0;
	DispObject *self = DispObject::Unwrap<DispObject>(args.This());
	if (!id || !self) return;
	NODE_DEBUG_FMT2("DispObject '%S.%S' set", self->name.c_str(), id);
    if (self->set(isolate, id, value)) 
        args.GetReturnValue().Set(args.This());
}

void DispObject::NodeSetByIndex(uint32_t index, Local<Value> value, const PropertyCallbackInfo<Value>& args)
{
    Isolate *isolate = args.GetIsolate();
    DispObject *self = DispObject::Unwrap<DispObject>(args.This());
	if (!self) return;
	NODE_DEBUG_FMT2("DispObject '%S[%u]' set", self->name.c_str(), index);
	isolate->ThrowException(Win32Error(isolate, E_NOTIMPL, __FUNCTIONW__));
}

void DispObject::NodeCall(const FunctionCallbackInfo<Value> &args)
{
    Isolate *isolate = args.GetIsolate();
    DispObject *self = DispObject::Unwrap<DispObject>(args.This());
	if (!self) return;
	NODE_DEBUG_FMT("DispObject '%S' call", self->name.c_str());
    self->call(isolate, args);
}

//-------------------------------------------------------------------------------------------------------
