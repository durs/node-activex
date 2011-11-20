//-------------------------------------------------------------------------------------------------------
// Project: NodeActiveX
// Author: Yuri Dursin
// Last Modification: 2011-11-22
// Description: DispObject class implementations
//-------------------------------------------------------------------------------------------------------

#include "stdafx.h"
#include "DispNodeJS.h"

Persistent<ObjectTemplate> DispObject::tmplate;

//-------------------------------------------------------------------------------------------------------
// DispObject implemetation

DispObject::DispObject(IDispatch *ptr, LPCOLESTR nm, DISPID id, LONG indx): 
	state(None), name(nm), owner_id(id), index(indx)
{ 
	if (owner_id == DISPID_UNKNOWN) disp = ptr;
	else owner_disp = ptr;
	if (disp) state |= Prepared;
	NODE_DEBUG_FMT("DispObject '%S' constructor", name.GetString());
}

DispObject::~DispObject()
{
	NODE_DEBUG_FMT("DispObject '%S' destructor", name.GetString());
}

HRESULT DispObject::Prepare(VARIANT *value)
{
	IDispatch *pdisp = (disp) ? disp : owner_disp;
	DISPID dispid = (disp) ? DISPID_VALUE : owner_id;
	CComVariant val; 
	if (!value) value = &val;

	// Get value using dispatch
	CComVariant arg(index);
	LONG argcnt = (index >= 0) ? 1 : 0;
	HRESULT hrcode = DispInvoke(pdisp, dispid, argcnt, &arg, value, DISPATCH_PROPERTYGET);
	if FAILED(hrcode) value->vt = VT_EMPTY;

	// Init dispatch interface
	if ((state & Prepared) == 0)
	{
		state |= Prepared;
		VariantDispGet(value, &disp);
	}

	return hrcode;
}

Handle<Value> DispObject::findElement(LPOLESTR name)
{
	// Prepare disp
	if (!isPrepared()) Prepare();

	// Search dispid
	DISPID dispid;
	IDispatch *pdisp = (disp) ? disp : owner_disp;
	if (!pdisp || FAILED(DispFind(pdisp, name, &dispid))) return Undefined();
	if (dispid == DISPID_UNKNOWN) return Undefined();
	return DispObject::NodeCreate(disp, name, dispid);
}

Handle<Value> DispObject::findElement(uint32_t index)
{
	if (!isOwned()) Undefined();
	return DispObject::NodeCreate(owner_disp, name.GetString(), owner_id, index);
}

Handle<Value> DispObject::valueOf()
{
	CComVariant value;
	HRESULT hrcode = Prepare(&value);
	//if FAILED(hrcode) return DispError(hrcode, L"DispPropertyGet");
	if FAILED(hrcode) return Undefined(); 
	return Variant2Value(value);
}

Handle<Value> DispObject::toString()
{
	CComVariant value;
	HRESULT hrcode = Prepare(&value);
	if FAILED(hrcode) return String::New((uint16_t*)name.GetString());
	return Variant2Value(value);
}

Handle<Value> DispObject::set(LPOLESTR name, Local<Value> value)
{
	// Prepare disp
	if (!isPrepared()) Prepare();
	if (!disp) return DispError(E_NOTIMPL, __FUNCTIONW__);

	// Set value using dispatch
	DISPID dispid;
	CComVariant ret;
	VarArgumets vargs(value);
	HRESULT hrcode = DispInvoke(disp, name, 0, 0, &ret, DISPATCH_PROPERTYPUT, &dispid);
	if FAILED(hrcode) return DispError(hrcode, L"DispPropertyPut");
	return DispObject::NodeCreate(disp, name, dispid);
}

Handle<Value> DispObject::call(const Arguments &args)
{
	CComVariant ret;
	VarArgumets vargs(args);
	IDispatch *pdisp = (disp) ? disp : owner_disp;
	DISPID dispid = (disp) ? DISPID_VALUE : owner_id;
	size_t argcnt = vargs.items.size();
	VARIANT *argptr = (argcnt > 0) ? &vargs.items.front() : 0;
	HRESULT hrcode = DispInvoke(pdisp, dispid, argcnt, argptr, &ret, DISPATCH_METHOD);
	if FAILED(hrcode) return DispError(hrcode, L"DispMethod");

	// Prepare result as object
	CComPtr<IDispatch> ret_disp;
	if (VariantDispGet(&ret, &ret_disp)) 
		return DispObject::NodeCreate(ret_disp, /*name.GetString()*/ L"Result");

	// Other result
	return Variant2Value(ret);
}

//-----------------------------------------------------------------------------------
// Static Node JS callbacks

void DispObject::NodeInit(Handle<Object> target)
{
    HandleScope scope;

	// Prepare disp template
	Local<FunctionTemplate> t = FunctionTemplate::New(NodeCreate);

	Local<ObjectTemplate> &t_obj = t->InstanceTemplate();
	t_obj->SetInternalFieldCount(1);
	t_obj->SetNamedPropertyHandler(NodeGet, NodeSet, 0, 0, 0, Handle<Value>());
	t_obj->SetIndexedPropertyHandler(NodeGetByIndex, NodeSetByIndex, 0, 0, 0, Handle<Value>());
	t_obj->SetCallAsFunctionHandler(NodeCall);

	//t_obj->SetAccessCheckCallbacks(Engine::GetNamedItem, 0, )
	//t_obj.MakeWeak(0, DestroyInstance);

    //NODE_SET_METHOD(t_obj, "test", toString);
    //NODE_SET_METHOD(t_obj, "valueOf", valueOf);
	
	target->Set(String::NewSymbol("Object"), t->GetFunction());
	Context::GetCurrent()->Global()->Set(String::NewSymbol("ActiveXObject"), t->GetFunction());
	NODE_DEBUG_MSG("DispObject initialized");
}

Handle<Value> DispObject::NodeCreate(IDispatch *ptr, LPCOLESTR name, DISPID id, LONG index)
{
	HandleScope scope;

	// Prepare disp item template
	if (tmplate.IsEmpty())
	{
		tmplate = Persistent<ObjectTemplate>::New(ObjectTemplate::New());
		tmplate->SetInternalFieldCount(1);
		tmplate->SetNamedPropertyHandler(NodeGet, NodeSet, 0, 0, 0, Handle<Value>());
		tmplate->SetIndexedPropertyHandler(NodeGetByIndex, NodeSetByIndex, 0, 0, 0, Handle<Value>());
		tmplate->SetCallAsFunctionHandler(NodeCall);
	}

	//
	Local<Object> &me = tmplate->NewInstance();
	(new DispObject(ptr, name, id, index))->Wrap(me);
	return scope.Close(me);
}

Handle<Value> DispObject::NodeCreate(const Arguments &args)
{
	HandleScope scope;

	// Prepare arguments
	Local<String> progid;
    if (args.Length() >= 1 && args[0]->IsString()) progid = args[0]->ToString();
	String::Value vname(progid);
	LPOLESTR id = (vname.length() > 0) ? (LPOLESTR)*vname : 0;

	// Initialize COM object
	CComPtr<IDispatch> disp;
	HRESULT hrcode = id ? disp.CoCreateInstance(id) : E_INVALIDARG;
	if FAILED(hrcode) return DispError(hrcode, L"CoCreateInstance");
	
	// Create object
	Local<Object> &me = args.This();
	(new DispObject(disp, id))->Wrap(me);
    return me;
}

Handle<Value> DispObject::NodeValueOf(const Arguments& args)
{
	DispObject *me = DispObject::Unwrap<DispObject>(args.This());
	if (!me) return args.This();
	Handle<Value> &val = me->valueOf();
	if (val == Undefined()) return args.This();
	return val;
}

Handle<Value> DispObject::NodeToString(const Arguments& args)
{
	DispObject *me = DispObject::Unwrap<DispObject>(args.This());
	if (!me) return Undefined();
	return me->toString();
}

Handle<Value> DispObject::NodeGet(Local<String> name, const AccessorInfo& info)
{
	HandleScope scope;
	String::Value vname(name);
	LPOLESTR id = (vname.length() > 0) ? (LPOLESTR)*vname : 0;
	DispObject *me = DispObject::Unwrap<DispObject>(info.This());
	if (!id || !me) return Undefined();
	NODE_DEBUG_FMT2("DispObject '%S.%S' get", me->name.GetString(), id);

	// Process standard elements
	if (_wcsicmp(id, L"valueOf") == 0) return scope.Close(FunctionTemplate::New(NodeValueOf, info.This())->GetFunction());
	if (_wcsicmp(id, L"toString") == 0) return scope.Close(FunctionTemplate::New(NodeToString, info.This())->GetFunction());

	// Other
	return scope.Close(me->findElement(id));
}

Handle<Value> DispObject::NodeGetByIndex(uint32_t index, const AccessorInfo& info)
{
	HandleScope scope;
	DispObject *me = DispObject::Unwrap<DispObject>(info.This());
	if (!me) return Undefined();
	NODE_DEBUG_FMT2("DispObject '%S[%u]' get", me->name.GetString(), index);
	return scope.Close(me->findElement(index));
}

Handle<Value> DispObject::NodeSet(Local<String> name, Local<Value> value, const AccessorInfo& info)
{
	HandleScope scope;
	String::Value vname(name);
	LPOLESTR id = (vname.length() > 0) ? (LPOLESTR)*vname : 0;
	DispObject *me = DispObject::Unwrap<DispObject>(info.This());
	if (!id || !me) return Undefined();
	NODE_DEBUG_FMT2("DispObject '%S.%S' set", me->name.GetString(), id);
	return scope.Close(me->set(id, value));
}

Handle<Value> DispObject::NodeSetByIndex(uint32_t index, Local<Value> value, const AccessorInfo& info)
{
	HandleScope scope;
	DispObject *me = DispObject::Unwrap<DispObject>(info.This());
	if (!me) return Undefined();
	NODE_DEBUG_FMT2("DispObject '%S[%u]' set", me->name.GetString(), index);
	return scope.Close(DispError(E_NOTIMPL, __FUNCTIONW__));
}

Handle<Value> DispObject::NodeCall(const Arguments &args)
{
	DispObject *me = DispObject::Unwrap<DispObject>(args.This());
	if (!me) return Undefined();
	NODE_DEBUG_FMT("DispObject '%S' call", me->name.GetString());
	return me->call(args);
}

//-------------------------------------------------------------------------------------------------------
