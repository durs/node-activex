//-------------------------------------------------------------------------------------------------------
// Project: NodeActiveX
// Author: Yuri Dursin
// Last Modification: 2017-04-16
// Description: DispObject class declarations. This class incapsulates COM IDispatch interface to Node JS Object
//-------------------------------------------------------------------------------------------------------

#pragma once

#include "utils.h"

class DispObject: public ObjectWrap
{
public:
	DispObject(IDispatch *ptr, LPCOLESTR name, DISPID id = DISPID_UNKNOWN, LONG indx = -1);
	~DispObject();

	static void NodeInit(Handle<Object> target);
	static Local<Object> NodeCreate(Isolate *isolate, IDispatch *ptr, LPCOLESTR name, DISPID id = DISPID_UNKNOWN, LONG indx = -1);

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
	virtual Local<Value> findElement(Isolate *isolate, LPOLESTR name);
	virtual Local<Value> findElement(Isolate *isolate, uint32_t index);
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
	inline bool isOwned() { return owner_disp != 0; }

	CComPtr<IDispatch> disp;
	std::wstring name;
	LONG index;

	CComPtr<IDispatch> owner_disp;
	DISPID owner_id;

	HRESULT Prepare(VARIANT *value = 0);
};
