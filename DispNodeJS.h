//-------------------------------------------------------------------------------------------------------
// Project: NodeActiveX
// Author: Yuri Dursin
// Last Modification: 2011-11-20
// Description: DispObject class declarations. This class incapsulates COM IDispatch interface to Node JS Object
//-------------------------------------------------------------------------------------------------------

#pragma once

#include "DispNodeUtils.h"

class DispObject: public ObjectWrap
{
public:
	DispObject(IDispatch *ptr, LPCOLESTR name, DISPID id = DISPID_UNKNOWN, LONG indx = -1);
	~DispObject();

	static void NodeInit(Handle<Object> target);
	static Handle<Value> NodeCreate(IDispatch *ptr, LPCOLESTR name, DISPID id = DISPID_UNKNOWN, LONG indx = -1);

private:
	static Handle<Value> NodeCreate(const Arguments &args);
	static Handle<Value> NodeValueOf(const Arguments& args);
	static Handle<Value> NodeToString(const Arguments& args);
	static Handle<Value> NodeGet(Local<String> name, const AccessorInfo& info);
	static Handle<Value> NodeSet(Local<String> name, Local<Value> value, const AccessorInfo& info);
	static Handle<Value> NodeGetByIndex(uint32_t index, const AccessorInfo& info);
	static Handle<Value> NodeSetByIndex(uint32_t index, Local<Value> value, const AccessorInfo& info);
	static Handle<Value> NodeCall(const Arguments &args);

protected:
	virtual Handle<Value> findElement(LPOLESTR name);
	virtual Handle<Value> findElement(uint32_t index);
	virtual Handle<Value> valueOf();
	virtual Handle<Value> toString();
	virtual Handle<Value> set(LPOLESTR name, Local<Value> value);
	virtual Handle<Value> call(const Arguments &args);

private:
	static Persistent<ObjectTemplate> tmplate;

	enum State { None = 0, Prepared = 1, };
	int state;
	inline bool isPrepared() { return (state & Prepared) != 0; }
	inline bool isOwned() { return owner_disp != 0; }

	CComPtr<IDispatch> disp;
	CStringW name;
	LONG index;

	CComPtr<IDispatch> owner_disp;
	DISPID owner_id;

	HRESULT Prepare(VARIANT *value = 0);
};
