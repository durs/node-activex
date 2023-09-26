//-------------------------------------------------------------------------------------------------------
// Project: NodeActiveX
// Author: Yuri Dursin
// Description:  Common utilities implementations
//-------------------------------------------------------------------------------------------------------

#include "stdafx.h"
#include "disp.h"
#include <OleCtl.h>
#include <oleacc.h>                  // for AccessibleObjectFromWindow
#pragma comment(lib, "OleAcc.lib")   // for AccessibleObjectFromWindow

const GUID CLSID_DispObjectImpl = { 0x9dce8520, 0x2efe, 0x48c0,{ 0xa0, 0xdc, 0x95, 0x1b, 0x29, 0x18, 0x72, 0xc0 } };

const IID IID_IReflect = { 0xAFBF15E5, 0xC37C, 0x11D2,{ 0xB8, 0x8E, 0x00, 0xA0, 0xC9, 0xB4, 0x71, 0xB8} };

//-------------------------------------------------------------------------------------------------------

#define ERROR_MESSAGE_WIDE_MAXSIZE 1024
#define ERROR_MESSAGE_UTF8_MAXSIZE 2048

void GetScodeString(HRESULT hr, wchar_t* buf, int bufSize)
{
	struct HRESULT_ENTRY
	{
		HRESULT hr;
		LPCWSTR lpszName;
	};
#define MAKE_HRESULT_ENTRY(hr)    { hr, L#hr }
	static const HRESULT_ENTRY hrNameTable[] =
	{
		MAKE_HRESULT_ENTRY(S_OK),
		MAKE_HRESULT_ENTRY(S_FALSE),

		MAKE_HRESULT_ENTRY(CACHE_S_FORMATETC_NOTSUPPORTED),
		MAKE_HRESULT_ENTRY(CACHE_S_SAMECACHE),
		MAKE_HRESULT_ENTRY(CACHE_S_SOMECACHES_NOTUPDATED),
		MAKE_HRESULT_ENTRY(CONVERT10_S_NO_PRESENTATION),
		MAKE_HRESULT_ENTRY(DATA_S_SAMEFORMATETC),
		MAKE_HRESULT_ENTRY(DRAGDROP_S_CANCEL),
		MAKE_HRESULT_ENTRY(DRAGDROP_S_DROP),
		MAKE_HRESULT_ENTRY(DRAGDROP_S_USEDEFAULTCURSORS),
		MAKE_HRESULT_ENTRY(INPLACE_S_TRUNCATED),
		MAKE_HRESULT_ENTRY(MK_S_HIM),
		MAKE_HRESULT_ENTRY(MK_S_ME),
		MAKE_HRESULT_ENTRY(MK_S_MONIKERALREADYREGISTERED),
		MAKE_HRESULT_ENTRY(MK_S_REDUCED_TO_SELF),
		MAKE_HRESULT_ENTRY(MK_S_US),
		MAKE_HRESULT_ENTRY(OLE_S_MAC_CLIPFORMAT),
		MAKE_HRESULT_ENTRY(OLE_S_STATIC),
		MAKE_HRESULT_ENTRY(OLE_S_USEREG),
		MAKE_HRESULT_ENTRY(OLEOBJ_S_CANNOT_DOVERB_NOW),
		MAKE_HRESULT_ENTRY(OLEOBJ_S_INVALIDHWND),
		MAKE_HRESULT_ENTRY(OLEOBJ_S_INVALIDVERB),
		MAKE_HRESULT_ENTRY(OLEOBJ_S_LAST),
		MAKE_HRESULT_ENTRY(STG_S_CONVERTED),
		MAKE_HRESULT_ENTRY(VIEW_S_ALREADY_FROZEN),

		MAKE_HRESULT_ENTRY(E_UNEXPECTED),
		MAKE_HRESULT_ENTRY(E_NOTIMPL),
		MAKE_HRESULT_ENTRY(E_OUTOFMEMORY),
		MAKE_HRESULT_ENTRY(E_INVALIDARG),
		MAKE_HRESULT_ENTRY(E_NOINTERFACE),
		MAKE_HRESULT_ENTRY(E_POINTER),
		MAKE_HRESULT_ENTRY(E_HANDLE),
		MAKE_HRESULT_ENTRY(E_ABORT),
		MAKE_HRESULT_ENTRY(E_FAIL),
		MAKE_HRESULT_ENTRY(E_ACCESSDENIED),

		MAKE_HRESULT_ENTRY(CACHE_E_NOCACHE_UPDATED),
		MAKE_HRESULT_ENTRY(CLASS_E_CLASSNOTAVAILABLE),
		MAKE_HRESULT_ENTRY(CLASS_E_NOAGGREGATION),
		MAKE_HRESULT_ENTRY(CLIPBRD_E_BAD_DATA),
		MAKE_HRESULT_ENTRY(CLIPBRD_E_CANT_CLOSE),
		MAKE_HRESULT_ENTRY(CLIPBRD_E_CANT_EMPTY),
		MAKE_HRESULT_ENTRY(CLIPBRD_E_CANT_OPEN),
		MAKE_HRESULT_ENTRY(CLIPBRD_E_CANT_SET),
		MAKE_HRESULT_ENTRY(CO_E_ALREADYINITIALIZED),
		MAKE_HRESULT_ENTRY(CO_E_APPDIDNTREG),
		MAKE_HRESULT_ENTRY(CO_E_APPNOTFOUND),
		MAKE_HRESULT_ENTRY(CO_E_APPSINGLEUSE),
		MAKE_HRESULT_ENTRY(CO_E_BAD_PATH),
		MAKE_HRESULT_ENTRY(CO_E_CANTDETERMINECLASS),
		MAKE_HRESULT_ENTRY(CO_E_CLASS_CREATE_FAILED),
		MAKE_HRESULT_ENTRY(CO_E_CLASSSTRING),
		MAKE_HRESULT_ENTRY(CO_E_DLLNOTFOUND),
		MAKE_HRESULT_ENTRY(CO_E_ERRORINAPP),
		MAKE_HRESULT_ENTRY(CO_E_ERRORINDLL),
		MAKE_HRESULT_ENTRY(CO_E_IIDSTRING),
		MAKE_HRESULT_ENTRY(CO_E_NOTINITIALIZED),
		MAKE_HRESULT_ENTRY(CO_E_OBJISREG),
		MAKE_HRESULT_ENTRY(CO_E_OBJNOTCONNECTED),
		MAKE_HRESULT_ENTRY(CO_E_OBJNOTREG),
		MAKE_HRESULT_ENTRY(CO_E_OBJSRV_RPC_FAILURE),
		MAKE_HRESULT_ENTRY(CO_E_SCM_ERROR),
		MAKE_HRESULT_ENTRY(CO_E_SCM_RPC_FAILURE),
		MAKE_HRESULT_ENTRY(CO_E_SERVER_EXEC_FAILURE),
		MAKE_HRESULT_ENTRY(CO_E_SERVER_STOPPING),
		MAKE_HRESULT_ENTRY(CO_E_WRONGOSFORAPP),
		MAKE_HRESULT_ENTRY(CONVERT10_E_OLESTREAM_BITMAP_TO_DIB),
		MAKE_HRESULT_ENTRY(CONVERT10_E_OLESTREAM_FMT),
		MAKE_HRESULT_ENTRY(CONVERT10_E_OLESTREAM_GET),
		MAKE_HRESULT_ENTRY(CONVERT10_E_OLESTREAM_PUT),
		MAKE_HRESULT_ENTRY(CONVERT10_E_STG_DIB_TO_BITMAP),
		MAKE_HRESULT_ENTRY(CONVERT10_E_STG_FMT),
		MAKE_HRESULT_ENTRY(CONVERT10_E_STG_NO_STD_STREAM),
		MAKE_HRESULT_ENTRY(DISP_E_ARRAYISLOCKED),
		MAKE_HRESULT_ENTRY(DISP_E_BADCALLEE),
		MAKE_HRESULT_ENTRY(DISP_E_BADINDEX),
		MAKE_HRESULT_ENTRY(DISP_E_BADPARAMCOUNT),
		MAKE_HRESULT_ENTRY(DISP_E_BADVARTYPE),
		MAKE_HRESULT_ENTRY(DISP_E_EXCEPTION),
		MAKE_HRESULT_ENTRY(DISP_E_MEMBERNOTFOUND),
		MAKE_HRESULT_ENTRY(DISP_E_NONAMEDARGS),
		MAKE_HRESULT_ENTRY(DISP_E_NOTACOLLECTION),
		MAKE_HRESULT_ENTRY(DISP_E_OVERFLOW),
		MAKE_HRESULT_ENTRY(DISP_E_PARAMNOTFOUND),
		MAKE_HRESULT_ENTRY(DISP_E_PARAMNOTOPTIONAL),
		MAKE_HRESULT_ENTRY(DISP_E_TYPEMISMATCH),
		MAKE_HRESULT_ENTRY(DISP_E_UNKNOWNINTERFACE),
		MAKE_HRESULT_ENTRY(DISP_E_UNKNOWNLCID),
		MAKE_HRESULT_ENTRY(DISP_E_UNKNOWNNAME),
		MAKE_HRESULT_ENTRY(DRAGDROP_E_ALREADYREGISTERED),
		MAKE_HRESULT_ENTRY(DRAGDROP_E_INVALIDHWND),
		MAKE_HRESULT_ENTRY(DRAGDROP_E_NOTREGISTERED),
		MAKE_HRESULT_ENTRY(DV_E_CLIPFORMAT),
		MAKE_HRESULT_ENTRY(DV_E_DVASPECT),
		MAKE_HRESULT_ENTRY(DV_E_DVTARGETDEVICE),
		MAKE_HRESULT_ENTRY(DV_E_DVTARGETDEVICE_SIZE),
		MAKE_HRESULT_ENTRY(DV_E_FORMATETC),
		MAKE_HRESULT_ENTRY(DV_E_LINDEX),
		MAKE_HRESULT_ENTRY(DV_E_NOIVIEWOBJECT),
		MAKE_HRESULT_ENTRY(DV_E_STATDATA),
		MAKE_HRESULT_ENTRY(DV_E_STGMEDIUM),
		MAKE_HRESULT_ENTRY(DV_E_TYMED),
		MAKE_HRESULT_ENTRY(INPLACE_E_NOTOOLSPACE),
		MAKE_HRESULT_ENTRY(INPLACE_E_NOTUNDOABLE),
		MAKE_HRESULT_ENTRY(MEM_E_INVALID_LINK),
		MAKE_HRESULT_ENTRY(MEM_E_INVALID_ROOT),
		MAKE_HRESULT_ENTRY(MEM_E_INVALID_SIZE),
		MAKE_HRESULT_ENTRY(MK_E_CANTOPENFILE),
		MAKE_HRESULT_ENTRY(MK_E_CONNECTMANUALLY),
		MAKE_HRESULT_ENTRY(MK_E_ENUMERATION_FAILED),
		MAKE_HRESULT_ENTRY(MK_E_EXCEEDEDDEADLINE),
		MAKE_HRESULT_ENTRY(MK_E_INTERMEDIATEINTERFACENOTSUPPORTED),
		MAKE_HRESULT_ENTRY(MK_E_INVALIDEXTENSION),
		MAKE_HRESULT_ENTRY(MK_E_MUSTBOTHERUSER),
		MAKE_HRESULT_ENTRY(MK_E_NEEDGENERIC),
		MAKE_HRESULT_ENTRY(MK_E_NO_NORMALIZED),
		MAKE_HRESULT_ENTRY(MK_E_NOINVERSE),
		MAKE_HRESULT_ENTRY(MK_E_NOOBJECT),
		MAKE_HRESULT_ENTRY(MK_E_NOPREFIX),
		MAKE_HRESULT_ENTRY(MK_E_NOSTORAGE),
		MAKE_HRESULT_ENTRY(MK_E_NOTBINDABLE),
		MAKE_HRESULT_ENTRY(MK_E_NOTBOUND),
		MAKE_HRESULT_ENTRY(MK_E_SYNTAX),
		MAKE_HRESULT_ENTRY(MK_E_UNAVAILABLE),
		MAKE_HRESULT_ENTRY(OLE_E_ADVF),
		MAKE_HRESULT_ENTRY(OLE_E_ADVISENOTSUPPORTED),
		MAKE_HRESULT_ENTRY(OLE_E_BLANK),
		MAKE_HRESULT_ENTRY(OLE_E_CANT_BINDTOSOURCE),
		MAKE_HRESULT_ENTRY(OLE_E_CANT_GETMONIKER),
		MAKE_HRESULT_ENTRY(OLE_E_CANTCONVERT),
		MAKE_HRESULT_ENTRY(OLE_E_CLASSDIFF),
		MAKE_HRESULT_ENTRY(OLE_E_ENUM_NOMORE),
		MAKE_HRESULT_ENTRY(OLE_E_INVALIDHWND),
		MAKE_HRESULT_ENTRY(OLE_E_INVALIDRECT),
		MAKE_HRESULT_ENTRY(OLE_E_NOCACHE),
		MAKE_HRESULT_ENTRY(OLE_E_NOCONNECTION),
		MAKE_HRESULT_ENTRY(OLE_E_NOSTORAGE),
		MAKE_HRESULT_ENTRY(OLE_E_NOT_INPLACEACTIVE),
		MAKE_HRESULT_ENTRY(OLE_E_NOTRUNNING),
		MAKE_HRESULT_ENTRY(OLE_E_OLEVERB),
		MAKE_HRESULT_ENTRY(OLE_E_PROMPTSAVECANCELLED),
		MAKE_HRESULT_ENTRY(OLE_E_STATIC),
		MAKE_HRESULT_ENTRY(OLE_E_WRONGCOMPOBJ),
		MAKE_HRESULT_ENTRY(OLEOBJ_E_INVALIDVERB),
		MAKE_HRESULT_ENTRY(OLEOBJ_E_NOVERBS),
		MAKE_HRESULT_ENTRY(REGDB_E_CLASSNOTREG),
		MAKE_HRESULT_ENTRY(REGDB_E_IIDNOTREG),
		MAKE_HRESULT_ENTRY(REGDB_E_INVALIDVALUE),
		MAKE_HRESULT_ENTRY(REGDB_E_KEYMISSING),
		MAKE_HRESULT_ENTRY(REGDB_E_READREGDB),
		MAKE_HRESULT_ENTRY(REGDB_E_WRITEREGDB),
		MAKE_HRESULT_ENTRY(RPC_E_ATTEMPTED_MULTITHREAD),
		MAKE_HRESULT_ENTRY(RPC_E_CALL_CANCELED),
		MAKE_HRESULT_ENTRY(RPC_E_CALL_REJECTED),
		MAKE_HRESULT_ENTRY(RPC_E_CANTCALLOUT_AGAIN),
		MAKE_HRESULT_ENTRY(RPC_E_CANTCALLOUT_INASYNCCALL),
		MAKE_HRESULT_ENTRY(RPC_E_CANTCALLOUT_INEXTERNALCALL),
		MAKE_HRESULT_ENTRY(RPC_E_CANTCALLOUT_ININPUTSYNCCALL),
		MAKE_HRESULT_ENTRY(RPC_E_CANTPOST_INSENDCALL),
		MAKE_HRESULT_ENTRY(RPC_E_CANTTRANSMIT_CALL),
		MAKE_HRESULT_ENTRY(RPC_E_CHANGED_MODE),
		MAKE_HRESULT_ENTRY(RPC_E_CLIENT_CANTMARSHAL_DATA),
		MAKE_HRESULT_ENTRY(RPC_E_CLIENT_CANTUNMARSHAL_DATA),
		MAKE_HRESULT_ENTRY(RPC_E_CLIENT_DIED),
		MAKE_HRESULT_ENTRY(RPC_E_CONNECTION_TERMINATED),
		MAKE_HRESULT_ENTRY(RPC_E_DISCONNECTED),
		MAKE_HRESULT_ENTRY(RPC_E_FAULT),
		MAKE_HRESULT_ENTRY(RPC_E_INVALID_CALLDATA),
		MAKE_HRESULT_ENTRY(RPC_E_INVALID_DATA),
		MAKE_HRESULT_ENTRY(RPC_E_INVALID_DATAPACKET),
		MAKE_HRESULT_ENTRY(RPC_E_INVALID_PARAMETER),
		MAKE_HRESULT_ENTRY(RPC_E_INVALIDMETHOD),
		MAKE_HRESULT_ENTRY(RPC_E_NOT_REGISTERED),
		MAKE_HRESULT_ENTRY(RPC_E_OUT_OF_RESOURCES),
		MAKE_HRESULT_ENTRY(RPC_E_RETRY),
		MAKE_HRESULT_ENTRY(RPC_E_SERVER_CANTMARSHAL_DATA),
		MAKE_HRESULT_ENTRY(RPC_E_SERVER_CANTUNMARSHAL_DATA),
		MAKE_HRESULT_ENTRY(RPC_E_SERVER_DIED),
		MAKE_HRESULT_ENTRY(RPC_E_SERVER_DIED_DNE),
		MAKE_HRESULT_ENTRY(RPC_E_SERVERCALL_REJECTED),
		MAKE_HRESULT_ENTRY(RPC_E_SERVERCALL_RETRYLATER),
		MAKE_HRESULT_ENTRY(RPC_E_SERVERFAULT),
		MAKE_HRESULT_ENTRY(RPC_E_SYS_CALL_FAILED),
		MAKE_HRESULT_ENTRY(RPC_E_THREAD_NOT_INIT),
		MAKE_HRESULT_ENTRY(RPC_E_UNEXPECTED),
		MAKE_HRESULT_ENTRY(RPC_E_WRONG_THREAD),
		MAKE_HRESULT_ENTRY(STG_E_ABNORMALAPIEXIT),
		MAKE_HRESULT_ENTRY(STG_E_ACCESSDENIED),
		MAKE_HRESULT_ENTRY(STG_E_CANTSAVE),
		MAKE_HRESULT_ENTRY(STG_E_DISKISWRITEPROTECTED),
		MAKE_HRESULT_ENTRY(STG_E_EXTANTMARSHALLINGS),
		MAKE_HRESULT_ENTRY(STG_E_FILEALREADYEXISTS),
		MAKE_HRESULT_ENTRY(STG_E_FILENOTFOUND),
		MAKE_HRESULT_ENTRY(STG_E_INSUFFICIENTMEMORY),
		MAKE_HRESULT_ENTRY(STG_E_INUSE),
		MAKE_HRESULT_ENTRY(STG_E_INVALIDFLAG),
		MAKE_HRESULT_ENTRY(STG_E_INVALIDFUNCTION),
		MAKE_HRESULT_ENTRY(STG_E_INVALIDHANDLE),
		MAKE_HRESULT_ENTRY(STG_E_INVALIDHEADER),
		MAKE_HRESULT_ENTRY(STG_E_INVALIDNAME),
		MAKE_HRESULT_ENTRY(STG_E_INVALIDPARAMETER),
		MAKE_HRESULT_ENTRY(STG_E_INVALIDPOINTER),
		MAKE_HRESULT_ENTRY(STG_E_LOCKVIOLATION),
		MAKE_HRESULT_ENTRY(STG_E_MEDIUMFULL),
		MAKE_HRESULT_ENTRY(STG_E_NOMOREFILES),
		MAKE_HRESULT_ENTRY(STG_E_NOTCURRENT),
		MAKE_HRESULT_ENTRY(STG_E_NOTFILEBASEDSTORAGE),
		MAKE_HRESULT_ENTRY(STG_E_OLDDLL),
		MAKE_HRESULT_ENTRY(STG_E_OLDFORMAT),
		MAKE_HRESULT_ENTRY(STG_E_PATHNOTFOUND),
		MAKE_HRESULT_ENTRY(STG_E_READFAULT),
		MAKE_HRESULT_ENTRY(STG_E_REVERTED),
		MAKE_HRESULT_ENTRY(STG_E_SEEKERROR),
		MAKE_HRESULT_ENTRY(STG_E_SHAREREQUIRED),
		MAKE_HRESULT_ENTRY(STG_E_SHAREVIOLATION),
		MAKE_HRESULT_ENTRY(STG_E_TOOMANYOPENFILES),
		MAKE_HRESULT_ENTRY(STG_E_UNIMPLEMENTEDFUNCTION),
		MAKE_HRESULT_ENTRY(STG_E_UNKNOWN),
		MAKE_HRESULT_ENTRY(STG_E_WRITEFAULT),
		MAKE_HRESULT_ENTRY(TYPE_E_AMBIGUOUSNAME),
		MAKE_HRESULT_ENTRY(TYPE_E_BADMODULEKIND),
		MAKE_HRESULT_ENTRY(TYPE_E_BUFFERTOOSMALL),
		MAKE_HRESULT_ENTRY(TYPE_E_CANTCREATETMPFILE),
		MAKE_HRESULT_ENTRY(TYPE_E_CANTLOADLIBRARY),
		MAKE_HRESULT_ENTRY(TYPE_E_CIRCULARTYPE),
		MAKE_HRESULT_ENTRY(TYPE_E_DLLFUNCTIONNOTFOUND),
		MAKE_HRESULT_ENTRY(TYPE_E_DUPLICATEID),
		MAKE_HRESULT_ENTRY(TYPE_E_ELEMENTNOTFOUND),
		MAKE_HRESULT_ENTRY(TYPE_E_INCONSISTENTPROPFUNCS),
		MAKE_HRESULT_ENTRY(TYPE_E_INVALIDSTATE),
		MAKE_HRESULT_ENTRY(TYPE_E_INVDATAREAD),
		MAKE_HRESULT_ENTRY(TYPE_E_IOERROR),
		MAKE_HRESULT_ENTRY(TYPE_E_LIBNOTREGISTERED),
		MAKE_HRESULT_ENTRY(TYPE_E_NAMECONFLICT),
		MAKE_HRESULT_ENTRY(TYPE_E_OUTOFBOUNDS),
		MAKE_HRESULT_ENTRY(TYPE_E_QUALIFIEDNAMEDISALLOWED),
		MAKE_HRESULT_ENTRY(TYPE_E_REGISTRYACCESS),
		MAKE_HRESULT_ENTRY(TYPE_E_SIZETOOBIG),
		MAKE_HRESULT_ENTRY(TYPE_E_TYPEMISMATCH),
		MAKE_HRESULT_ENTRY(TYPE_E_UNDEFINEDTYPE),
		MAKE_HRESULT_ENTRY(TYPE_E_UNKNOWNLCID),
		MAKE_HRESULT_ENTRY(TYPE_E_UNSUPFORMAT),
		MAKE_HRESULT_ENTRY(TYPE_E_WRONGTYPEKIND),
		MAKE_HRESULT_ENTRY(VIEW_E_DRAW),

		MAKE_HRESULT_ENTRY(CONNECT_E_NOCONNECTION),
		MAKE_HRESULT_ENTRY(CONNECT_E_ADVISELIMIT),
		MAKE_HRESULT_ENTRY(CONNECT_E_CANNOTCONNECT),
		MAKE_HRESULT_ENTRY(CONNECT_E_OVERRIDDEN),

		MAKE_HRESULT_ENTRY(CLASS_E_NOTLICENSED),
		MAKE_HRESULT_ENTRY(CLASS_E_NOAGGREGATION),
		MAKE_HRESULT_ENTRY(CLASS_E_CLASSNOTAVAILABLE),

		MAKE_HRESULT_ENTRY(CTL_E_ILLEGALFUNCTIONCALL),
		MAKE_HRESULT_ENTRY(CTL_E_OVERFLOW),
		MAKE_HRESULT_ENTRY(CTL_E_OUTOFMEMORY),
		MAKE_HRESULT_ENTRY(CTL_E_DIVISIONBYZERO),
		MAKE_HRESULT_ENTRY(CTL_E_OUTOFSTRINGSPACE),
		MAKE_HRESULT_ENTRY(CTL_E_OUTOFSTACKSPACE),
		MAKE_HRESULT_ENTRY(CTL_E_BADFILENAMEORNUMBER),
		MAKE_HRESULT_ENTRY(CTL_E_FILENOTFOUND),
		MAKE_HRESULT_ENTRY(CTL_E_BADFILEMODE),
		MAKE_HRESULT_ENTRY(CTL_E_FILEALREADYOPEN),
		MAKE_HRESULT_ENTRY(CTL_E_DEVICEIOERROR),
		MAKE_HRESULT_ENTRY(CTL_E_FILEALREADYEXISTS),
		MAKE_HRESULT_ENTRY(CTL_E_BADRECORDLENGTH),
		MAKE_HRESULT_ENTRY(CTL_E_DISKFULL),
		MAKE_HRESULT_ENTRY(CTL_E_BADRECORDNUMBER),
		MAKE_HRESULT_ENTRY(CTL_E_BADFILENAME),
		MAKE_HRESULT_ENTRY(CTL_E_TOOMANYFILES),
		MAKE_HRESULT_ENTRY(CTL_E_DEVICEUNAVAILABLE),
		MAKE_HRESULT_ENTRY(CTL_E_PERMISSIONDENIED),
		MAKE_HRESULT_ENTRY(CTL_E_DISKNOTREADY),
		MAKE_HRESULT_ENTRY(CTL_E_PATHFILEACCESSERROR),
		MAKE_HRESULT_ENTRY(CTL_E_PATHNOTFOUND),
		MAKE_HRESULT_ENTRY(CTL_E_INVALIDPATTERNSTRING),
		MAKE_HRESULT_ENTRY(CTL_E_INVALIDUSEOFNULL),
		MAKE_HRESULT_ENTRY(CTL_E_INVALIDFILEFORMAT),
		MAKE_HRESULT_ENTRY(CTL_E_INVALIDPROPERTYVALUE),
		MAKE_HRESULT_ENTRY(CTL_E_INVALIDPROPERTYARRAYINDEX),
		MAKE_HRESULT_ENTRY(CTL_E_SETNOTSUPPORTEDATRUNTIME),
		MAKE_HRESULT_ENTRY(CTL_E_SETNOTSUPPORTED),
		MAKE_HRESULT_ENTRY(CTL_E_NEEDPROPERTYARRAYINDEX),
		MAKE_HRESULT_ENTRY(CTL_E_SETNOTPERMITTED),
		MAKE_HRESULT_ENTRY(CTL_E_GETNOTSUPPORTEDATRUNTIME),
		MAKE_HRESULT_ENTRY(CTL_E_GETNOTSUPPORTED),
		MAKE_HRESULT_ENTRY(CTL_E_PROPERTYNOTFOUND),
		MAKE_HRESULT_ENTRY(CTL_E_INVALIDCLIPBOARDFORMAT),
		MAKE_HRESULT_ENTRY(CTL_E_INVALIDPICTURE),
		MAKE_HRESULT_ENTRY(CTL_E_PRINTERERROR),
		MAKE_HRESULT_ENTRY(CTL_E_CANTSAVEFILETOTEMP),
		MAKE_HRESULT_ENTRY(CTL_E_SEARCHTEXTNOTFOUND),
		MAKE_HRESULT_ENTRY(CTL_E_REPLACEMENTSTOOLONG),
	};

	// look for scode in the table
	for (int i = 0; i < sizeof(hrNameTable) / sizeof(hrNameTable[0]); i++)
	{
		if (hr == hrNameTable[i].hr) {
			swprintf_s(buf, (size_t)(bufSize - 1), hrNameTable[i].lpszName);
			return;
		}
	}
	// not found - make one up
	swprintf_s(buf, (size_t)(bufSize - 1), L"OLE error 0x%08x", hr);
}


uint16_t *GetWin32ErrorMessage(uint16_t *buf, size_t buflen, Isolate *isolate, HRESULT hrcode, LPCOLESTR msg, LPCOLESTR msg2, LPCOLESTR desc) {
	uint16_t *bufptr = buf;
	size_t len;
	if (msg) {
		len = wcslen(msg);
		if (len >= buflen) len = buflen - 1;
		if (len > 0) memcpy(bufptr, msg, len * sizeof(uint16_t));
		buflen -= len;
		bufptr += len;
		if (buflen > 2) {
			bufptr[0] = ':';
			bufptr[1] = ' ';
			buflen -= 2;
			bufptr += 2;
		}
	}
	if (msg2) {
		len = wcslen(msg2);
		if (len >= buflen) len = buflen - 1;
		if (len > 0) memcpy(bufptr, msg2, len * sizeof(uint16_t));
		buflen -= len;
		bufptr += len;
		if (buflen > 1) {
			bufptr[0] = ' ';
			buflen -= 1;
			bufptr += 1;
		}
	}
	if (buflen > 1) {
		len = desc ? wcslen(desc) : 0;
		if (len > 0) {
			if (len >= buflen) len = buflen - 1;
			memcpy(bufptr, desc, len * sizeof(OLECHAR));
		}
		else {
			len = FormatMessageW(FORMAT_MESSAGE_FROM_SYSTEM, 0, hrcode, MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), (LPOLESTR)bufptr, (DWORD)buflen - 1, 0);
			if (len == 0) len = swprintf_s((LPOLESTR)bufptr, buflen - 1, L"Error 0x%08X", hrcode);
		}
		buflen -= len;
		bufptr += len;
	}
	if (buflen > 0) bufptr[0] = 0;
	return buf;
}

char *GetWin32ErrorMessage(char *buf, size_t buflen, Isolate *isolate, HRESULT hrcode, LPCOLESTR msg, LPCOLESTR msg2, LPCOLESTR desc) {
	uint16_t buf_wide[ERROR_MESSAGE_WIDE_MAXSIZE];
	GetWin32ErrorMessage(buf_wide, ERROR_MESSAGE_WIDE_MAXSIZE, isolate, hrcode, msg, msg2, desc);
	int rcode = WideCharToMultiByte(CP_UTF8, 0, (WCHAR*)buf_wide, -1, buf, buflen, NULL, NULL);
	if (rcode < 0) rcode = 0;
	buf[rcode] = 0;
	return buf;
}

Local<String> GetWin32ErrorMessage(Isolate *isolate, HRESULT hrcode, LPCOLESTR msg, LPCOLESTR msg2, LPCOLESTR desc) {
	uint16_t buf_wide[ERROR_MESSAGE_WIDE_MAXSIZE];
	return v8str(isolate, GetWin32ErrorMessage(buf_wide, ERROR_MESSAGE_WIDE_MAXSIZE, isolate, hrcode, msg, msg2, desc));
}

//-------------------------------------------------------------------------------------------------------

Local<Value> Variant2Array(Isolate *isolate, const VARIANT &v) {
	if ((v.vt & VT_ARRAY) == 0) return Null(isolate);
	SAFEARRAY *varr = (v.vt & VT_BYREF) != 0 ? *v.pparray : v.parray;
	if (!varr || varr->cDims > 2 || varr->cDims == 0) return Null(isolate);
	else if ( varr->cDims == 2 ) return Variant2Array2( isolate, v );
	Local<Context> ctx = isolate->GetCurrentContext();
	VARTYPE vt = v.vt & VT_TYPEMASK;
	LONG cnt = (LONG)varr->rgsabound[0].cElements;
	Local<Array> arr = Array::New(isolate, cnt);
	for (LONG i = varr->rgsabound[0].lLbound; i < varr->rgsabound[0].lLbound + cnt; i++) {
		CComVariant vi;
		if SUCCEEDED(SafeArrayGetElement(varr, &i, (vt == VT_VARIANT) ? (void*)&vi : (void*)&vi.byref)) {
			if (vt != VT_VARIANT) vi.vt = vt;
			uint32_t jsi = i - varr->rgsabound[ 0 ].lLbound;
			arr->Set(ctx, jsi, Variant2Value(isolate, vi, true));
		}
	}
	return arr;
}

static Local<Value> Variant2Array2(Isolate *isolate, const VARIANT &v) {
	if ((v.vt & VT_ARRAY) == 0) return Null(isolate);
	SAFEARRAY *varr = (v.vt & VT_BYREF) != 0 ? *v.pparray : v.parray;
	if (!varr || varr->cDims != 2) return Null(isolate);
	Local<Context> ctx = isolate->GetCurrentContext();
	VARTYPE vt = v.vt & VT_TYPEMASK;
	LONG cnt1 = (LONG)varr->rgsabound[0].cElements;
	LONG cnt2 = (LONG)varr->rgsabound[1].cElements;
	Local<Array> arr1 = Array::New(isolate, cnt2);
	LONG rgIndices[ 2 ];
	for (LONG i2 = varr->rgsabound[1].lLbound; i2 < varr->rgsabound[1].lLbound + cnt2; i2++) {
		rgIndices[ 0 ] = i2;
		Local<Array> arr2 = Array::New(isolate, cnt1);
		for (LONG i1 = varr->rgsabound[0].lLbound; i1 < varr->rgsabound[0].lLbound + cnt1; i1++) {
			CComVariant vi;
			rgIndices[ 1 ] = i1;
			if SUCCEEDED(SafeArrayGetElement(varr, &rgIndices[0], (vt == VT_VARIANT) ? (void*)&vi : (void*)&vi.byref)) {
				if (vt != VT_VARIANT) vi.vt = vt;
				uint32_t jsi = (uint32_t)i1 - varr->rgsabound[ 0 ].lLbound;
				arr2->Set(ctx, jsi, Variant2Value(isolate, vi, true));
			}
		}
		uint32_t jsi = i2 - varr->rgsabound[ 1 ].lLbound;
		arr1->Set(ctx, jsi, arr2);
	}
	return arr1;
}

Local<Value> Variant2Value(Isolate *isolate, const VARIANT &v, bool allow_disp) {
	if ((v.vt & VT_ARRAY) != 0) return Variant2Array(isolate, v);
	VARTYPE vt = (v.vt & VT_TYPEMASK);
	bool by_ref = (v.vt & VT_BYREF) != 0;
	switch (vt) {
	case VT_NULL:
		return Null(isolate);
	case VT_I1:
		return Int32::New(isolate, (int32_t)(by_ref ? *v.pcVal : v.cVal));
	case VT_I2:
		return Int32::New(isolate, (int32_t)(by_ref ? *v.piVal : v.iVal));
	case VT_I4:
		return Int32::New(isolate, (int32_t)(by_ref ? *v.plVal : v.lVal));
	case VT_INT:
		return Int32::New(isolate, (int32_t)(by_ref ? *v.pintVal : v.intVal));
	case VT_UI1:
		return Int32::New(isolate, (uint32_t)(by_ref ? *v.pbVal : v.bVal));
	case VT_UI2:
		return Int32::New(isolate, (uint32_t)(by_ref ? *v.puiVal : v.uiVal));
	case VT_UI4:
		return Int32::New(isolate, (uint32_t)(by_ref ? *v.pulVal : v.ulVal));
	case VT_UINT:
		return Int32::New(isolate, (uint32_t)(by_ref ? *v.puintVal : v.uintVal));
	case VT_I8:
		return Number::New(isolate, (double)(by_ref ? *v.pllVal : v.llVal));
	case VT_UI8:
		return Number::New(isolate, (double)(by_ref ? *v.pullVal : v.ullVal));
	case VT_CY:
		return Number::New(isolate, (double)(by_ref ? v.pcyVal : &v.cyVal)->int64 / 10000.);
	case VT_R4:
		return Number::New(isolate, by_ref ? *v.pfltVal : v.fltVal);
	case VT_R8:
		return Number::New(isolate, by_ref ? *v.pdblVal : v.dblVal);
	case VT_DATE: {
		Local<Value> ret;
		if (!Date::New(isolate->GetCurrentContext(), FromOleDate(by_ref ? *v.pdate : v.date)).ToLocal(&ret))
			ret = Undefined(isolate);
		return ret;
	}
	case VT_DECIMAL: {
		DOUBLE dblval;
		if FAILED(VarR8FromDec(by_ref ? v.pdecVal : &v.decVal, &dblval)) return Undefined(isolate);
		return Number::New(isolate, dblval);		
	}
	case VT_BOOL:
		return Boolean::New(isolate, (by_ref ? *v.pboolVal : v.boolVal) != VARIANT_FALSE);
	case VT_DISPATCH: {
		IDispatch *disp = (by_ref ? *v.ppdispVal : v.pdispVal);
		if (!disp) return Null(isolate);
		if (allow_disp) {
			DispObjectImpl *impl;
			if (disp->QueryInterface(CLSID_DispObjectImpl, (void**)&impl) == S_OK) {
				return impl->obj.Get(isolate);
			}
			return DispObject::NodeCreate(isolate, disp, L"Dispatch", option_auto);
		}
		return v8str(isolate, "[Dispatch]");
	}
	case VT_UNKNOWN: {
		if (allow_disp)
		{
			CComPtr<IDispatch> disp;
			if (UnknownDispGet(by_ref ? *v.ppunkVal : v.punkVal, &disp)) {
				return DispObject::NodeCreate(isolate, disp, L"Unknown", option_auto);
			}
		}
		// Check for NULL value
		if ((by_ref && *v.ppunkVal) || v.punkVal)
		{
			if (allow_disp) {
				return VariantObject::NodeCreate(isolate, v);
			}
			else {
				return v8str(isolate, "[Unknown]");
			}
		}
		else if (!v.punkVal) {
			return Null(isolate);
		}
	}
	case VT_BSTR: {
        BSTR bstr = by_ref ? (v.pbstrVal ? *v.pbstrVal : nullptr) : v.bstrVal;
        //if (!bstr) return String::Empty(isolate);
		// Sometimes we need to distinguish between NULL and empty string
		if (!bstr) return Null(isolate);
		return v8str(isolate, bstr);
    }
	case VT_VARIANT: 
		if (v.pvarVal) return Variant2Value(isolate, *v.pvarVal, allow_disp);
	}
	return Undefined(isolate);
}

Local<Value> Variant2String(Isolate *isolate, const VARIANT &v) {
	char buf[256] = {};
	VARTYPE vt = (v.vt & VT_TYPEMASK);
	bool by_ref = (v.vt & VT_BYREF) != 0;
	switch (vt) {
	case VT_EMPTY:
		strcpy(buf, "EMPTY");
		break;
	case VT_NULL:
		strcpy(buf, "NULL");
		break;
	case VT_I1:
		sprintf_s(buf, "%i", (int)(by_ref ? *v.pcVal : v.cVal));
		break;
	case VT_I2:
		sprintf_s(buf, "%i", (int)(by_ref ? *v.piVal : v.iVal));
		break;
	case VT_I4:
		sprintf_s(buf, "%i", (int)(by_ref ? *v.plVal : v.lVal));
		break;
	case VT_INT:
		sprintf_s(buf, "%i", (int)(by_ref ? *v.pintVal : v.intVal));
		break;
	case VT_UI1:
		sprintf_s(buf, "%u", (unsigned int)(by_ref ? *v.pbVal : v.bVal));
		break;
	case VT_UI2:
		sprintf_s(buf, "%u", (unsigned int)(by_ref ? *v.puiVal : v.uiVal));
		break;
	case VT_UI4:
		sprintf_s(buf, "%u", (unsigned int)(by_ref ? *v.pulVal : v.ulVal));
		break;
	case VT_UINT:
		sprintf_s(buf, "%u", (unsigned int)(by_ref ? *v.puintVal : v.uintVal));
		break;
	case VT_CY:
	case VT_I8:
		sprintf_s(buf, "%lld", (by_ref ? *v.pllVal : v.llVal));
		break;
	case VT_UI8:
		sprintf_s(buf, "%llu", (by_ref ? *v.pullVal : v.ullVal));
		break;
	case VT_R4:
		sprintf_s(buf, "%f", (double)(by_ref ? *v.pfltVal : v.fltVal));
		break;
	case VT_R8:
		sprintf_s(buf, "%f", (double)(by_ref ? *v.pdblVal : v.dblVal));
		break;
	case VT_DATE: {
		Local<Value> ret;
		if (!Date::New(isolate->GetCurrentContext(), FromOleDate(by_ref ? *v.pdate : v.date)).ToLocal(&ret))
			ret = Undefined(isolate);
		return ret;
	}
	case VT_DECIMAL: {
		DOUBLE dblval;
		if FAILED(VarR8FromDec(by_ref ? v.pdecVal : &v.decVal, &dblval)) return Undefined(isolate); 
		sprintf_s(buf, "%f", (double)dblval);
		break;		
	}
	case VT_BOOL:
		strcpy(buf, ((by_ref ? *v.pboolVal : v.boolVal) == VARIANT_FALSE) ? "false" : "true");
	case VT_DISPATCH:
		strcpy(buf, "[Dispatch]");
		break;
	case VT_UNKNOWN: 
		strcpy(buf, "[Unknown]");
		break;
	case VT_VARIANT:
		if (v.pvarVal) return Variant2String(isolate, *v.pvarVal);
		break;
	default:
		CComVariant tmp;
		if (SUCCEEDED(VariantChangeType(&tmp, &v, 0, VT_BSTR)) && tmp.vt == VT_BSTR && v.bstrVal != nullptr) {
			return v8str(isolate, v.bstrVal);
		}
	}
	return v8str(isolate, buf);
}

void Value2Variant(Isolate *isolate, Local<Value> &val, VARIANT &var, VARTYPE vt) {
	Local<Context> ctx = isolate->GetCurrentContext();
	if (val.IsEmpty() || val->IsUndefined()) {
		var.vt = VT_EMPTY;
	}
	else if (val->IsNull()) {
		var.vt = VT_NULL;
	}
	else if (val->IsInt32()) {
		//var.lVal = val->Int32Value();
        var.lVal = val->Int32Value(ctx).FromMaybe(0);
        var.vt = VT_I4;
    }
	else if (val->IsUint32()) {
		//var.ulVal = val->Uint32Value();
        var.ulVal = val->Uint32Value(ctx).FromMaybe(0);
		var.vt = (var.ulVal <= 0x7FFFFFFF) ? VT_I4 : VT_UI4;
	}
	else if (val->IsNumber()) {
		//var.dblVal = val->NumberValue();
        var.dblVal = val->NumberValue(ctx).FromMaybe(0);
        var.vt = VT_R8;
    }
	else if (val->IsDate()) {
		//var.date = ToOleDate(val->NumberValue());
        var.date = ToOleDate(val->NumberValue(ctx).FromMaybe(0));
        var.vt = VT_DATE;
    }
	else if (val->IsBoolean()) {
		var.boolVal = NODE_BOOL(isolate, val) ? VARIANT_TRUE : VARIANT_FALSE;
        var.vt = VT_BOOL;
    }
	else if (val->IsArray() && (vt != VT_NULL)) {
		Local<Array> arr = v8::Local<Array>::Cast(val);
		uint32_t len = arr->Length();
		if (vt == VT_EMPTY) vt = VT_VARIANT;
		var.vt = VT_ARRAY | vt;
		// if array of arrays, create a 2 dim array, choose the 2nd bound
		uint32_t second_len = 0;
		if (len) {
			Local<Value> first_value;
			if (arr->Get(ctx, 0).ToLocal(&first_value)) {
				if (first_value->IsArray()) second_len = v8::Local<Array>::Cast(first_value)->Length();
			}
		}
		if ( second_len == 0 ) {
			var.parray = SafeArrayCreateVector(vt, 0, len);
			for (uint32_t i = 0; i < len; i++) {
				CComVariant v;
				Local<Value> val;
				if (!arr->Get(ctx, i).ToLocal(&val)) val = Undefined(isolate);
				Value2Variant(isolate, val, v, vt);
				void *pv;
				if (vt == VT_VARIANT) pv = (void*)&v;
				else if (vt == VT_DISPATCH || vt == VT_UNKNOWN || vt == VT_BSTR) pv = v.byref;
				else pv = (void*)&v.byref;
				SafeArrayPutElement(var.parray, (LONG*)&i, pv);
			}
		}
		else {
			SAFEARRAYBOUND rgsabounds[ 2 ];
			rgsabounds[ 0 ].lLbound = rgsabounds[ 1 ].lLbound = 0;
			rgsabounds[ 0 ].cElements = len;
			rgsabounds[ 1 ].cElements = second_len;
			var.parray = SafeArrayCreate(vt, 2, rgsabounds);
			LONG rgIndices[ 2 ];
			for (uint32_t i = 0; i < len; i++) {
				rgIndices[ 0 ] = i;
				Local<Value> maybearray;
				Local<Array> arr2;
				bool bGotArray = false;
				if (arr->Get( ctx, i ).ToLocal( &maybearray ) && maybearray->IsArray()) {
					bGotArray = true;
					arr2 = v8::Local<Array>::Cast(maybearray);
				}
				for (uint32_t j = 0; j < second_len; j++) {
					rgIndices[ 1 ] = j;
					Local<Value> val;
					if ( bGotArray ) {
						if (!arr2->Get(ctx, j).ToLocal(&val)) val = Undefined(isolate);
					}
					else {
						// if no arrays for a "row", the value is put only in first "col"
						if ( j == 0 ) val = maybearray;
						else val = Undefined(isolate);
					}
					CComVariant v;
					Value2Variant(isolate, val, v, vt);
					void *pv;
					if (vt == VT_VARIANT) pv = (void*)&v;
					else if (vt == VT_DISPATCH || vt == VT_UNKNOWN || vt == VT_BSTR) pv = v.byref;
					else pv = (void*)&v.byref;
					SafeArrayPutElement(var.parray, rgIndices, pv);
				}
			}
		}
		vt = VT_EMPTY;
	}
	else if (val->IsObject()) {
		auto obj = Local<Object>::Cast(val);
		if (!DispObject::GetValueOf(isolate, obj, var) && !VariantObject::GetValueOf(isolate, obj, var)) {
			var.vt = VT_DISPATCH;
			var.pdispVal = new DispObjectImpl(obj);
			var.pdispVal->AddRef();
		}
	}
	else {
		String::Value str(isolate, val);
		var.vt = VT_BSTR;
		// For some apps there is still a difference between "" and NULL string, so we should support it here gracefully
		if (*str)
		{
			var.bstrVal = SysAllocString((LPOLESTR)*str);
		}
		else {
			var.bstrVal = 0;
		}
		//var.bstrVal = (str.length() > 0) ? SysAllocString((LPOLESTR)*str) : 0;
	}
	if (vt != VT_EMPTY && vt != VT_NULL && vt != VT_VARIANT) {
		if FAILED(VariantChangeType(&var, &var, 0, vt))
			VariantClear(&var);
	}
}

void Value2SafeArray(Isolate *isolate, Local<Value> &val, VARIANT &var, VARTYPE vt) {
    Local<Context> ctx = isolate->GetCurrentContext();
    if (val.IsEmpty() || val->IsUndefined()) {
        var.vt = VT_EMPTY;
    }
    else if (val->IsNull()) {
        var.vt = VT_NULL;
    }
    else if (val->IsObject()) {
        // Test conversion dispatch pointer to Uint8Array
        if (vt == VT_UI1) {
            auto obj = Local<Object>::Cast(val);
            auto ptr = DispObject::GetDispPtr(isolate, obj);
            size_t len = sizeof(UINT_PTR);
            var.vt = VT_ARRAY | vt;
            var.parray = SafeArrayCreateVector(vt, 0, len);
            if (var.parray != nullptr) memcpy(var.parray->pvData, &ptr, len);
        }
    }
    else {
        var.vt = VT_EMPTY;
    }
}

bool Value2Unknown(Isolate *isolate, Local<Value> &val, IUnknown **unk) {
    if (val.IsEmpty() || !val->IsObject()) return false;
    auto obj = Local<Object>::Cast(val);
    CComVariant var;
    if (!DispObject::GetValueOf(isolate, obj, var) && !VariantObject::GetValueOf(isolate, obj, var)) return false;
    return VariantUnkGet(&var, unk);
}

bool UnknownDispGet(IUnknown *unk, IDispatch **disp) {
	if (!unk) return false;
	if SUCCEEDED(unk->QueryInterface(__uuidof(IDispatch), (void**)disp)) {
		return true;
	}
	CComPtr<IEnumVARIANT> enum_ptr;
	if SUCCEEDED(unk->QueryInterface(__uuidof(IEnumVARIANT), (void**)&enum_ptr)) {
		*disp = new DispEnumImpl(enum_ptr);
		(*disp)->AddRef();
		return true;
	}
	return false;
}

bool VariantUnkGet(VARIANT *v, IUnknown **punk) {
	IUnknown *unk = NULL;
    if ((v->vt & VT_TYPEMASK) == VT_DISPATCH) {
		unk = ((v->vt & VT_BYREF) != 0) ? *v->ppdispVal : v->pdispVal;
    }
    else if ((v->vt & VT_TYPEMASK) == VT_UNKNOWN) {
        unk = ((v->vt & VT_BYREF) != 0) ? *v->ppunkVal : v->punkVal;
    }
	if (!unk) return false;
	unk->AddRef();
	*punk = unk;
	return true;
}

bool VariantDispGet(VARIANT *v, IDispatch **pdisp) {
	/*
	if ((v->vt & VT_ARRAY) != 0) {
		*disp = new DispArrayImpl(*v);
		(*disp)->AddRef();
		return true;
	}
	*/
	if ((v->vt & VT_TYPEMASK) == VT_DISPATCH) {
		IDispatch *disp = ((v->vt & VT_BYREF) != 0) ? *v->ppdispVal : v->pdispVal;
		if (!disp) return false;
        disp->AddRef();
		*pdisp = disp;
        return true;
    }
    if ((v->vt & VT_TYPEMASK) == VT_UNKNOWN) {
		return UnknownDispGet(((v->vt & VT_BYREF) != 0) ? *v->ppunkVal : v->punkVal, pdisp);
    }
    return false;
}

//-------------------------------------------------------------------------------------------------------
// DispArrayImpl implemetation

HRESULT STDMETHODCALLTYPE DispArrayImpl::GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) {
	if (cNames != 1 || !rgszNames[0]) return DISP_E_UNKNOWNNAME;
	LPOLESTR name = rgszNames[0];
	if (wcscmp(name, L"length") == 0) *rgDispId = 1;
	else return DISP_E_UNKNOWNNAME;
	return S_OK;
}

HRESULT STDMETHODCALLTYPE DispArrayImpl::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) {
	HRESULT hrcode = S_OK;
	UINT argcnt = pDispParams->cArgs;
	VARIANT *args = pDispParams->rgvarg;

	if ((var.vt & VT_ARRAY) == 0) return E_NOTIMPL;
	SAFEARRAY *arr = ((var.vt & VT_BYREF) != 0) ? *var.pparray : var.parray;

	switch (dispIdMember) {
	case 1: {
		if (pVarResult) {
			pVarResult->vt = VT_INT;
			pVarResult->intVal = (INT)(arr ? arr->rgsabound[0].cElements : 0);
		}
		return hrcode; }
	}
	return E_NOTIMPL;
}

//-------------------------------------------------------------------------------------------------------
// DispEnumImpl implemetation

HRESULT STDMETHODCALLTYPE DispEnumImpl::GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) {
    if (cNames != 1 || !rgszNames[0]) return DISP_E_UNKNOWNNAME;
    LPOLESTR name = rgszNames[0];
    if (wcscmp(name, L"Next") == 0) *rgDispId = 1;
    else if (wcscmp(name, L"Skip") == 0) *rgDispId = 2;
    else if (wcscmp(name, L"Reset") == 0) *rgDispId = 3;
    else if (wcscmp(name, L"Clone") == 0) *rgDispId = 4;
    else return DISP_E_UNKNOWNNAME;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE DispEnumImpl::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) {
    HRESULT hrcode = S_OK;
    UINT argcnt = pDispParams->cArgs;
    VARIANT *args = pDispParams->rgvarg;
    switch (dispIdMember) {
    case 1: {
        CComVariant arr;
        ULONG fetched, celt = (argcnt > 0) ? Variant2Int(args[argcnt - 1], (ULONG)1) : 1;
        if (!pVarResult || celt == 0) hrcode = E_INVALIDARG;
        if SUCCEEDED(hrcode) hrcode = arr.ArrayCreate(VT_VARIANT, celt);
        if SUCCEEDED(hrcode) hrcode = ptr->Next(celt, arr.ArrayGet<VARIANT>(0), &fetched);
        if SUCCEEDED(hrcode) {
            if (fetched == 0) pVarResult->vt = VT_EMPTY;
            else if (fetched == 1) {
                VARIANT *v = arr.ArrayGet<VARIANT>(0);
                *pVarResult = *v;
                v->vt = VT_EMPTY;
            }
            else {
                if (fetched < celt) hrcode = arr.ArrayResize(fetched);
                if SUCCEEDED(hrcode) arr.Detach(pVarResult);
            }
        }
        return hrcode; 
        }
    case 2: {
        if (pVarResult) pVarResult->vt = VT_EMPTY;
        ULONG celt = (argcnt > 0) ? Variant2Int(args[argcnt - 1], (ULONG)1) : 1;
        return ptr->Skip(celt); 
        }
    case 3: {
        if (pVarResult) pVarResult->vt = VT_EMPTY;
        return ptr->Reset(); 
        }
    case 4: {
        std::unique_ptr<DispEnumImpl> disp;
        hrcode = pVarResult ? ptr->Clone(&disp->ptr) : E_INVALIDARG;
        if SUCCEEDED(hrcode) {
            disp->AddRef();
            pVarResult->vt = VT_DISPATCH;
            pVarResult->pdispVal = disp.release();
        }
        return hrcode; }
    }
    return E_NOTIMPL;
}

//-------------------------------------------------------------------------------------------------------
// DispObjectImpl implemetation

HRESULT STDMETHODCALLTYPE DispObjectImpl::GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) {
	if (cNames != 1 || !rgszNames[0]) return DISP_E_UNKNOWNNAME;
	std::wstring name(rgszNames[0]);
	name_ptr &ptr = names[name];
	if (!ptr) {
		ptr.reset(new name_t(dispid_next++, name));
		index.insert(index_t::value_type(ptr->dispid, ptr));
	}
	*rgDispId = ptr->dispid;
	return S_OK;
}

HRESULT STDMETHODCALLTYPE DispObjectImpl::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) {
	Isolate *isolate = Isolate::GetCurrent();
	Local<Context> ctx = isolate->GetCurrentContext();
	Local<Object> self = obj.Get(isolate);
	Local<Value> name, val, ret;

	// Prepare name by member id
	if (!index.empty()) {
		index_t::const_iterator p = index.find(dispIdMember);
		if (p == index.end())
		{
			// DispID may be 0 for  regular member or for DISPID_VALUE.
			// Since regular method not found assume DISPID_VALUE if it is 0.
			if(dispIdMember != DISPID_VALUE) return DISP_E_MEMBERNOTFOUND;
		} else {
			name_t& info = *p->second;
			name = v8str(isolate, info.name.c_str());
		}
	}
	// Set property value
	if ((wFlags & DISPATCH_PROPERTYPUT) != 0) {
		UINT argcnt = pDispParams->cArgs;
		VARIANT *key = (argcnt > 1) ? &pDispParams->rgvarg[--argcnt] : nullptr;
		if (argcnt > 0) val = Variant2Value(isolate, pDispParams->rgvarg[--argcnt], true);
		else val = Undefined(isolate);
		bool rcode;

		// Set simple object property value
		if (!key) {
			if (name.IsEmpty()) return DISP_E_MEMBERNOTFOUND;
			rcode = self->Set(ctx, name, val).FromMaybe(false);
		}

		// Set object/array item value
		else {
			Local<Object> target;
			if (name.IsEmpty()) target = self;
			else {
				Local<Value> obj;
				if (self->Get(ctx, name).ToLocal(&obj) && !obj.IsEmpty()) target = Local<Object>::Cast(obj);
				if (target.IsEmpty()) return DISP_E_BADCALLEE;
			}

			LONG index = Variant2Int<LONG>(*key, -1);
			if (index >= 0) rcode = target->Set(ctx, (uint32_t)index, val).FromMaybe(false);
			else rcode = target->Set(ctx, Variant2Value(isolate, *key, false), val).FromMaybe(false);
		}

		// Store result
		if (pVarResult) {
			pVarResult->vt = VT_BOOL;
			pVarResult->boolVal = rcode ? VARIANT_TRUE : VARIANT_FALSE;
		}
		return S_OK;
	}

	// Prepare property item
	if (name.IsEmpty()) val = self;
	else self->Get(ctx, name).ToLocal(&val);

	// Call property as method
	if ((wFlags & DISPATCH_METHOD) != 0) {
		wFlags = 0;
		NodeArguments args(isolate, pDispParams, true, reverse_arguments);
		int argcnt = (int)args.items.size();
		Local<Value> *argptr = (argcnt > 0) ? &args.items[0] : nullptr;
		if (val->IsFunction()) {
			Local<Function> func = Local<Function>::Cast(val);
			if (func.IsEmpty()) return DISP_E_BADCALLEE;
			func->Call(isolate->GetCurrentContext(), self, argcnt, argptr).ToLocal(&ret);
		}
		else if (val->IsObject()) {
			wFlags = DISPATCH_PROPERTYGET;
			//Local<Object> target = val->ToObject();
			//target->CallAsFunction(isolate->GetCurrentContext(), target, args.items.size(), &args.items[0]).ToLocal(&ret);
		}
		else {
			ret = val;
		}
	}

	// Get property value
	if ((wFlags & DISPATCH_PROPERTYGET) != 0) {
		if (pDispParams->cArgs == 1) {
			Local<Object> target;
			if (!val.IsEmpty()) target = Local<Object>::Cast(val);
			if (target.IsEmpty()) return DISP_E_BADCALLEE;
			VARIANT &key = pDispParams->rgvarg[0];
			LONG index = Variant2Int<LONG>(key, -1);
			if (index >= 0) target->Get(ctx, (uint32_t)index).ToLocal(&ret);
			else target->Get(ctx, Variant2Value(isolate, key, false)).ToLocal(&ret);
		}
		else {
			ret = val;
		}
	}

	// Store result
	if (pVarResult) {
		Value2Variant(isolate, ret, *pVarResult, VT_NULL);
	}
	return S_OK;
}

/*
* Microsoft OLE Date type:
* https://docs.microsoft.com/en-us/previous-versions/visualstudio/visual-studio-2008/82ab7w69(v=vs.90)
*/

double FromOleDate(double oleDate) {
    double posixDate = oleDate - 25569; // days from 1899 dec 30
    posixDate *= 24 * 60 * 60 * 1000;   // days to milliseconds
    return posixDate;
}

double ToOleDate(double posixDate) {
    double oleDate = posixDate / (24 * 60 * 60 * 1000); // milliseconds to days
    oleDate += 25569;                                   // days from 1899 dec 30
    return oleDate;
}

//-------------------------------------------------------------------------------------------------------

bool NodeMethods::get(Isolate* isolate, const std::wstring &name, Local<Function>* value) {
    auto ptr = items.find(name);
    if (ptr == items.end()) return false;
    Local<FunctionTemplate> ft = ptr->second->Get(isolate);
    return ft->GetFunction(isolate->GetCurrentContext()).ToLocal(value);
}

void NodeMethods::add(Isolate* isolate, Local<FunctionTemplate>& clazz, const char* name, FunctionCallback callback) {
    // see NODE_SET_PROTOTYPE_METHOD
    HandleScope handle_scope(isolate);
    Local<Signature> s = Signature::New(isolate, clazz);
    Local<FunctionTemplate> t = FunctionTemplate::New(isolate, callback, Local<Value>(), s);
    Local<String> fn_name = String::NewFromUtf8(isolate, name, NewStringType::kInternalized).ToLocalChecked();
    t->SetClassName(fn_name);
    
    clazz->PrototypeTemplate()->Set(fn_name, t);

    String::Value vname(isolate, fn_name);
    item_type item(new Persistent<FunctionTemplate>(isolate, t));
    items.emplace(std::wstring((const wchar_t *)*vname), item);
}

//-------------------------------------------------------------------------------------------------------

/* Message loop. Just like with WScript it is executed while the script is waiting. 
   So if you have something showing, for example, a popup message, then you will only
   see it when you do WScript.Sleep.
*/

void DoEvents()
{
	MSG msg;
	BOOL result;

	if (::PeekMessage(&msg, NULL, 0, 0, PM_NOREMOVE))
	{
		result = ::GetMessage(&msg, NULL, 0, 0);
		if (result == 0) // WM_QUIT
		{
			::PostQuitMessage(msg.wParam);
		}
		else if (result == -1)
		{
			// Handle errors/exit application, etc.
		}
		else
		{
			::TranslateMessage(&msg);
			::DispatchMessage(&msg);
		}
	}
}

/* Returns the amount of milliseconds elapsed since the UNIX epoch. Works on both
 * windows and linux. */

long long GetTimeMs64()
{
	/* Windows */
	FILETIME ft;
	LARGE_INTEGER li;

	/* Get the amount of 100 nano seconds intervals elapsed since January 1, 1601 (UTC) and copy it
	 * to a LARGE_INTEGER structure. */
	GetSystemTimeAsFileTime(&ft);
	li.LowPart = ft.dwLowDateTime;
	li.HighPart = ft.dwHighDateTime;

	long long ret = li.QuadPart;
	ret -= 116444736000000000LL; /* Convert from file time to UNIX epoch time. */
	ret /= 10000; /* From 100 nano seconds (10^-7) to 1 millisecond (10^-3) intervals */

	return ret;

}

// Sleep is essential to have proper WScript emulation
void WinaxSleep(const FunctionCallbackInfo<Value>& args) {
	Isolate* isolate = args.GetIsolate();
	Local<Context> ctx = isolate->GetCurrentContext();

	if (args.Length() == 0 && !args[0]->IsUint32()) {
		isolate->ThrowException(InvalidArgumentsError(isolate));
		return;
	}
	uint32_t ms = (args[0]->Uint32Value(ctx)).FromMaybe(0);
	long long start = GetTimeMs64();
	do
	{
		DoEvents();
		Sleep(1);
	} while (GetTimeMs64() - start < ms);
	args.GetReturnValue().SetUndefined();
}

// Get a COM pointer from a window found by it's text
HRESULT GetAccessibleObject(const wchar_t* pszWindowText, CComPtr<IUnknown>& spIUnknown) {
	struct ew {
		static BOOL CALLBACK ecp(HWND hWnd, LPARAM lParam) {
			wchar_t szWindowText[128];
			if (GetWindowTextW(hWnd, szWindowText, _countof(szWindowText))) {
				ewp* pparams = reinterpret_cast<ewp*>(lParam);
				if (!wcscmp(szWindowText, pparams->pszWindowText)) {
					pparams->hWnd = hWnd;
					return FALSE;
				}
			}
			return TRUE;
		}
		struct ewp {
			const wchar_t* pszWindowText;
			HWND hWnd;
		};
	};
	ew::ewp params{ pszWindowText, nullptr };
	EnumChildWindows(GetDesktopWindow(), ew::ecp, reinterpret_cast<LPARAM>(&params));
	if (params.hWnd == nullptr) return _HRESULT_TYPEDEF_(0x80070057L); // ERROR_INVALID_PARAMETER
	return AccessibleObjectFromWindow(params.hWnd, OBJID_NATIVEOM, IID_IUnknown,
	                                  reinterpret_cast<void**>(&spIUnknown));
}