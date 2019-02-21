#pragma once

#define TEST_ADVISE

#include <SDKDDKVer.h>

// Windows Header Files:
#define WIN32_LEAN_AND_MEAN                     // Exclude rarely-used stuff from Windows headers
#define _ATL_CSTRING_EXPLICIT_CONSTRUCTORS      // some CString constructors will be explicit
#include <windows.h>

// ATL headers
//#define USE_ATL
#ifdef USE_ATL
#include <atlbase.h>
#include <atlstr.h>
#else
#include <ole2.h>
#include <ocidl.h>
#endif

// STD headers
#include <iostream>
#include <stdio.h>
#include <string>
#include <vector>
#include <map>
#include <memory>
#include <initializer_list>

// Node JS headers
#include <v8.h>
#include <node.h>
#include <node_version.h>
#include <node_object_wrap.h>
#include <node_buffer.h>
using namespace v8;
using namespace node;

// Application defines
//#define TEST_ADVISE