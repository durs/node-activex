#pragma once

#include "targetver.h"

// Windows Header Files:
#define WIN32_LEAN_AND_MEAN             // Exclude rarely-used stuff from Windows headers
#include <windows.h>
#define _ATL_CSTRING_EXPLICIT_CONSTRUCTORS      // some CString constructors will be explicit
#include <atlbase.h>
#include <atlstr.h>

// STD headers
#include <iostream>
#include <stdio.h>
#include <string>
#include <vector>
using namespace std;

// Node JS headers
#include <v8.h>
#include <node.h>
#include <node_version.h>
#include <node_object_wrap.h>
#include <node_buffer.h>
using namespace v8;
using namespace node;