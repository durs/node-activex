//-------------------------------------------------------------------------------------------------------
// Project: node-activex
// Author: Yuri Dursin
// Description: Defines the entry point for the NodeJS addon
//-------------------------------------------------------------------------------------------------------

#include "stdafx.h"
#include "disp.h"

//----------------------------------------------------------------------------------

namespace node_activex {

    void Init(Local<Object> exports) {
        DispObject::NodeInit(exports);
		VariantObject::NodeInit(exports);

#ifdef TEST_ADVISE 
        ConnectionPointObject::NodeInit(exports);
#endif
    }

    NODE_MODULE(node_activex, Init)
}

//----------------------------------------------------------------------------------

BOOL APIENTRY DllMain(HMODULE hModule, DWORD  ulReason, LPVOID lpReserved) {
	
    switch (ulReason) {
    case DLL_PROCESS_ATTACH:
        CoInitialize(0);
        break;
    case DLL_PROCESS_DETACH:
        CoUninitialize();
        break;
    case DLL_THREAD_ATTACH:
        break;
    case DLL_THREAD_DETACH:
        break;
    }
    return TRUE;
}

//----------------------------------------------------------------------------------
