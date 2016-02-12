//-------------------------------------------------------------------------------------------------------
// Project: NodeActiveX
// Author: Yuri Dursin
// Last Modification: 2011-11-22
// Description: Defines the entry point for the NodeJS addon
//-------------------------------------------------------------------------------------------------------

#include "stdafx.h"
#include "DispNodeJS.h"

#pragma comment(lib, "node")

//----------------------------------------------------------------------------------

BOOL APIENTRY DllMain( HMODULE hModule,
                       DWORD  ul_reason_for_call,
                       LPVOID lpReserved
					 )
{
	switch (ul_reason_for_call)
	{
	case DLL_PROCESS_ATTACH:
		CoInitialize(0);
		break;

	case DLL_THREAD_ATTACH:
		break;

	case DLL_THREAD_DETACH:
		break;

	case DLL_PROCESS_DETACH:
		CoUninitialize();
		break;
	}
	return TRUE;
}

//----------------------------------------------------------------------------------

extern "C" void NODE_EXTERN init (Handle<Object> target)
{
	DispObject::NodeInit(target);
}

NODE_MODULE(activex, init);

//----------------------------------------------------------------------------------
