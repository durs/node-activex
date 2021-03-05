//-------------------------------------------------------------------------------------------------------
// Project: node-activex
// Author: Yuri Dursin
// Description: Defines the entry point for the NodeJS addon
//-------------------------------------------------------------------------------------------------------

#include "stdafx.h"
#include "disp.h"

//----------------------------------------------------------------------------------

// Initialize this addon
NODE_MODULE_INIT(/*exports, module, context*/) {
  Isolate* isolate = context->GetIsolate();

  DispObject::NodeInit(exports, isolate, context);
  VariantObject::NodeInit(exports, isolate, context);
  ConnectionPointObject::NodeInit(exports, isolate, context);

  // Sleep is essential to have proper WScript emulation
  exports->Set(context,
      String::NewFromUtf8(isolate, "winaxsleep", NewStringType::kNormal)
      .ToLocalChecked(),
      FunctionTemplate::New(isolate, WinaxSleep)
      ->GetFunction(context).ToLocalChecked()).FromJust();

  /* Example implementation of a context-aware addon:
  // Create a new instance of `AddonData` for this instance of the addon and
  // tie its life cycle to that of the Node.js environment.
  AddonData* data = new AddonData(isolate);

  // Wrap the data in a `v8::External` so we can pass it to the method we
  // expose.
  Local<External> external = External::New(isolate, data);

  // Expose the method `Method` to JavaScript, and make sure it receives the
  // per-addon-instance data we created above by passing `external` as the
  // third parameter to the `FunctionTemplate` constructor.
  exports->Set(context,
               String::NewFromUtf8(isolate, "method", NewStringType::kNormal)
                  .ToLocalChecked(),
               FunctionTemplate::New(isolate, Method, external)
                  ->GetFunction(context).ToLocalChecked()).FromJust();
  */
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
