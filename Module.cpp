/*!-----------------------------------------------------------------------
	module.cpp
-----------------------------------------------------------------------!*/
#include "stdafx.h"

CAddInModule _AtlModule;

// DLL Entry Point
extern "C" BOOL WINAPI DllMain(HINSTANCE hInstance, DWORD dwReason, LPVOID lpReserved)
{
	_AtlModule.SetResourceInstance(hInstance);
	return _AtlModule.DllMain(dwReason, lpReserved);
}

// Used to determine whether the DLL can be unloaded by OLE
STDAPI DllCanUnloadNow()
{
	return _AtlModule.DllCanUnloadNow();
}

// Returns a class factory to create an object of the requested type
STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID* ppv)
{
	return _AtlModule.DllGetClassObject(rclsid, riid, ppv);
}

// DllRegisterServer - Adds entries to the system registry
STDAPI DllRegisterServer()
{
	// registers object, typelib and all interfaces in typelib
	return _AtlModule.DllRegisterServer();
}

// DllUnregisterServer - Removes entries from the system registry
STDAPI DllUnregisterServer()
{
	return _AtlModule.DllUnregisterServer();
}