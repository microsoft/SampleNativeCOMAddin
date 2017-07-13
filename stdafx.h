/*!-----------------------------------------------------------------------
	stdafx.h
-----------------------------------------------------------------------!*/
#pragma once

#include <atlbase.h>
#include <atlcom.h>
using namespace ATL;

#include "ids.h"
#include "resource.h"
#include "outlook.h"

// Import for IDTExtensibility2
#import "libid:AC0714F2-3D04-11D1-AE7D-00A0C90F26F4"\
	auto_rename auto_search raw_interfaces_only rename_namespace("AddinDesign")
using namespace AddinDesign;

// This addin type library
#import "addin.tlb" auto_rename auto_search raw_interfaces_only rename_namespace("SampleAddin")
using namespace SampleAddin;

class CAddInModule : public CAtlDllModuleT< CAddInModule >
{
public:
	CAddInModule()
	{
		m_hInstance = NULL;
	}

	DECLARE_LIBID(__uuidof(__SampleNativeCOMAddinLib))

	inline HINSTANCE GetResourceInstance()
	{
		return m_hInstance;
	}

	inline void SetResourceInstance(HINSTANCE hInstance)
	{
		m_hInstance = hInstance;
	}

private:
	HINSTANCE m_hInstance;
};

extern CAddInModule _AtlModule;
