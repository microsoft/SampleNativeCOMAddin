/*!-----------------------------------------------------------------------
	outlook.h

	Imports for some of the standard outlook type libraries.
-----------------------------------------------------------------------!*/
#pragma once

// Office type library (i.e. mso.dll)
#import "libid:2DF8D04C-5BFA-101B-BDE5-00AA0044DE52"\
	auto_rename auto_search raw_interfaces_only rename_namespace("Office")\
	rename("DocumentProperties", "__DocumentProperties")\
	rename("SearchPath", "__SearchPath")\
	rename("RGB", "__RGB")

// Outlook type library (i.e. msoutl.olb)
#import "libid:00062FFF-0000-0000-C000-000000000046"\
	auto_rename auto_search raw_interfaces_only rename_namespace("Outlook")\
	rename("CopyFile", "__CopyFile")\
	rename("PlaySound", "__PlaySound")

// Forms type library (i.e. fm20.dll)
#import "libid:0D452EE1-E08F-101A-852E-02608C4D0BB4"\
	auto_rename auto_search raw_interfaces_only rename_namespace("Forms")\
	rename("OLE_COLOR", "__OLE_COLOR")\
	exclude("IFont")

using namespace Outlook;
using namespace Office;
using namespace Forms;

// dispids used for various events
const DISPID dispidEventOnClose = 0xF004;
const DISPID dispidEventOptionsPagesAdd = 0xF005;
