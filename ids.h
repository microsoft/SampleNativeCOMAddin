/*!-----------------------------------------------------------------------
	ids.h

	Contains all the progids/guids/clsids used across the addin. If an
	id is needed to be updated then it should only need to be changed
	here and nowhere else.

	Remember these ids need to be unique on each system so they should be
	changed if this sample addin is used as a base for a another addin.
-----------------------------------------------------------------------!*/
#pragma once

#define ADDIN_PROGID            OLESTR("SampleNativeCOMAddin.Connect")
#define SAMPLECONTROL_PROGID    OLESTR("SampleNativeCOMAddin.SampleControl")

#define TYPELIB_GUID            B715B90F-6F65-4848-B73A-37D550C7E726
#define ADDIN_CLSID             1E1AC2A3-72BB-453F-92AE-1BE63F1F88BC
#define IRIBBONCALLBACK_IID     0BDAF081-9E88-48F2-9265-1FF7136BF3ED
#define ISAMPLECONTROL_IID      88F2DFD1-B331-4D14-889F-43481CC11E40
#define SAMPLECONTROL_CLSID     D062F7B8-FBD7-46B6-9A34-A5FCBD6DBC78

// Remember to keep these string version in sync with the ones above
#define TYPELIB_GUID_STR        OLESTR("B715B90F-6F65-4848-B73A-37D550C7E726")
#define ADDIN_CLSID_STR         OLESTR("1E1AC2A3-72BB-453F-92AE-1BE63F1F88BC")
#define IRIBBONCALLBACK_IID_STR OLESTR("0BDAF081-9E88-48F2-9265-1FF7136BF3ED")
#define ISAMPLECONTROL_IID_STR  OLESTR("88F2DFD1-B331-4D14-889F-43481CC11E40")
#define SAMPLECONTROL_CLSID_STR OLESTR("D062F7B8-FBD7-46B6-9A34-A5FCBD6DBC78")

