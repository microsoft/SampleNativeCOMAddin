#pragma once

#define IDR_ADDIN                       100
#define IDR_SAMPLECONTROL               112
#define IDD_SAMPLECONTROL               113
#define IDC_BUTTON1                     201
#define IDS_CUSTOMRIBBON                1000
#define IDS_FORMREGIONMANIFEST          1001
#define IDS_FORMREGIONSTORAGE           1002

HRESULT HrGetResource(int nId, LPCTSTR lpType, LPVOID* ppvResourceData, DWORD* pdwSizeInBytes);
BSTR GetXMLResource(int nId);
SAFEARRAY* GetOFSResource(int nId);
