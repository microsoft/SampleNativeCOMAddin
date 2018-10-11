# SampleNativeCOMAddin
Basic native/C++ com add-in for quick and dirty tests

To install the addin run:
``` batch
regsvr32 SampleNativeCOMAddin.dll
```

To uninstall the addin run:
``` batch
regsvr32 /u SampleNativeCOMAddin.dll
```

The register and unregister will take care of all the registry that is needed including:
1) COM registration of addin and sample control
2) Registration of the addin with Outlook
3) Registration of the addin form region with Outlook

## Confirm load
File/Options/Add-ins, look for "Outlook Sample Native COM Addin Connect Class Object"

## Debugging
Simplest debugging is on the machine where the add-in was built.
1) Register addin as above
2) Start Outlook
3) Debug/Attach to process. Make sure to attach as native code. Choose Outlook.
4) Set breakpoints and go!
