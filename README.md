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
