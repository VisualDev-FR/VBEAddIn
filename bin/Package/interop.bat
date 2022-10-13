@ECHO OFF
"C:\Program Files (x86)\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.7.1 Tools\TlbImp.exe" ^
   "C:\Program Files (x86)\Common Files\DESIGNER\MSADDNDR.DLL" ^
   /out:"VBEAddIn.Interop.Extensibility.dll" ^
   /keyfile:"C:\TMenanteau.snk" ^
   /strictref:nopia /nologo /asmversion:1.0.0.0 /sysarray

PAUSE
CLS

"C:\Program Files (x86)\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.7.1 Tools\TlbImp.exe" ^
   "C:\Windows\SysWOW64\stdole2.tlb" ^
   /out:"VBEAddIn.Interop.Stdole.dll" ^
   /keyfile:"C:\TMenanteau.snk" ^
   /strictref:nopia /nologo /asmversion:1.0.0.0

PAUSE
CLS

"C:\Program Files (x86)\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.7.1 Tools\TlbImp.exe" ^
   "C:\Program Files (x86)\Common Files\Microsoft Shared\OFFICE11\MSO.DLL" ^
   /out:"VBEAddIn.Interop.Office11.dll" ^
   /keyfile:"C:\TMenanteau.snk" ^
   /strictref:nopia /nologo /asmversion:1.0.0.0 ^
   /reference:VBEAddIn.Interop.Stdole.dll

PAUSE
CLS

"C:\Program Files (x86)\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.7.1 Tools\TlbImp.exe" ^
   "C:\Program Files (x86)\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB" ^
   /out:"VBEAddIn.Interop.VBAExtensibility.dll" ^
   /keyfile:"C:\TMenanteau.snk" ^
   /strictref:nopia /nologo /asmversion:1.0.0.0 ^
   /reference:VBEAddIn.Interop.Office11.dll

PAUSE