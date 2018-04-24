copy aryolib2.* %SYSTEMROOT%\system /y
copy mfbus15.* %SYSTEMROOT%\system /y
copy gjfbus15.dll %SYSTEMROOT%\system /y
copy msgdll.* %SYSTEMROOT%\system /y
%SYSTEMROOT%\system\regsvr32 %SYSTEMROOT%\system\aryolib2.dll
%SYSTEMROOT%\system\regsvr32 %SYSTEMROOT%\system\mfbus15.ocx
%SYSTEMROOT%\system\regsvr32 %SYSTEMROOT%\system\msgdll.dll