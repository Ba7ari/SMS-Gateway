copy aryolib2.* %SYSTEMROOT%\system32 /y
copy mfbus15.* %SYSTEMROOT%\system32 /y
copy gjfbus15.dll %SYSTEMROOT%\system32 /y
copy msgdll.* %SYSTEMROOT%\system32 /y
%SYSTEMROOT%\system32\regsvr32 %SYSTEMROOT%\system32\aryolib2.dll
%SYSTEMROOT%\system32\regsvr32 %SYSTEMROOT%\system32\mfbus15.ocx
%SYSTEMROOT%\system32\regsvr32 %SYSTEMROOT%\system32\msgdll.dll
