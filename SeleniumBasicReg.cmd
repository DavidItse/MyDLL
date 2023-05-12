:: 
:: This script installs SeleniumBasic without administrator privileges.
:: It registers a COM API running on the .NET Framework.
::
:: The required files can be extracted with innoextract from the original setup :
::   Selenium.dll  Selenium32.tlb  Selenium64.tlb  Selenium.pdb
::
:: The drivers are not provided. You'll have to download and place the desired driver :
::  * in a folder defined in the "PATH" environment variable
::  * or in this folder before running this script
::  * or in the installed folder after installation
:: Note that the group policy may block the drivers depending on the location or whitelist.
::
:: By default, the files are installed in "%APPDATA%\SeleniumBasic".
:: To change the install folder, edit "set _LOCATION=..." in this script.
::
:: By default, the library will associate with the latest installed .NET runtime.
:: To run a specific runtime version, edit "set _RUNTIME=v?.?.*" in this script.
:: It will fail in Office if a different runtime version is already loaded by an extension.
:: If it's the case, try a different version (ex: v1.0.3705 v2.0.50727 v4.0.30319).
::
:: To uninstall, open "Programs and features" and double-click on "SeleniumBasic"
:: or execute "%APPDATA%\SeleniumBasic\uninstall.cmd"
::

@echo off
setlocal EnableDelayedExpansion
set PATH=%SYSTEMROOT%\System32;%SYSTEMROOT%\System32\wbem
chcp 65001 >nul 2>nul
pushd "%~dp0" || goto :failed

set _NAME=SeleniumBasic
set _HOMEPAGE=https://github.com/florentbr/SeleniumBasic
set _LOCATION=%APPDATA%\SeleniumBasic
set _FILES=Selenium.dll Selenium32.tlb Selenium64.tlb Selenium.pdb
set _RUNTIME=v?.?.*

for %%f in (%_FILES%) do if not exist "%%~f" >&2 echo Error: file %%f not found & goto :failed
call :get_version _VERSION %_FILES% || goto :failed
call :get_runtime _RUNTIME || goto :failed

echo =============================================================================
echo  %_NAME%  %_VERSION%
echo  Microsoft .NET Framework %_RUNTIME%
echo  %_HOMEPAGE%
echo =============================================================================

if exist "%_LOCATION%\uninstall.cmd" (
	echo Uninstall previous version ...
	cmd /c "%_LOCATION%\uninstall.cmd" >nul
	timeout /t 1 >nul
)

echo Install to %_LOCATION% ...
xcopy Selenium* "%_LOCATION%\" /y >nul || goto :failed
xcopy *driver.exe "%_LOCATION%\" /y >nul 2>nul || goto :failed
call :build_setup >"%_LOCATION%\setup.inf" || goto :failed
> "%_LOCATION%\uninstall.cmd" (
	echo @set PATH=%%SYSTEMROOT%%\System32
	echo @%%SYSTEMROOT%%\SysWOW64\rundll32 advpack.dll,LaunchINFSection "%%~dp0setup.inf",DefaultUninstall,3
	echo @%%SYSTEMROOT%%\System32\rundll32 advpack.dll,LaunchINFSection "%%~dp0setup.inf",DefaultUninstall,3
	echo @start "" /b cmd /c rmdir /s /q "%%~dp0"
)

echo Register application ...
%SYSTEMROOT%\SysWOW64\rundll32 advpack.dll,LaunchINFSection "%_LOCATION%\setup.inf",DefaultInstall,3 2>nul
%SYSTEMROOT%\System32\rundll32 advpack.dll,LaunchINFSection "%_LOCATION%\setup.inf",DefaultInstall,2 || goto :failed

echo Test CreateObject with VBS script ...
> "%_LOCATION%\test.vbs" echo CreateObject "Selenium.ChromeDriver"
(cscript //nologo "%_LOCATION%\test.vbs" >nul) 2>&1 | findstr /r "." && goto :failed

echo Done ^^!
popd & endlocal & pause >nul & exit \b 0

:failed
>&2 echo Failed ^^!
popd & endlocal & pause >nul & exit \b 1



:get_runtime
for /d %%f in ("%SYSTEMROOT%\Microsoft.NET\Framework\!%1!") do if exist "%%f\mscorlib.dll" set "%1=%%~nxf"
if exist "%SYSTEMROOT%\Microsoft.NET\Framework\!%1!\mscorlib.dll" exit /b 0
>&2 echo Error: .NET Framework !%1! not found.
exit /b 1


:get_version
set _file=%~f2
set _query=wmic datafile where "Name='%_file:\=\\%'" get Version
for /f "usebackq skip=1" %%a in (`%_query% 2^>nul`) do for %%l in (%%a) do set "%1=%%l" & exit /b 0
>&2 echo Error: failed to read version of %_file%
exit /b 1


:build_setup
pushd "%_LOCATION%" || exit /b 0
call :build_reg 4>1 5>2 6>3 >nul
cmd /d /u /c for %%f in (1 2 3) do @type %%f ^& del %%f
popd >nul
goto :eof


:build_reg

set _libname=Selenium Type Library
set _libguid={0277FC34-FD1B-4616-BB19-A9AABCAF2A70}
set _libsign=Selenium, Version=!_VERSION!, Culture=neutral
set _libvers=2.0

>&4 echo [Version]
>&4 echo Signature="$Windows NT$"
>&4 echo AdvancedINF=2.5
>&4 echo.
>&4 echo [DefaultInstall]
>&4 echo DelReg = DelReg
>&4 echo AddReg = AddReg
>&4 echo.
>&4 echo [DefaultUninstall]
>&4 echo DelReg = DelReg

>&5 echo.
>&5 echo [AddReg]
>&5 echo HKCU,SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\SeleniumBasic,DisplayName,,"!_NAME!"
>&5 echo HKCU,SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\SeleniumBasic,DisplayVersion,,"!_VERSION!"
>&5 echo HKCU,SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\SeleniumBasic,InstallLocation,,"!CD!"
>&5 echo HKCU,SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\SeleniumBasic,UninstallString,,"""!CD!\uninstall.cmd"""
>&5 echo HKCU,SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\SeleniumBasic,URLInfoAbout,,"!_HOMEPAGE!"
>&5 echo HKCU,SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\SeleniumBasic,NoModify,0x00010001,1
>&5 echo HKCU,SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\SeleniumBasic,NoRepair,0x00010001,1
>&5 echo HKCU,SOFTWARE\Classes\TypeLib\!_libguid!\!_libvers!,,,"!_libname!"
>&5 echo HKCU,SOFTWARE\Classes\TypeLib\!_libguid!\!_libvers!\FLAGS,,,"0"
>&5 echo HKCU,SOFTWARE\Classes\TypeLib\!_libguid!\!_libvers!\0\win32,,,"!CD!\Selenium32.tlb"
>&5 echo HKCU,SOFTWARE\Classes\TypeLib\!_libguid!\!_libvers!\0\win64,,,"!CD!\Selenium64.tlb"

>&6 echo.
>&6 echo [DelReg]
>&6 echo HKCU,SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\SeleniumBasic
>&6 echo HKCU,SOFTWARE\Classes\TypeLib\!_libguid!

for %%l in (
	0809389E78C4:PhantomJSDriver
	14DB1E4916D4:FirefoxDriver
	3C406728F1A2:EdgeDriver
	44A424DB3F50:Timeouts
	5D556733E8C9:ChromeDriver
	5DB46A739EEA:List
	6AAF7EDD33D6:Assert
	7D30CBC3F6BB:Waiter
	80B2B91F0D44:By
	9E7F9EF1D002:OperaDriver
	A34FCBA29598:Utils
	B0C8C528C673:Verify
	B719752452AA:Table
	BE75D14E7B41:Keys
	CDCD9EB97FD6:PdfFile
	CEA7D8FD6954:Dictionary
	E3CCFFAB4234:WebDriver
	E9AAFA695FFB:Application
	EED04A1E4CD1:IEDriver
) do for /f "tokens=1,2 delims=:" %%i in ("%%l") do (
	set _name=Selenium.%%j
	set _guid={0277FC34-FD1B-4616-BB19-%%i}
	>&6 echo HKCU,SOFTWARE\Classes\!_name!
	>&6 echo HKCU,SOFTWARE\Classes\CLSID\!_guid!
	>&5 echo HKCU,SOFTWARE\Classes\!_name!,,,"!_name!"
	>&5 echo HKCU,SOFTWARE\Classes\!_name!\CLSID,,,"!_guid!"
	>&5 echo HKCU,SOFTWARE\Classes\CLSID\!_guid!,,,"!_name!"
	>&5 echo HKCU,SOFTWARE\Classes\CLSID\!_guid!\InprocServer32,,,"!SYSTEMROOT!\System32\mscoree.dll"
	>&5 echo HKCU,SOFTWARE\Classes\CLSID\!_guid!\InprocServer32,Class,,"!_name!"
	>&5 echo HKCU,SOFTWARE\Classes\CLSID\!_guid!\InprocServer32,Assembly,,"!_libsign!"
	>&5 echo HKCU,SOFTWARE\Classes\CLSID\!_guid!\InprocServer32,CodeBase,,"!CD!\Selenium.dll"
	>&5 echo HKCU,SOFTWARE\Classes\CLSID\!_guid!\InprocServer32,RuntimeVersion,,"!_RUNTIME!"
	>&5 echo HKCU,SOFTWARE\Classes\CLSID\!_guid!\InprocServer32,ThreadingModel,,"Both"
	>&5 echo HKCU,SOFTWARE\Classes\CLSID\!_guid!\ProgId,,,"!_name!"
	>&5 echo HKCU,SOFTWARE\Classes\CLSID\!_guid!\VersionIndependentProgID,,,"!_name!"
	>&5 echo HKCU,SOFTWARE\Classes\CLSID\!_guid!\Implemented Categories\{62C8FE65-4EBB-45e7-B440-6E39B2CDBF29},,,""
)

for %%l in (
	01D514FE0B1A:20:_Utils
	0B61E370369D:24:_TableRow
	0EA52ACB97D1:20:_Assert
	11660D7615B7:20:_Manage
	1456C48D8E5C:24:_Dictionary
	2276E80F5CF7:20:_Cookie
	384C7E50EFA8:20:_Waiter
	495CC9DBFB96:20:_Verify
	4CE442A16502:20:_SelectElement
	54BA7C175990:20:_Image
	61DAD6C51012:20:_Keyboard
	637431245D48:24:_Keys
	63F894CA99E9:20:_Mouse
	6E0522EA435E:20:_Application
	74F5D5680428:20:_Timeouts
	7C9763568492:20:_WebElements
	7E2EBB6C82E9:20:_Size
	8B145197B76C:20:_WebElement
	A398E67A519B:24:_DictionaryItem
	A3DE5685A27E:20:_By
	ACE280CD7780:20:_Point
	B51CB7C5A694:20:_Alert
	B825A6BF9610:24:_Table
	BBE48A6D09DB:20:_Actions
	BE15C121F199:20:_TableElement
	C539CB44B63F:24:_List
	C6F450B6EE52:20:_Storage
	CC6284398AA5:20:_WebDriver
	D0E30A5D0697:20:_Proxy
	D5DE929CF018:20:_TouchActions
	E6E7ED329824:20:_Cookies
	F2A56C3A68D4:20:_PdfFile
	FBDA3A91C82B:20:_Window
	FFD6FAEF290A:20:_TouchScreen
) do for /f "tokens=1,2,3 delims=:" %%i in ("%%l") do (
	set _name=%%k
	set _guid={0277FC34-FD1B-4616-BB19-%%i}
	set _stub={000204%%j-0000-0000-C000-000000000046}
	>&6 echo HKCU,SOFTWARE\Classes\Interface\!_guid!
	>&5 echo HKCU,SOFTWARE\Classes\Interface\!_guid!,,,"!_name!"
	>&5 echo HKCU,SOFTWARE\Classes\Interface\!_guid!\ProxyStubClsid32,,,"!_stub!"
	>&5 echo HKCU,SOFTWARE\Classes\Interface\!_guid!\TypeLib,,,"!_libguid!"
	>&5 echo HKCU,SOFTWARE\Classes\Interface\!_guid!\TypeLib,Version,,"!_libvers!"
)

for %%l in (
	300DAA508541:Strategy
	B342CE81CB2A:MouseButton
	C724C5135B6E:CacheState
) do for /f "tokens=1,2 delims=:" %%i in ("%%l") do (
	set _name=Selenium.%%j
	set _guid={0277FC34-FD1B-4616-BB19-%%i}
	>&6 echo HKCU,SOFTWARE\Classes\Record\!_guid!
	>&5 echo HKCU,SOFTWARE\Classes\Record\!_guid!,Class,,"!_name!"
	>&5 echo HKCU,SOFTWARE\Classes\Record\!_guid!,Assembly,,"!_libsign!"
	>&5 echo HKCU,SOFTWARE\Classes\Record\!_guid!,RuntimeVersion,,"!_RUNTIME!"
	>&5 echo HKCU,SOFTWARE\Classes\Record\!_guid!,CodeBase,,"!CD!\Selenium.dll"
)

goto :eof