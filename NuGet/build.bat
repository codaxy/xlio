pushd %~dp0
del *.nupkg

mkdir Codaxy.Xlio\lib
mkdir Codaxy.Xlio\lib\net35
copy ..\Source\Codaxy.Xlio\bin\Release\Codaxy*.* Codaxy.Xlio\lib\net35
..\Source\.nuget\nuget pack Codaxy.Xlio\Codaxy.Xlio.nuspec /verbose

if not ERRORLEVEL 0 pause

popd