del *.nupkg

mkdir Codaxy.Xlio\lib
copy ..\Libraries\Codaxy.Xlio\bin\Release\Codaxy*.* Codaxy.Xlio\lib

..\.nuget\nuget pack Codaxy.Xlio\Codaxy.Xlio.nuspec /verbose

pause