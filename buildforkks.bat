@echo off

set CSC="C:\Windows\Microsoft.NET\Framework\v3.5\csc.exe"

set GAME=D:\Games\KoikatsuSunshine
set MANAGED=%GAME%\KoikatsuSunshine_Data\Managed

set SRC=KKS_RandomPngToDragAndDrop.cs
set OUT=%GAME%\BepInEx\plugins\KKS_dragdrop_random_png.dll

%CSC% /noconfig /target:library /optimize+ /warn:4 /warnaserror- ^
 /r:%MANAGED%\UnityEngine.dll ^
 /r:%GAME%\BepInEx\core\BepInEx.dll ^
 /r:%MANAGED%\UnityEngine.CoreModule.dll ^
 /r:%MANAGED%\UnityEngine.InputLegacyModule.dll ^
 /r:%MANAGED%\UnityEngine.UI.dll ^
 /r:%MANAGED%\Assembly-CSharp.dll ^
 /r:%GAME%\BepInEx\core\0Harmony.dll ^
 /out:%OUT% ^
 %SRC%

if errorlevel 1 (
  echo.
  echo BUILD FAILED
  pause
  exit /b 1
)

echo.
echo Build OK:
echo %OUT%
pause
