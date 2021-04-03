[CmdletBinding()]
param (
  [Parameter()]
  [ValidateScript( {
      (Test-Path -PathType Leaf -Path $_) -eq $true
    })]
  [String]$Path
)

$Win32WallpaperClass = @'
using System.Runtime.InteropServices;

namespace Win32{

    public class Wallpaper{

      [DllImport("user32.dll", CharSet=CharSet.Auto)]
      static extern int SystemParametersInfo (int uAction , int uParam , string lpvParam , int fuWinIni);

      public static void SetWallpaper(string thePath){
         SystemParametersInfo(20,0,thePath,3);
      }
    }
}
'@
Add-Type $Win32WallpaperClass

[Win32.Wallpaper]::SetWallpaper($Path)