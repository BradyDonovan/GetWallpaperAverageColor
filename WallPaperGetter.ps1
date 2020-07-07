[CmdletBinding(DefaultParameterSetName = 'Help')]
param (
  [Parameter(ParameterSetName = 'Help')]
  [Switch]
  $Help # not doing anything with this on purpose; I just needed CmdletBinding available to use -Verbose.
)
$sLib = @'
using System;
using System.Runtime.InteropServices;
using System.Text;

namespace Win32{

    public class Wallpaper{

      [DllImport("user32.dll", CharSet=CharSet.Auto)]
      static extern int SystemParametersInfo (uint uiAction, uint uiParam, StringBuilder pvParam, int fWinIni);

      public static string GetWallpaperPath() {
        const uint SPI_GETDESKWALLPAPER = 0x0073;
        StringBuilder sb = new StringBuilder(500);
        SystemParametersInfo(SPI_GETDESKWALLPAPER,(uint)sb.Capacity,sb,0);
        return sb.ToString();
      }
    }
}
'@
Add-Type $sLib
Add-Type -AssemblyName System.Drawing

function Add-ToColorStatsTable {
  param (
    [AllowNull()]
    [int]$R,
    [AllowNull()]
    [int]$G,
    [AllowNull()]
    [int]$B,
    [int]$TotalAdds = 1,
    [PSCustomObject]$Table
  )
  process {
    $Table.R = $Table.R + $R
    $Table.G = $Table.G + $G
    $Table.B = $Table.B + $B
    $Table.TotalAdds = $Table.TotalAdds + $TotalAdds

    $Table
  }
}
function Get-AverageColorFromStatsTable {
  param (
    [PSCustomObject]$Table
  )
  process {
    [int]$Table.R = $Table.R / $Table.TotalAdds
    [int]$Table.G = $Table.G / $Table.TotalAdds
    [int]$Table.B = $Table.B / $Table.TotalAdds

    $Table
  }
}
$oRGBColorStats = [PSCustomObject]@{
  R         = 0
  G         = 0
  B         = 0
  TotalAdds = 0
}

# retrieve the wallpaper path
$sWpPath = [Win32.Wallpaper]::GetWallpaperPath()

IF ($null -ne $sWpPath) {
  [System.Drawing.Image]$image = [System.Drawing.Image]::FromFile($sWpPath)
  $totalXpixels = $image.Size.Width
  $totalYpixels = $image.Size.Height
  $y = 0

  # get X RGB values
  while ($y -ne $totalYpixels) {
    # display a progress bar aligned with current Y value as a measurement of progress
    [decimal]$dProgressInDec = $y/$totalYpixels
    $progress = [decimal]::Truncate($dProgressInDec * 100) / 100
    Write-Progress -Activity "Processing in progress" -Status "$progress% Complete:" -PercentComplete $progress

    # revert the processing bit
    $bSkipProcessingBit = $null

    # x needs to be reset after it reaches the maximum of the x axis; we want to start back at (0, Y)
    $x = 0

    # begin the processing loop
    do {
      Write-Verbose "Processing pixel: [$($x), $($y)]"

      # retrieve the current pixel in the bitmap
      $pixel = $image.GetPixel($x, $y)

      # declare an empty hashtable to hold the highest key value pair in case there's more than one
      $oHighestKey = @{}

      # create subset hashtable with just RGB values
      $oCurrentPixelRGB = @{
        R = [int]$pixel.R
        G = [int]$pixel.G
        B = [int]$pixel.B
      }

      # is the current pixel solid black or white? Set skipProcessingBit if either.
      IF ($oCurrentPixelRGB.R -eq 0 -and $oCurrentPixelRGB.G -eq 0 -and $oCurrentPixelRGB.B -eq 0) {
        Write-Verbose "Pixel [$($x), $($y)] is black."
        $bSkipProcessingBit = $true

        $oRGBColorStats = Add-ToColorStatsTable -R 0 -G 0 -B 0 -Table $oRGBColorStats
      }
      IF ($oCurrentPixelRGB.R -eq 255 -and $oCurrentPixelRGB.G -eq 255 -and $oCurrentPixelRGB.B -eq 255) {
        Write-Verbose "Pixel [$($x), $($y)] is white."
        $bSkipProcessingBit = $true

        $oRGBColorStats = Add-ToColorStatsTable -R 255 -G 255 -B 255 -Table $oRGBColorStats
      }

      # don't process the pixel if it's solidly white or black
      IF ($null -eq $bSkipProcessingBit) {
        # declare an empty array the size of all values in the hashtable
        [int[]]$aDstArray = @(0) * $oCurrentPixelRGB.Values.Count

        # copy the RGB values to the destination array. This is needed to use setup for using Linq's .Max()
        $oCurrentPixelRGB.Values.CopyTo($aDstArray, 0)
        [System.Collections.Generic.IEnumerable[int]]$iEnum = $aDstArray

        # grab the highest value from the array
        $iHighestVal = [System.Linq.Enumerable]::Max($iEnum)

        # recurse the dictionary looking for the key associated with the highest value found
        foreach ($key in $oCurrentPixelRGB.GetEnumerator()) {
          IF ($key.Value -eq $iHighestVal) {
            Write-Verbose "Found the strongest RGB color value for pixel [$($x), $($y)]. It's [$($key.Key) : $($key.Value)]."
            $oHighestKey.Add($key.Key, $key.Value)

            # add the current strongest colors to the $oRGBColorStats table
            switch ($key.Key) {
              "R" { $oRGBColorStats = Add-ToColorStatsTable -R $key.Value -Table $oRGBColorStats }
              "G" { $oRGBColorStats = Add-ToColorStatsTable -G $key.Value -Table $oRGBColorStats }
              "B" { $oRGBColorStats = Add-ToColorStatsTable -B $key.Value -Table $oRGBColorStats }
            }
          }
        }
      }

      # now that processing has finished, increment the x axis
      $x++
    }
    while ($x -ne $totalXpixels)

    # increment the y axis once x has incremented to its maximum
    $y++
  }
  Get-AverageColorFromStatsTable -Table $oRGBColorStats
}