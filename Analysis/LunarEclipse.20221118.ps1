if (-not (get-module psxisfreader)){import-module psxisfreader}

<#
$stats = Get-ChildItem D:\Backups\Camera\Dropoff\NINA -Filter "Lunar*" |
    Get-XisfFitsStats

$stats|
group-object Object,Filter,Exposure|
foreach-object {
    $object = $_.Group[0].Object
    $filter = $_.Group[0].Filter
    $exposure = $_.Group[0].Exposure
    $targetDirectory = "E:\Astrophotography\135mm\$object\$filter\$($exposure)s"
    [IO.Directory]::CreateDirectory($targetDirectory)
    $_.Group | foreach-object {
        $x=$_
        $fileName = [IO.Path]::GetFileName($x.Path)
        move-item -Path $_.Path -Destination "$targetDirectory/$fileName"
    }
}
#>
Clear-Content -Path "E:\Astrophotography\Test.scp"
$exposure = 0.0012
$r = Get-ChildItem "E:\Astrophotography\135mm\Lunar Eclipse and Pleiades\R\$($exposure)s" -Filter "Lunar*" | Get-XisfFitsStats
$g = Get-ChildItem "E:\Astrophotography\135mm\Lunar Eclipse and Pleiades\G\$($exposure)s" -Filter "Lunar*" | Get-XisfFitsStats
$b = Get-ChildItem "E:\Astrophotography\135mm\Lunar Eclipse and Pleiades\B\$($exposure)s" -Filter "Lunar*" | Get-XisfFitsStats

$rgbStats = $g | ForEach-Object {
    $x = $_
    $matchingRed = $r | sort-object {
        [Math]::Abs(($_.ObsDate - $x.ObsDate).TotalSeconds)
    } | Select-Object -First 1
    $matchingBlue = $b | sort-object {
        [Math]::Abs(($_.ObsDate - $x.ObsDate).TotalSeconds)
    } | Select-Object -First 1

    new-object psobject -Property @{
        LocalDate = $x.LocalDate
        Red = $matchingRed
        Green = $x
        Blue = $matchingBlue
    }
}


$rgbStats | 
where-object LocalDate -gt "2021-11-19 4:02:00 AM" |
where-object LocalDate -lt "2021-11-19 4:05:00 AM" |
ForEach-Object {
    $r1 = $_.Red.Path | Format-PiPath
    $g1 = $_.Green.Path | Format-PiPath
    $b1 = $_.Blue.Path | Format-PiPath
    $outpuFile = [IO.Path]::GetFileNameWithoutExtension($_.Green.Path)+".tiff"
    $outputPath = Format-PiPath -Path "S:\PixInsight\Timelapse\RGB\$outpuFile"
    Add-Content -Path "E:\Astrophotography\Test.scp" -Value "
open `"$r1`"
ImageIdentifier -id=`"R`"
open `"$g1`"
ImageIdentifier -id=`"G`"
open `"$b1`"
ImageIdentifier -id=`"B`"
PixelMath -n -rgb -x0=R -x1=G -x2=B --interface -g-
ImageIdentifier -id=`"RGB`"
SCNR -green -a=0.8
save RGB -p=`"$outputPath`" --nodialog --nomessages --noverify
close --force R
close --force G
close --force B
close --force RGB
    "
}
