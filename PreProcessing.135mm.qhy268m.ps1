if (-not (get-module psxisfreader)){import-module psxisfreader}

$ErrorActionPreference="STOP"
$WarningPreference="Continue"
$DropoffLocation = "D:\Backups\Camera\Dropoff\NINA"
$ArchiveDirectory="E:\Astrophotography"
$CalibratedOutput = "F:\PixInsightLT\Calibrated"
<#
Invoke-BiasFrameSorting `
    -DropoffLocation $DropoffLocation `
    -ArchiveDirectory $ArchiveDirectory `
    -PixInsightSlot 200 `
    -Verbose 
Invoke-DarkFrameSorting `
    -DropoffLocation $DropoffLocation `
    -ArchiveDirectory $ArchiveDirectory `
    -PixInsightSlot 200 `
    -Verbose 
Invoke-DarkFlatFrameSorting `
    -DropoffLocation $DropoffLocation `
    -ArchiveDirectory $ArchiveDirectory `
    -PixInsightSlot 200
Invoke-FlatFrameSorting `
    -DropoffLocation $DropoffLocation `
    -ArchiveDirectory $ArchiveDirectory `
    -CalibratedFlatsOutput "F:\PixInsightLT\CalibratedFlats" `
    -PixInsightSlot 200
exit
#>
#exit

Get-ChildItem $DropoffLocation *.xisf |
    Get-XisfFitsStats | 
    where-object Instrument -eq "QHY268M" |
    where-object ImageType -eq "LIGHT" |
    #where-object {$_.Rotator -gt 355 -or $_.Rotator -lt 5} | 
    #Where-Object SetTemp -eq -15 |
    #where-object Filter -eq "Oiii6nm" |
    where-object FocalLength -eq 1000 |
    where-object SetTemp -eq 0 |
    #where-object Object -eq "Abell 31" |
    where-object Exposure -EQ 180 |
    #select-object -First 3 |
    group-object Instrument,Gain,Exposure |
    foreach-object {
        $lights = $_.Group
        $lights |
            group-object Filter,FocalLength |
            foreach-object {
                $filter = $_.Group[0].Filter
                $focalLength=$_.Group[0].FocalLength

                Write-Host "Sorting $($_.Group.Count) frames at ($focalLength)mm with filter $filter"

                Invoke-LightFrameSorting `
                    -XisfStats ($_.Group) `
                    -ArchiveDirectory $ArchiveDirectory `
                    <#-MasterDark:$null#> `
                    <#-MasterBias:"E:\Astrophotography\BiasLibrary\QHYCCD-Cameras-Capture (ASCOM)\20210215.SuperBias.Gain.56.Offset.10.100x0.001s.xisf"#> `
                    -MasterFlat "E:\Astrophotography\$($focalLength)mm\Flats\20210712.MasterFlatCal.$($filter).xisf" `
                    <#-MasterDark "E:\Astrophotography\DarkLibrary\QHYCCD-Cameras-Capture (ASCOM)\20210217.MasterDark.Gain.56.Offset.10.-15C.54x60s.xisf"#> `
                    <#-MasterDark "E:\Astrophotography\DarkLibrary\QHYCCD-Cameras-Capture (ASCOM)\20210217.MasterDark.Gain.56.Offset.10.-15C.53x90s.xisf"#> `
                    <#-MasterDark "E:\Astrophotography\DarkLibrary\QHYCCD-Cameras-Capture (ASCOM)\20210217.MasterDark.Gain.56.Offset.10.-15C.53x120s.xisf"#> `
                    <#-MasterDark "E:\Astrophotography\DarkLibrary\QHYCCD-Cameras-Capture (ASCOM)\20210217.MasterDark.Gain.56.Offset.10.-15C.53x240s.xisf"#> `
                    <#-MasterDark "E:\Astrophotography\DarkLibrary\QHYCCD-Cameras-Capture (ASCOM)\20210228.MasterDark.Gain.56.Offset.0.-7C.88x180s.xisf"#> `
                    -MasterDark "E:\Astrophotography\DarkLibrary\QHYCCD-Cameras-Capture (ASCOM)\20210529.MasterDark.Gain.56.Offset.0.0C.48x180s.xisf" `                    
                    <#-MasterDark "E:\Astrophotography\DarkLibrary\QHYCCD-Cameras-Capture (ASCOM)\20210221.MasterDark.Gain.56.Offset.10.-15C.23x600s.xisf" #>`
                    <#-CalibrateDark -OptimizeDark#> `
                    -OutputPath $CalibratedOutput `
                    -PixInsightSlot 200 `
                    -OutputPedestal 70
            }
    }

Get-ChildItem $DropoffLocation *.xisf |
    Get-XisfFitsStats | 
    where-object Instrument -eq "QHYCCD-Cameras-Capture (ASCOM)" |
    where-object ImageType -eq "LIGHT" | 
    #Where-Object SetTemp -eq -15 |
    #where-object Object -eq "Alpha Cam" |
    where-object FocalLength -eq 1000 |
    #where-object SetTemp -eq -7 |
    where-object Exposure -EQ 180 |
    #where-object Filter -ne "Sii6nm" |
    #select-object -First 3 |
    group-object Instrument,Gain,Exposure |
    foreach-object {
        $lights = $_.Group
        $lights |
            group-object Filter,FocalLength |
            foreach-object {
                $filter = $_.Group[0].Filter
                $focalLength=$_.Group[0].FocalLength

                Write-Host "Sorting $($_.Group.Count) frames at ($focalLength)mm with filter $filter"

                Invoke-LightFrameSorting `
                    -XisfStats ($_.Group) `
                    -ArchiveDirectory $ArchiveDirectory `
                    <#-MasterDark:$null#> `
                    -MasterBias:"E:\Astrophotography\BiasLibrary\QHYCCD-Cameras-Capture (ASCOM)\20210215.SuperBias.Gain.56.Offset.10.100x0.001s.xisf" `
                    -MasterFlat "E:\Astrophotography\$($focalLength)mm\Flats\20210712.MasterFlatCal.$($filter).xisf" `
                    <#-MasterDark "E:\Astrophotography\DarkLibrary\QHYCCD-Cameras-Capture (ASCOM)\20210227.MasterDark.Gain.56.Offset.0.-7C.41x60s.xisf"#> `
                    <#-MasterDark "E:\Astrophotography\DarkLibrary\QHYCCD-Cameras-Capture (ASCOM)\20210227.MasterDark.Gain.56.Offset.0.-7C.41x90s.xisf"#> `
                    <#-MasterDark "E:\Astrophotography\DarkLibrary\QHYCCD-Cameras-Capture (ASCOM)\20210226.MasterDark.Gain.56.Offset.0.-7C.34x120s.xisf"#> `
                    -MasterDark "E:\Astrophotography\DarkLibrary\QHYCCD-Cameras-Capture (ASCOM)\20210228.MasterDark.Gain.56.Offset.0.-7C.88x180s.xisf" `
                    <#-MasterDark "E:\Astrophotography\DarkLibrary\QHYCCD-Cameras-Capture (ASCOM)\20210227.MasterDark.Gain.56.Offset.0.-7C.40x240s.xisf"#> `
                    <#-MasterDark "E:\Astrophotography\DarkLibrary\QHYCCD-Cameras-Capture (ASCOM)\20210226.MasterDark.Gain.56.Offset.0.-7C.34x360s.xisf"#> `
                    -CalibrateDark -OptimizeDark  `
                    -OutputPath $CalibratedOutput `
                    -PixInsightSlot 200 `
                    -OutputPedestal 70
            }
    }
