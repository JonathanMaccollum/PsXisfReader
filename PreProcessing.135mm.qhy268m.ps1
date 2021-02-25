import-module $PSScriptRoot/PsXisfReader.psd1 -Force

$ErrorActionPreference="STOP"
$WarningPreference="Continue"
$DropoffLocation = "D:\Backups\Camera\Dropoff\NINA"
$ArchiveDirectory="E:\Astrophotography"
$CalibratedOutput = "E:\PixInsightLT\Calibrated"
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
    -CalibratedFlatsOutput "E:\PixInsightLT\CalibratedFlats" `
    -PixInsightSlot 200
exit
#>
Get-ChildItem $DropoffLocation *.xisf |
    Get-XisfFitsStats | 
    where-object Instrument -eq "QHYCCD-Cameras-Capture (ASCOM)" |
    where-object ImageType -eq "LIGHT" |
    where-object Object -eq "IC1871" |
    where-object FocalLength -eq 1000 |
    where-object Exposure -EQ 360 |
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
                    -MasterFlat "E:\Astrophotography\$($focalLength)mm\Flats\20210223.MasterFlatCal.$($filter).xisf" `
                    <#-MasterDark "E:\Astrophotography\DarkLibrary\QHYCCD-Cameras-Capture (ASCOM)\20210218.MasterDark.Gain.56.Offset.10.-15C.52x360s.xisf" #> `
                    -MasterDark "E:\Astrophotography\DarkLibrary\QHYCCD-Cameras-Capture (ASCOM)\20210218.MasterDark.Gain.56.Offset.10.-15C.52x360s.xisf" `
                    <#-CalibrateDark -OptimizeDark#>  `
                    -OutputPath $CalibratedOutput `
                    -PixInsightSlot 200 `
                    -OutputPedestal 70
            }
    }

Get-ChildItem $DropoffLocation *.xisf |
    Get-XisfFitsStats | 
    where-object Instrument -eq "QHYCCD-Cameras-Capture (ASCOM)" |
    where-object ImageType -eq "LIGHT" |
    #where-object Object -eq "Seagull" |
    where-object FocalLength -eq 135 |
    where-object Exposure -eq 360 |
    group-object Instrument,Gain,Exposure |
    foreach-object {
        $lights = $_.Group
        $lights |
            group-object Filter,FocalLength |
            foreach-object {
                $filter = $_.Group[0].Filter
                $focalLength=$_.Group[0].FocalLength

                if(-not (test-path $masterFlat)) {
                    Write-Warning "Skipping $($_.Group.Count) frames at ($focalLength)mm with filter $filter. Reason: No master flat was found."
                }
                else{
                    Write-Host "Sorting $($_.Group.Count) frames at ($focalLength)mm with filter $filter"

                    Invoke-LightFrameSorting `
                        -XisfStats ($_.Group) `
                        -ArchiveDirectory $ArchiveDirectory `
                        <#-MasterDark:$null#> `
                        <#-MasterBias:"E:\Astrophotography\BiasLibrary\QHYCCD-Cameras-Capture (ASCOM)\20210215.SuperBias.Gain.56.Offset.10.100x0.001s.xisf" #> `
                        -MasterFlat "E:\Astrophotography\$($focalLength)mm\Flats\20210220.MasterFlatCal.$filter.xisf" `
                        -MasterDark "E:\Astrophotography\DarkLibrary\QHYCCD-Cameras-Capture (ASCOM)\20210218.MasterDark.Gain.56.Offset.10.-15C.52x360s.xisf" `
                        <#-CalibrateDark -OptimizeDark #> `
                        -OutputPath $CalibratedOutput `
                        -PixInsightSlot 200 `
                        -OutputPedestal 70 -KeepOpen
                }
            }
    }