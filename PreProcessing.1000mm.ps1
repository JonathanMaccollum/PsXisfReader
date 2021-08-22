import-module $PSScriptRoot/PsXisfReader.psd1 -Force
$ErrorActionPreference="STOP"

$DropoffLocation = "D:\Backups\Camera\Dropoff\NINA"
$ArchiveDirectory="E:\Astrophotography"
$CalibratedOutput = "F:\PixInsightLT\Calibrated"

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



$DarkLibraryFiles = Get-MasterDarkLibrary `
    -Path "E:\Astrophotography\DarkLibrary\ZWO ASI183MM Pro" `
    -Pattern "^(?<date>\d+).MasterDark.Gain.(?<gain>\d+).Offset.(?<offset>\d+).(?<temp>-?\d+)C.(?<numberOfExposures>\d+)x(?<exposure>\d+)s.xisf$"
$DarkLibrary=($DarkLibraryFiles|group-object Instrument,Gain,Offset,Exposure,SetTemp|foreach-object {
    $instrument=$_.Group[0].Instrument
    $gain=$_.Group[0].Gain
    $offset=$_.Group[0].Offset
    $exposure=$_.Group[0].Exposure
    $setTemp=$_.Group[0].SetTemp
    
    $dark=$_.Group | sort-object {(Get-Item $_.Path).LastWriteTime} -Descending | select-object -First 1
    new-object psobject -Property @{
        Instrument=$instrument
        Gain=$gain
        Offset=$offset
        Exposure=$exposure
        SetTemp=$setTemp
        Path=$dark.Path
    }
})

Get-ChildItem $DropoffLocation *.xisf |
    Get-XisfFitsStats | 
    where-object Instrument -eq "ZWO ASI183MM Pro" |
    where-object ImageType -eq "LIGHT" |
    group-object Instrument,SetTemp,Gain,Offset,Exposure |
    foreach-object {
        $lights = $_.Group
        $x=$lights[0]

        $instrument=$x.Instrument
        $gain=[decimal]$x.Gain
        $offset=[decimal]$x.Offset
        $exposure=[decimal]$x.Exposure
        $setTemp=[decimal]$x.SetTemp
        $masterDark = $DarkLibrary | where-object {
            $dark = $_
            ($dark.Instrument-eq $instrument) -and
            ($dark.Gain-eq $gain) -and
            ($dark.Offset-eq $offset) -and
            ($dark.Exposure-eq $exposure) -and
            ($dark.SetTemp-eq $setTemp)
        } | select-object -first 1

        if(-not $masterDark){
            Write-Warning "Unable to process $($lights.Count) images: No master dark available for $instrument at Gain=$gain Offset=$offset Exposure=$exposure (s) and SetTemp $setTemp"
        }else {
            Write-Host "Master dark available for $instrument at Gain=$gain Offset=$offset Exposure=$exposure (s) and SetTemp $setTemp"
            $lights |
                group-object Filter,FocalLength |
                foreach-object {
                    $filter = $_.Group[0].Filter
                    $focalLength=$_.Group[0].FocalLength
                    $masterFlat = "E:\Astrophotography\$($focalLength)mm\Flats\20201223.MasterFlatCal.$filter.xisf"

                    if(-not (test-path $masterFlat)) {
                        Write-Warning "Calibrating $($_.Group.Count) frames at ($focalLength)mm with filter $filter without flats. Reason: No master flat was found."
                    }

                        Write-Host "Sorting $($_.Group.Count) frames at ($focalLength)mm with filter $filter"

                        Invoke-LightFrameSorting `
                            -XisfStats ($_.Group) -ArchiveDirectory $ArchiveDirectory `
                            -MasterDark ($masterDark.Path) `
                            -MasterFlat $masterFlat `
                            -OutputPath $CalibratedOutput `
                            -PixInsightSlot 200 `
                            -OutputPedestal 900
                    
                }
        }
    }


exit
$('L','R','G','B')|foreach-object {
    Invoke-LightFrameSorting `
    -DropoffLocation $DropoffLocation -ArchiveDirectory $ArchiveDirectory `
    -MasterDark "E:\Astrophotography\DarkLibrary\ZWO ASI183MM Pro\20200422.MasterDark.Gain.111.Offset.8.-15C.60x15s.xisf" `
    -MasterFlat "E:\Astrophotography\1000mm\Flats\20200324.MasterFlatCal.$_.xisf" `
    -Filter $_ -FocalLength 1000 -Exposure 15 -Gain 111 -Offset 8 -SetTemp -15  `
    -OutputPath $CalibratedOutput `
    -PixInsightSlot 200 `
    -Verbose
}

$('L','R','G','B')|foreach-object {
    Invoke-LightFrameSorting `
    -DropoffLocation $DropoffLocation -ArchiveDirectory $ArchiveDirectory `
    -MasterDark "E:\Astrophotography\DarkLibrary\ZWO ASI183MM Pro\20200422.MasterDark.Gain.111.Offset.8.-15C.60x30s.xisf" `
    -MasterFlat "E:\Astrophotography\1000mm\Flats\20200324.MasterFlatCal.$_.xisf" `
    -Filter $_ -FocalLength 1000 -Exposure 30 -Gain 111 -Offset 8 -SetTemp -15  `
    -OutputPath $CalibratedOutput `
    -PixInsightSlot 200 `
    -Verbose
}
$('L','R','G','B')|foreach-object {
    Invoke-LightFrameSorting `
    -DropoffLocation $DropoffLocation -ArchiveDirectory $ArchiveDirectory `
    -MasterDark "E:\Astrophotography\DarkLibrary\ZWO ASI183MM Pro\20200329.MasterDark.Gain.111.Offset.8.-5C.60x60s.xisf" `
    -MasterFlat "E:\Astrophotography\1000mm\Flats\20200324.MasterFlatCal.$_.xisf" `
    -Filter $_ -FocalLength 1000 -Exposure 60 -Gain 111 -Offset 8 -SetTemp -5  `
    -OutputPath $CalibratedOutput `
    -PixInsightSlot 200 `
    -Verbose
}
$('L','R','G','B')|foreach-object {
    Invoke-LightFrameSorting `
    -DropoffLocation $DropoffLocation -ArchiveDirectory $ArchiveDirectory `
    -MasterDark "D:\Backups\Camera\ASI183mm-Pro\20190525\60x30s.-10C.G111.O8.masterdark.nocalibration.xisf" `
    -MasterFlat "E:\Astrophotography\1000mm\Flats\20200324.MasterFlatCal.$_.xisf" `
    -Filter $_ -FocalLength 1000 -Exposure 30 -Gain 111 -Offset 8 -SetTemp -10  `
    -OutputPath $CalibratedOutput `
    -PixInsightSlot 200
}
$('L','R','G','B')|foreach-object {
    Invoke-LightFrameSorting `
    -DropoffLocation $DropoffLocation -ArchiveDirectory $ArchiveDirectory `
    -MasterDark "E:\Astrophotography\DarkLibrary\ZWO ASI183MM Pro\20200422.MasterDark.Gain.111.Offset.8.-15C.45x60s.xisf" `
    -MasterFlat "E:\Astrophotography\1000mm\Flats\20200418.MasterFlatCal.$_.xisf" `
    -Filter $_ -FocalLength 1000 -Exposure 60 -Gain 111 -Offset 8 -SetTemp -15 -OutputPedestal 150 `
    -OutputPath $CalibratedOutput `
    -PixInsightSlot 200
}
$('L','R','G','B')|foreach-object {
    Invoke-LightFrameSorting `
        -DropoffLocation $DropoffLocation -ArchiveDirectory $ArchiveDirectory `
        -MasterDark "D:\Backups\Camera\ASI183mm-Pro\20191220\30x120s.-15C.G111.O8.masterdark.nocalibration.xisf" `
        -MasterFlat "E:\Astrophotography\1000mm\Flats\20200418.MasterFlatCal.$_.xisf" `
        -Filter $_ -FocalLength 1000 -Exposure 120 -Gain 111 -Offset 8 -SetTemp -15 -OutputPedestal 150  `
        -OutputPath $CalibratedOutput `
        -PixInsightSlot 200
}
$('Ha','Oiii','Sii','L','R','G','B')|foreach-object {
    Invoke-LightFrameSorting `
        -DropoffLocation $DropoffLocation -ArchiveDirectory $ArchiveDirectory `
        -MasterDark "E:\Astrophotography\DarkLibrary\ZWO ASI183MM Pro\20200422.MasterDark.Gain.111.Offset.8.-15C.45x240s.xisf" `
        -MasterFlat "E:\Astrophotography\1000mm\Flats\20200418.MasterFlatCal.$_.xisf" `
        -Filter $_ -FocalLength 1000 -Exposure 240 -Gain 111 -Offset 8 -SetTemp -15 -OutputPedestal 150  `
        -OutputPath $CalibratedOutput `
        -PixInsightSlot 200
}
$('Ha','Oiii','Sii')|foreach-object {
    Invoke-LightFrameSorting `
        -DropoffLocation $DropoffLocation -ArchiveDirectory $ArchiveDirectory `
        -MasterDark "D:\Backups\Camera\ASI183mm-Pro\20191220\74x600s.-15C.G53.O10.masterdark.nocalibration.xisf" `
        -MasterFlat "E:\Astrophotography\1000mm\Flats\20200418.MasterFlatCal.$_.xisf" `
        -Filter $_ -FocalLength 1000 -Exposure 600 -Gain 53 -Offset 10 -SetTemp -15  `
        -OutputPath $CalibratedOutput `
        -PixInsightSlot 200
}