if (-not (get-module psxisfreader)){import-module psxisfreader}
$ErrorActionPreference="STOP"
$VerbosePreference="Continue"
$DropoffLocation = "D:\Backups\Camera\Dropoff\NINA"
$ArchiveDirectory="E:\Astrophotography"
$CalibratedOutput = "F:\PixInsightLT\Calibrated"

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
#$DarkLibrary|FT
Get-ChildItem $DropoffLocation *.xisf |
    Get-XisfFitsStats | 
    where-object Instrument -eq "ZWO ASI183MM Pro" |
    where-object ImageType -eq "LIGHT" |
    where-object FocalLength -eq 40 |
    group-object Instrument,SetTemp,Gain,Offset,Exposure |
    foreach-object {
        $lights = $_.Group
        $x=$lights[0]

        $instrument=$x.Instrument
        $gain=[decimal]$x.Gain
        $offset=[decimal]$x.Offset
        $exposure=[decimal]$x.Exposure
        $ccdTemp=[decimal]$x.CCDTemp
        $setTemp=[decimal]$x.SetTemp
        $masterDark = $DarkLibrary | where-object {
            $dark = $_
            ($dark.Instrument-eq $instrument) -and
            ($dark.Gain-eq $gain) -and
            ($dark.Offset-eq $offset) -and
            ($dark.Exposure-eq $exposure) -and
            ($dark.SetTemp-eq $setTemp) -and
            ([Math]::abs($dark.SetTemp - $ccdTemp) -lt 3)
        } |
        select-object -first 1

        if(-not $masterDark){
            Write-Warning "Unable to process $($lights.Count) images: No master dark available for $instrument at Gain=$gain Offset=$offset Exposure=$exposure (s) and SetTemp $setTemp"
        }else {
            Write-Host "Master dark available for $instrument at Gain=$gain Offset=$offset Exposure=$exposure (s) and SetTemp $setTemp"
            $lights |
                group-object Filter,FocalLength |
                foreach-object {
                    $filter = $_.Group[0].Filter
                    $focalLength=$_.Group[0].FocalLength
                    $masterFlat = "E:\Astrophotography\$($focalLength)mm\Flats\20210714.MasterFlatCal.$filter.xisf"

                    if(-not (test-path $masterFlat)) {
                        Write-Warning "Skipping $($_.Group.Count) frames at ($focalLength)mm with filter $filter. Reason: No master flat was found."
                    }
                    else{

                        Write-Host "Sorting $($_.Group.Count) frames at ($focalLength)mm with filter $filter"
                        Write-Host " Dark: $($masterDark.Path)"
                        Write-Host " Flat: $($masterFlat)"

                        Invoke-LightFrameSorting `
                            -XisfStats ($_.Group) -ArchiveDirectory $ArchiveDirectory `
                            -MasterDark ($masterDark.Path) `
                            -MasterFlat $masterFlat `
                            -OutputPath $CalibratedOutput `
                            -PixInsightSlot 200 `
                            -OutputPedestal 200 `
                            -Verbose
                    }
                }
        }
    }