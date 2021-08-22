#import-module $PSScriptRoot/PsXisfReader.psd1 -Force
if (-not (get-module psxisfreader)){import-module psxisfreader}
$ErrorActionPreference="STOP"

$DropoffLocation = "D:\Backups\Camera\Dropoff\NINACS"
$ArchiveDirectory="E:\Astrophotography"
$CalibratedOutput = "F:\PixInsightLT\Calibrated"

Invoke-DarkFrameSorting `
    -DropoffLocation $DropoffLocation `
    -ArchiveDirectory $ArchiveDirectory `
    -PixInsightSlot 201 `
    -Verbose -KeepOpen
Invoke-DarkFlatFrameSorting `
    -DropoffLocation $DropoffLocation `
    -ArchiveDirectory $ArchiveDirectory `
    -PixInsightSlot 201
Invoke-FlatFrameSorting `
    -DropoffLocation $DropoffLocation `
    -ArchiveDirectory $ArchiveDirectory `
    -CalibratedFlatsOutput "F:\PixInsightLT\CalibratedFlats" `
    -PixInsightSlot 201

$DarkLibraryFiles=Get-MasterDarkLibrary `
    -Path "E:\Astrophotography\DarkLibrary\ZWO ASI071MC Pro" `
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
    where-object Instrument -eq "ZWO ASI071MC Pro" |
    where-object ImageType -eq "LIGHT" |
    where-object FocalLength -eq "40" |
    group-object Instrument,SetTemp,Gain,Offset,Exposure |
    foreach-object {
        $lights = $_.Group
        $x=$lights[0]

        $instrument=$x.Instrument
        $gain=[decimal]$x.Gain
        $offset=[decimal]$x.Offset
        $exposure=[decimal]$x.Exposure
        $ccdTemp = [decimal]$x.CCDTemp
        $setTemp=[decimal]$x.SetTemp
        $masterDark = $DarkLibrary | where-object {
            $dark = $_
            ($dark.Instrument-eq $instrument) -and
            ($dark.Gain-eq $gain) -and
            ($dark.Offset-eq $offset) -and
            ($dark.Exposure-eq $exposure) -and
            #($dark.SetTemp -eq $setTemp)
            ([Math]::abs($dark.SetTemp - $ccdTemp) -lt 3)
        } | select-object -first 1

        
        # if(-not $masterDark){
        #     $masterDark = $DarkLibrary | where-object {$_.Path.FullName -eq "E:\Astrophotography\DarkLibrary\ZWO ASI071MC Pro\20200517.MasterDark.Gain.90.Offset.65.-10C.110x120s.xisf"}
        # }


        if(-not $masterDark){
            Write-Warning "Unable to process $($lights.Count) images: No master dark available for $instrument at Gain=$gain Offset=$offset Exposure=$exposure (s) and SetTemp $setTemp"
        }else {
            Write-Host "Master dark available for $instrument at Gain=$gain Offset=$offset Exposure=$exposure (s) and SetTemp $setTemp"
            $lights |
                group-object Filter,FocalLength |
                foreach-object {
                    $filter = $_.Group[0].Filter
                    $focalLength=$_.Group[0].FocalLength
                    $masterFlat = "E:\Astrophotography\$($focalLength)mm\Flats\20210530.MasterFlatCal.F4.$filter.xisf"

                    if(-not (test-path $masterFlat)) {
                        Write-Warning "Skipping $($_.Group.Count) frames at ($focalLength)mm with filter $filter. Reason: No master flat was found."
                    }
                    else{

                        Write-Host "Sorting $($_.Group.Count) frames at ($focalLength)mm with filter $filter"

                        Invoke-LightFrameSorting `
                            -XisfStats ($_.Group) -ArchiveDirectory $ArchiveDirectory `
                            -MasterDark ($masterDark.Path) `
                            -MasterFlat $masterFlat `
                            -OutputPath $CalibratedOutput `
                            -PixInsightSlot 201 `
                            -OutputPedestal 50 `
                            -Verbose
                    }
                }
        }
    }

