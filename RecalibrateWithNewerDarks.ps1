import-module $PSScriptRoot/PsXisfReader.psd1 -Force
$ErrorActionPreference="STOP"
$VerbosePreference="Continue"
$target="E:\Astrophotography\1000mm\Heart Nebula"
$outputPath = "E:\PixInsightLT\ReCalibrated"
$WhatIf=$false
$DarkLibraryFiles = Get-MasterDarkLibrary `
    -Path "E:\Astrophotography\DarkLibrary\ZWO ASI183MM Pro" `
    -Pattern "^(?<date>\d+).MasterDark.Gain.(?<gain>\d+).Offset.(?<offset>\d+).(?<temp>-?\d+)C.(?<numberOfExposures>\d+)x(?<exposure>\d+)s.xisf$"
$FlatFiles = 
    @(Get-ChildItem "D:\Backups\Camera\2019\20191027.Flats.Newt.Efw" -File *.xisf) +
    @(Get-ChildItem "D:\Backups\Camera\2019\20190929.Flats.Newt.Efw" -File *.xisf) |
    Get-XisfFitsStats
$rawSubs =
    Get-XisfLightFrames -Path $target -Recurse -UseCache -SkipOnError |
    Where-Object {-not $_.HasTokensInPath(@("reject","process","testing","clouds","draft","cloudy","_ez_LS_"))} |
    Where-Object {-not $_.IsIntegratedFile()}  |
    Where-Object Object -eq "Melotte 15 Panel 2"

$calibrationState = $rawSubs |
    Get-XisfCalibrationState -CalibratedPath "E:\PixInsightLT\Calibrated\Melotte 15 Panel 2\Old" | 
    foreach-object {
        $calibratedFile = $_.Calibrated | Get-XisfFitsStats
        $state=New-XisfPreprocessingState -Stats $calibratedFile 
        $dark = $DarkLibraryFiles | where-object {$_.Path.Name -eq $state.MasterDark}
        $flat = $FlatFiles | where-object {$_.Path.Name -eq $state.MasterFlat}

        if(-not $dark){
            Write-Warning "Unable to locate dark $($state.MasterDark)"
        }
        if(-not $flat){
            Write-Warning "Unable to locate flat $($state.MasterFlat)"
        }

        new-object psobject -Property @{
            Path=$_.Path
            MasterDark = $dark.Path
            MasterFlat = $flat.Path
            Pedestal = $calibratedFile.Pedestal
            Gain=$calibratedFile.Gain
            Offset=$calibratedFile.Offset
            SetTemp=$calibratedFile.SetTemp
            Exposure=$calibratedFile.Exposure
        }
    }
$calibrationState|group-object MasterFlat,MasterDark,Pedestal,Gain,Offset,SetTemp,Exposure |
    foreach-object {
        $x=$_.Group[0]
        $masterFlat=$x.MasterFlat
        $masterDark=$x.MasterDark
        $pedestal=$x.Pedestal
        $gain=$x.Gain
        $offset=$x.Offset
        $setTemp=$x.SetTemp
        $exposure=$x.Exposure
        $files = $_.Group

        $availableDarks = $DarkLibraryFiles |
            where-object {
                $dark=$_
                ($dark.Exposure -eq $exposure) -and 
                ($dark.Gain -eq $gain) -and 
                ($dark.Offset -eq $offset) -and 
                #($dark.SetTemp -eq $setTemp) -and
                1-eq 1
            }
        $availableDarks|ForEach-Object {
            Write-Host "Available Dark: $($_.Path)"
        }
        $latestDark = $availableDarks | sort-object ObsDate -desc | select-object Path -first 1
        if(-not $latestDark){
            Write-Error "No darks currently available."
        }
        elseif($masterDark -eq $latestDark.Path){
            Write-Host "Current data is already calibrated with the latest dark: $($latestDark.Path)"
        }
        else{
            Write-Host "Re-calibrating $($files.Count) using flat master $($masterFlat.Name) and output pedestal $pedestal"
            Write-Host "Previous Dark: $($masterDark.Name)"
            Write-Host "Updated  Dark: $($latestDark.Path.Name)"
            if(-not $WhatIf){
                Invoke-PiLightCalibration `
                -PixInsightSlot 200 `
                -Images ($files.Path) `
                -MasterDark ($latestDark.Path) `
                -MasterFlat $masterFlat `
                -OutputPath $outputPath `
                -OutputPedestal 200 
            }else{
                Write-Host "What-If was set... no calibration was performed."
            }
        }
    }
