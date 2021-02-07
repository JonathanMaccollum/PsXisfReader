import-module $PSScriptRoot/PsXisfReader.psd1 -Force
$ErrorActionPreference="STOP"
$VerbosePreference="Continue"
$target="E:\Astrophotography\50mm\Taurus to Pleiades 1 50mm"

$DarkLibraryFiles = Get-MasterDarkLibrary `
    -Path "E:\Astrophotography\DarkLibrary\ZWO ASI183MM Pro" `
    -Pattern "^(?<date>\d+).MasterDark.Gain.(?<gain>\d+).Offset.(?<offset>\d+).(?<temp>-?\d+)C.(?<numberOfExposures>\d+)x(?<exposure>\d+)s.xisf$"
$FlatFiles = Get-ChildItem "E:\Astrophotography\50mm\Flats" -File *.xisf |
    Get-XisfFitsStats
$rawSubs =
    Get-XisfLightFrames -Path $target -Recurse -UseCache -SkipOnError |
    Where-Object {-not $_.HasTokensInPath(@("reject","process","testing","clouds","draft","cloudy","_ez_LS_"))} |
    Where-Object {-not $_.IsIntegratedFile()} 

$groups = $rawSubs|
    group-object Instrument,Filter,Exposure,SetTemp,FocalLength,Gain,Offset|
    foreach-object{
        new-object psobject -Property @{
            Instrument=$_.Group[0].Instrument
            Exposure=$_.Group[0].Exposure
            SetTemp=$_.Group[0].SetTemp
            Filter=$_.Group[0].Filter
            FocalLength=$_.Group[0].FocalLength
            Gain=$_.Group[0].Gain
            Offset=$_.Group[0].Offset
            Group=$_.Group
        }        
    }
foreach($group in $groups){
    $calibratedResults = $group.Group |
        Get-XisfCalibrationState -CalibratedPath "E:\PixInsightLT\Calibrated" | 
        foreach-object {
            $calibratedFile = $_.Calibrated | Get-XisfFitsStats
            $state=New-XisfPreprocessingState -Stats $calibratedFile 
        }


}
$calibrationState = $rawSubs |
    Get-XisfCalibrationState -CalibratedPath "E:\PixInsightLT\Calibrated" | 
    foreach-object {
        $calibratedFile = $_.Calibrated | Get-XisfFitsStats
        $state=New-XisfPreprocessingState -Stats $calibratedFile 
        $dark = $DarkLibraryFiles | where-object {$_.Path.Name -eq $state.MasterDark}
        $flat = $FlatFiles | where-object {$_.Path.Name -eq $state.MasterFlat}
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

        Write-Host "Re-calibrating $($files.Count) using flat master $($masterFlat.Name) and output pedestal $pedestal"
        Write-Host "Previous Dark: $($masterDark.Name)"
        $availableDarks = $DarkLibraryFiles |
            where-object {
                $dark=$_
                ($dark.Exposure -eq $exposure) -and 
                ($dark.Gain -eq $gain) -and 
                ($dark.Offset -eq $offset) -and 
                ($dark.SetTemp -eq $setTemp)
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
            Invoke-PiLightCalibration `
                -PixInsightSlot 200 `
                -Images ($files.Path) `
                -MasterDark ($latestDark.Path) `
                -MasterFlat $masterFlat `
                -OutputPath "E:\PixInsightLT\ReCalibrated" `
                -OutputPedestal 200 -KeepOpen
        }
    }
