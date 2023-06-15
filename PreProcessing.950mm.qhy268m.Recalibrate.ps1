if (-not (get-module psxisfreader)){import-module psxisfreader}

$ErrorActionPreference="STOP"
$WarningPreference="Continue"
$ArchiveDirectory="E:\Astrophotography"
$CalibratedOutput = "E:\Calibrated\950mm"
$RecalibratedOutput = "E:\Recalibrated\950mm"

$BiasLibraryFiles=Get-MasterBiasLibrary `
    -Path "E:\Astrophotography\BiasLibrary\QHY268M" `
    -Pattern "^(?<date>\d+).MasterBias.Gain.(?<gain>\d+).Offset.(?<offset>\d+).(?<numberOfExposures>\d+)x(?<exposure>\d+\.?\d*)s.xisf$" |
    where-object Geometry -eq "6252:4176:1"
$DarkLibraryFiles=Get-MasterDarkLibrary `
    -Path "E:\Astrophotography\DarkLibrary\QHY268M" `
    -Pattern "^(?<date>\d+).MasterDark.Gain.(?<gain>\d+).Offset.(?<offset>\d+).(?<temp>-?\d+)C.(?<numberOfExposures>\d+)x(?<exposure>\d+)s.xisf$" |
    where-object Geometry -eq "6252:4176:1"
$DarkLibrary=($DarkLibraryFiles|group-object Instrument,Gain,Offset,Exposure,SetTemp|foreach-object {
    $instrument=$_.Group[0].Instrument
    $gain=$_.Group[0].Gain
    $offset=$_.Group[0].Offset
    $exposure=$_.Group[0].Exposure
    $setTemp=$_.Group[0].SetTemp
    $geometry=$_.Group[0].Geometry
    
    $dark=$_.Group | sort-object {(Get-Item $_.Path).LastWriteTime} -Descending | select-object -First 1
    new-object psobject -Property @{
        Instrument=$instrument
        Gain=$gain
        Offset=$offset
        Exposure=$exposure
        SetTemp=$setTemp
        Geometry=$geometry
        Path=$dark.Path
    }
})

$FlatLibrary = Get-XisfFile -Path "$ArchiveDirectory\950mm\Flats" |
    Group-Object Instrument,Geometry,Filter,Telescope
$lightFrames=Get-Childitem -Path "$ArchiveDirectory\950mm" -Directory |
    foreach-object {
        Get-XisfLightFrames -Path $_.FullName `
            -SkipOnError `
            -Recurse `
            -UseCache `
            -PathTokensToIgnore @("reject","process","testing","clouds","draft","cloudy","_ez_LS_","drizzle","skyflats","quickedit","calibrated") 
    } |
    Where-Object {$_.LocalDate } |
    Where-Object {-not $_.IsIntegratedFile()} |
    where-object {-not [string]::IsNullOrWhiteSpace($_.Object)} |
    where-object Instrument -eq "QHY268M" |
    where-object ImageType -eq "LIGHT" |
    where-object FocalLength -eq "950" |
    where-object Gain -eq 26 |
    where-object Geometry -eq "6252:4176:1" |
    #where-object Filter -in @("L","R","G","B") |
    where-object Filter -eq "L" |
    Where-Object ObsDateMinus12hr -ge ([DateTime]"2023-03-01") |
    #where-object FocalRatio -eq 4 |
    where-object Object -eq "M51"
$lightFrames.Length
#$lightFrames|group-object {$_.ObsDateMinus12hr}


$toCalibrate=$lightFrames |
    group-object Instrument,Geometry,Filter,Telescope |
    foreach-object {
        $instrument=$_.Group[0].Instrument
        $geometry=$_.Group[0].Geometry
        $filter=$_.Group[0].Filter
        $matchingFlats = ($FlatLibrary|where-object Name -eq $_.Name).Group
        $mostRecentFlat = $matchingFlats | sort-object {(get-item $_.Path).LastWriteTime} -Descending | select-object -first 1

        $_.Group |
            group-object Exposure,Gain,Offset | 
            foreach-object {
                $exposure = $_.Group[0].Exposure
                $gain = $_.Group[0].Gain
                $offset = $_.Group[0].Offset

                $darkCandidates = $DarkLibrary |
                    where-object Instrument -eq $instrument |
                    where-object Geometry -eq $geometry |
                    where-object Exposure -eq $exposure |
                    where-object Gain -eq $gain |
                    where-object Offset -eq $offset
                $biasCandidates = $BiasLibraryFiles |
                    where-object Instrument -eq $instrument |
                    where-object Geometry -eq $geometry |
                    where-object Gain -eq $gain |
                    where-object Offset -eq $offset
                if(-not $darkCandidates){
                    Write-Warning "No suitable dark masters found for $instrument $geometry $($exposure)s Gain $gain Offset $offset"
                }

                $_.Group | foreach-object {
                    $lightFrame=$_
                    $bestDark = $darkCandidates |
                        where-object {[Math]::Abs($_.setTemp-$lightFrame.CCDTemp) -lt 4 } |
                        sort-object {$_.ObsDate} -Descending |
                        Select-Object -First 1
                    if(-not $bestDark){
                        $bestDark = $darkCandidates |
                        where-object {[Math]::Abs($_.setTemp-$lightFrame.CCDTemp) -lt 7 } |
                        sort-object {$_.ObsDate} -Descending |
                        Select-Object -First 1
                    }
                    if(-not $bestDark){
                        $bestDark = $darkCandidates |
                        where-object {[Math]::Abs($_.setTemp-$lightFrame.CCDTemp) -lt 12 } |
                        sort-object {$_.ObsDate} -Descending |
                        Select-Object -First 1
                    }
                    $bestBias = $biasCandidates |
                        sort-object {$_.ObsDate} -Descending |
                        select-object -First 1
                    new-object psobject -Property @{
                        LightFrameToCalibrate=$lightFrame
                        MasterDark=$bestDark
                        MasterBias=$bestBias
                        MasterFlat=$mostRecentFlat
                        Key="BIAS: $($bestBias.Name); DARK: $($bestDark.Name);"
                    }
                }
            }
    }
[IO.Directory]::CreateDirectory($RecalibratedOutput)>>$null
$toCalibrate|
    group-object {
        "Object: $($_.LightFrameToCalibrate.Object.Trim());BIAS: $($_.MasterBias.Path.Name); DARK: $($_.MasterDark.Path.Name); FLAT: $($_.MasterFlat.Path.Name);"
    } |
    foreach-object {
        $masterBias=$_.Group[0].MasterBias
        $masterFlat=$_.Group[0].MasterFlat
        $masterDark=$_.Group[0].MasterDark
        $object=$_.Group[0].LightFrameToCalibrate.Object.Trim()
        $lightFramesToCalibrate=$_.Group.LightFrameToCalibrate
        write-host "Recalibrating $($lightFramesToCalibrate.Count) lights for $object using"
        write-host " - Master Bias: $($masterBias.Path.Name)"
        write-host " - Master Dark: $($masterDark.Path.Name)"
        write-host " - Master Flat: $($masterFlat.Path.Name)"
        $calibrateDark=($masterDark -and $masterBias)
        $optimizeDark=($masterDark -and $masterBias)

        Invoke-LightFrameSorting -XisfStats $lightFramesToCalibrate `
            -DoNotArchive `
            -MasterBias $masterBias.Path `
            -MasterDark $masterDark.Path `
            -MasterFlat $masterFlat.Path `
            -CalibrateDark:$calibrateDark `
            -OptimizeDark:$optimizeDark `
            -PixInsightSlot 201 `
            -OutputPedestal 90 `
            -OutputPath $RecalibratedOutput 
    }




$lightFrames.Count
$calibrated=$lightFrames |
        Get-XisfCalibrationState `
            -CalibratedPath $CalibratedOutput `
            -AdditionalSearchPaths @() `
            -Verbose -Recurse -ShowProgress -ProgressTotalCount ($lightFrames.Count) |
        foreach-object {
            $x = $_
            if(-not $x.IsCalibrated()){
                Write-Warning "Skipping file: Uncalibrated: $($x.Path)"
            }
            else {
                $calibratedStats= Get-XisfFitsStats -Path $x.Calibrated
                $calibrationState=New-XisfPreprocessingState -Stats $calibratedStats
                new-object psobject -Property @{
                    LightFrame=$x.Stats
                    CalibrationState=$calibrationState
                }
            }
         } 
$calibrated.Count
$calibrated | 
    group-object {
        "FLAT: $($_.CalibrationState.MasterFlat)"
    } |
    foreach-object {$_.Name}
$lightFrames |
    Export-Csv -Path ".\Triage.90mm.Calibration.CSV"