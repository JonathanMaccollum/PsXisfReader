if (-not (get-module psxisfreader)){import-module psxisfreader}

$ErrorActionPreference="STOP"
$WarningPreference="Continue"
$DropoffLocation = "D:\Backups\Camera\Dropoff\NINACS\All-Sky"
$ArchiveDirectory="E:\Astrophotography"
$CalibratedOutput = "S:\PixInsight\Timelapse\Calibrated"

<#
Invoke-BiasFrameSorting `
    -DropoffLocation $DropoffLocation `
    -ArchiveDirectory $ArchiveDirectory `
    -PixInsightSlot 201 `
    -Verbose 
Invoke-DarkFrameSorting `
    -DropoffLocation $DropoffLocation `
    -ArchiveDirectory $ArchiveDirectory `
    -PixInsightSlot 201 `
    -Verbose 
Invoke-DarkFlatFrameSorting `
    -DropoffLocation $DropoffLocation `
    -ArchiveDirectory $ArchiveDirectory `
    -PixInsightSlot 201
Invoke-FlatFrameSorting `
    -DropoffLocation $DropoffLocation `
    -ArchiveDirectory $ArchiveDirectory `
    -CalibratedFlatsOutput "F:\PixInsightLT\CalibratedFlats" `
    -PixInsightSlot 201
exit
#>
#exit
[IO.Directory]::CreateDirectory("E:\Astrophotography\DarkLibrary\ZWO ASI678MC")>>$null
$BiasLibraryFiles=Get-MasterBiasLibrary `
    -Path "E:\Astrophotography\BiasLibrary\ZWO ASI678MC" `
    -Pattern "^(?<date>\d+).MasterBias.Gain.(?<gain>\d+).Offset.(?<offset>\d+).(?<numberOfExposures>\d+)x(?<exposure>\d+\.?\d*)s.xisf$"
$DarkLibraryFiles=Get-MasterDarkLibrary `
    -Path "E:\Astrophotography\DarkLibrary\ZWO ASI678MC" `
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

while($true){
    $toCalibrate=Get-ChildItem $DropoffLocation/"2024-05-12" *.xisf -ErrorAction Continue |
        foreach-object { try{ $_ | Get-XisfFitsStats -ErrorAction Continue}catch{} }|
        where-object Instrument -eq "ZWO ASI678MC" |
        where-object ImageType -eq "LIGHT" |
        #where-object Object -eq "All-Sky 20230831 Ride along" |
        #where-object FocalLength -eq "135" |
        #where-object Offset -eq 65 |
        #where-object Object -eq "m101 at 135mm P1" |
        #where-object Filter -eq "L" |
        #select-object -first 50 |
        group-object ObsDateMinus12hr,Instrument,SetTemp,Gain,Offset,Exposure 
    $toCalibrate |
        foreach-object {
            $lights = $_.Group
            $x=$lights[0]

            $instrument=$x.Instrument
            $gain=[decimal]$x.Gain
            $offset=[decimal]$x.Offset
            $exposure=[decimal]$x.Exposure
            $ccdTemp = [decimal]$x.CCDTemp
            $setTemp=[decimal]$x.SetTemp

            $masterBias = $BiasLibraryFiles |
                where-object Gain -eq $gain |
                where-object Offset -eq $offset |
                where-object Instrument -eq $instrument |
                sort-object ObsDate -Descending |
                select-object -First 1
            if($masterBias){
                write-host "Master bias available for $instrument at Gain=$gain Offset=$offset. $($masterBias.Path.Name)"
            }
            else{
                Write-Warning "No master bias available for $instrument at Gain=$gain Offset=$offset."
            }

            $masterDark = $DarkLibrary | where-object {
                $dark = $_
                ($dark.Instrument-eq $instrument) -and
                ($dark.Gain-eq $gain) -and
                ($dark.Offset-eq $offset) -and
                ($dark.Exposure-eq $exposure) -and
                ([Math]::abs($dark.SetTemp - $ccdTemp) -lt 3)
            } | select-object -first 1

            if(-not $masterDark){
                if($masterBias){
                    $masterDark = $DarkLibrary | where-object {
                        $dark = $_
                        ($dark.Instrument-eq $instrument) -and
                        ($dark.Gain-eq $gain) -and
                        ($dark.Offset-eq $offset) -and
                        ($dark.Exposure-eq $exposure)
                    } | select-object -first 1
                    if($masterDark){
                        Write-Warning "No master dark available for $instrument at Gain=$gain Offset=$offset Exposure=$exposure (s) and SetTemp $setTemp. Attempting to scale temperature."
                    }
                    else{
                        Write-Warning "No master dark available for $instrument at Gain=$gain Offset=$offset Exposure=$exposure (s) and SetTemp $setTemp. Using bias only."
                    }
                }
                else{
                    Write-Warning "No master dark available for $instrument at Gain=$gain Offset=$offset Exposure=$exposure (s) and SetTemp $setTemp. Using bias only."
                }
            }
            $lights |
                    group-object Filter,FocalLength |
                    foreach-object {
                        $filter = $_.Group[0].Filter
                        $focalLength=$_.Group[0].FocalLength
                    
                        $masterBiasFile=$masterBias.Path
                        $masterDarkFile=$masterDark.Path
                        $optimizeDark = ($masterBiasFile -and $masterDarkFile)
                        $calibrateDark= ($masterBiasFile -and $masterDarkFile)
                        Write-Host "Sorting $($_.Group.Count) frames at ($focalLength)mm with filter $filter"
                        if($masterBias){
                            Write-Host " Bias: $($masterBias.Path)"
                        }
                        if($masterDark) {
                            Write-Host " Dark: $($masterDark.Path)"
                        }
                        Invoke-LightFrameSorting `
                            -XisfStats ($_.Group) -ArchiveDirectory $ArchiveDirectory `
                            -MasterBias $masterBiasFile -OptimizeDark:$OptimizeDark -CalibrateDark:$calibrateDark `
                            -MasterDark $masterDarkFile `
                            -OutputPath $CalibratedOutput `
                            -PixInsightSlot 201 `
                            -OutputPedestal 80 `
                            -Verbose -EnableCFA `
                            -AfterImagesCalibrated {
                                param($LightFrames)
                            try {
                                $targetOutputPath=Join-Path $CalibratedOutput $LightFrames[0].Object
                                $calibrated=($LightFrames|Get-XisfCalibrationState -CalibratedPath $targetOutputPath).Calibrated
                                #$corrected = $LightFrames|Get-XisfCosmeticCorrectionState -CosmeticCorrectionPath
                                Invoke-PiDebayer -PixInsightSlot 201 -Images $calibrated -OutputPath "S:\PixInsight\Timelapse\Debayered" -Compress -KeepOpen
                                
                            }
                            catch {
                                Write-Warning $_.Exception.ToString()
                            }
                        }
            }
        }
    write-host "waiting for next set of data..."
    Wait-Event -Timeout 30
}