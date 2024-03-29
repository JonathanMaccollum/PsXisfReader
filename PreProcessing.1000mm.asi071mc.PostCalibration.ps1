Clear-Host
if (-not (get-module psxisfreader)){import-module psxisfreader}
$ErrorActionPreference="STOP"

$target="E:\Astrophotography\1000mm\Sh 2-108"
$alignmentReference=$null
#$alignmentReference = "DWB111.Ha.16x600s.Int2.xisf"
#$alignmentReference="Sh 2-108.L3.70x120s.xisf"
#$alignmentReference="NGC188.L3.175x120s.xisf"
#$alignmentReference="LDN 1251 135mm OSC.L3.89x60s.L3.15x120s.xisf"
#$alignmentReference="Double Cluster.L3.39x90s.L3.7x120s.L3.17x180s.xisf"
#$alignmentReference = "_NGC 2112 OSC.L3.51x180s.xisf"
#$alignmentReference=Join-Path $target "Rosette Nebula OSC.L3.32x240s.Drizzled.xisf"
#$alignmentReference=Join-Path $target "LDN1463 OSC.L3.76x180s.xisf"
#$alignmentReference=Join-Path $target "NGC1333 Panel 3 OSC.L3.124x180s.xisf"
$alignmentReference = Join-Path $target "Sh 2-108.Oiii6nm.54x240s.Oiii6nm.31x360s.ESD.xisf"
Clear-Host

$rawSubs =
    Get-XisfLightFrames -Path $target -Recurse -SkipOnError -UseCache -PathTokensToIgnore $Ignorables |
    Where-Object {-not $_.HasTokensInPath(@("reject","process","testing","clouds","draft","cloudy","_ez_LS_"))} |
    Where-Object {-not $_.IsIntegratedFile()} |
    Where-Object Filter -eq "L3" |
    where-object Instrument -eq "ZWO ASI071MC Pro"
    

$data = Invoke-XisfPostCalibrationColorImageWorkflow `
    -RawSubs $rawSubs `
    -CalibrationPath "F:\PixInsightLT\Calibrated" `
    -CorrectedOutputPath "S:\PixInsight\Corrected" `
    -WeightedOutputPath "S:\PixInsight\Weighted" `
    -DebayeredOutputPath "S:\PixInsight\Debayered" `
    -DarkLibraryPath "E:\Astrophotography\DarkLibrary\ZWO ASI071MC Pro" `
    -AlignedOutputPath "S:\PixInsight\Aligned" `
    -BackupCalibrationPaths @("M:\PixInsightLT\Calibrated","S:\PixInsightLT\Calibrated") `
    -PixInsightSlot 200 `
    -RerunWeighting:$false `
    -RerunAlignment `
    -IntegratedImageOutputDirectory $target `
    -AlignmentReference $alignmentReference `
    -GenerateDrizzleData `
    -CfaPattern "RGGB" `
    -ApprovalExpression "Median<50 && FWHM<2.25 && Eccentricity<0.67" `
    -WeightingExpression "(15*(1-(FWHM-FWHMMin)/(FWHMMax-FWHMMin))
    +  5*(1-(Eccentricity-EccentricityMin)/(EccentricityMax-EccentricityMin))
    + 25*(SNRWeight-SNRWeightMin)/(SNRWeightMax-SNRWeightMin)
    + 10*((Median-MedianMin)/(MedianMax-MedianMin))
    + 20*(Stars-StarsMin)/(StarsMax-StarsMin))
    + 20" `
    -Verbose
                    
    
$stacked = $data | where-object {$_.Aligned -and (Test-Path $_.Aligned)}
$toReject = $data | where-object {-not $_.Aligned -or (-not (Test-Path $_.Aligned))}
Write-Host "Stacked: $($stacked.Stats | Measure-ExposureTime -TotalMinutes)"
Write-Host "Rejected: $($toReject.Stats | Measure-ExposureTime -TotalMinutes)"
if($toReject -and (Read-Host -Prompt "Move $($toReject.Count) Rejected files?") -eq "Y"){
    [System.IO.Directory]::CreateDirectory("$target\Rejection")>>$null
    $toReject | foreach-object {
        if(test-path $_.Path){
            Move-Item ($_.Path) -Destination "$target\Rejection\" -Verbose
        }
    }
}

if((Read-Host -Prompt "Cleanup intermediate files (corrected, debayered, weighted, aligned, drizzle)?") -eq "Y"){
    $data|foreach-object{
        $_.RemoveAlignedAndDrizzleFiles()
        $_.RemoveWeightedFiles()
        $_.RemoveDebayeredFiles()
        $_.RemoveCorrectedFiles()
    }
}

$mostRecent=
    $data.Stats|
    Sort-Object LocalDate -desc |
    Select-Object -first 1
if($data){
    $data | Export-Clixml -Path (Join-Path $target "Stats.$($mostRecent.LocalDate.ToString('yyyyMMdd HHmmss')).clixml")
}