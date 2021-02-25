import-module $PSScriptRoot/PsXisfReader.psd1 -Force
$ErrorActionPreference="STOP"
$VerbosePreference="Continue"
$target="E:\Astrophotography\135mm\Rosette Sh2-280 Sh2-284"

Clear-Host
$rawSubs = 
    Get-XisfLightFrames -Path $target -Recurse -UseCache -SkipOnError |
    where-object Instrument -eq "QHYCCD-Cameras-Capture (ASCOM)" |
    Where-Object {-not $_.HasTokensInPath(@("reject","process","testing","clouds","draft","cloudy","_ez_LS_"))} |
    Where-Object Filter -ne "V4" |
    Where-Object {-not $_.IsIntegratedFile()}

$data=Invoke-XisfPostCalibrationMonochromeImageWorkflow `
    -RawSubs $rawSubs `
    -CalibrationPath "E:\PixInsightLT\Calibrated" `
    -CorrectedOutputPath "S:\PixInsight\Corrected" `
    -WeightedOutputPath "S:\PixInsight\Weighted" `
    -DarkLibraryPath "E:\Astrophotography\DarkLibrary\QHYCCD-Cameras-Capture (ASCOM)" `
    -AlignedOutputPath "S:\PixInsight\Aligned" `
    -BackupCalibrationPaths @("T:\PixInsightLT\Calibrated","N:\PixInsightLT\Calibrated","S:\PixInsightLT\Calibrated") `
    -PixInsightSlot 200 `
    -RerunCosmeticCorrection:$false `
    -SkipCosmeticCorrection:$false `
    -RerunWeighting:$false `
    -SkipWeighting:$false `
    -RerunAlignment:$false `
    -IntegratedImageOutputDirectory $target `
    -AlignmentReference $alignmentReference `
    -GenerateDrizzleData `
    -ApprovalExpression "Median<50 && FWHM<4.5" `
    -WeightingExpression "(15*(1-(FWHM-FWHMMin)/(FWHMMax-FWHMMin))
    +  5*(1-(Eccentricity-EccentricityMin)/(EccentricityMax-EccentricityMin))
    + 15*(SNRWeight-SNRWeightMin)/(SNRWeightMax-SNRWeightMin)
    + 30*(1-(Median-MedianMin)/(MedianMax-MedianMin))
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
        if($_.Path -and (test-path $_.Path)){
            Move-Item ($_.Path) -Destination "$target\Rejection\" -Verbose
        }
    }
}

if((Read-Host -Prompt "Cleanup intermediate files (corrected, weighted, aligned, drizzle)?") -eq "Y"){
    $data|foreach-object{
        $_.RemoveAlignedAndDrizzleFiles()
        $_.RemoveWeightedFiles()
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
