if (-not (get-module psxisfreader)){import-module psxisfreader}
$ErrorActionPreference="STOP"

$target="E:\Astrophotography\40mm\Large Dark Nebula Complex in Aquila 40mm OSC"
$alignmentReference=$null


#$alignmentReference="E:\Astrophotography\50mm\Orion 50mm\Orion 50mm.Ha3nm.30x6min.ESD.Drizzle2x.xisf"
#$alignmentReference=Join-Path $target "Omega 40mm OSC.D1.63x180s.ESD.Drizzled.xisf"
Clear-Host

$rawSubs =
    Get-XisfLightFrames -Path $target -Recurse |
    Where-Object {-not $_.HasTokensInPath(@("reject","process","testing","clouds","draft","cloudy","_ez_LS_"))} |
    Where-Object {-not $_.IsIntegratedFile()} 

#    -CalibrationPath "F:\PixInsightLT\Calibrated\M81 M82 50mm\SuperFlatted" `

$data = Invoke-XisfPostCalibrationColorImageWorkflow `
    -RawSubs $rawSubs `
    -CorrectedOutputPath "S:\PixInsight\Corrected" `
    -CalibrationPath "F:\PixInsightLT\Calibrated" `
    -WeightedOutputPath "S:\PixInsight\Weighted" `
    -DebayeredOutputPath "S:\PixInsight\Debayered" `
    -DarkLibraryPath "E:\Astrophotography\DarkLibrary\ZWO ASI071MC Pro" `
    -AlignedOutputPath "S:\PixInsight\Aligned" `
    -BackupCalibrationPaths @("T:\PixInsightLT\Calibrated","N:\PixInsightLT\Calibrated","S:\PixInsightLT\Calibrated") `
    -PixInsightSlot 201 `
    -RerunCosmeticCorrection:$false `
    -SkipCosmeticCorrection:$false `
    -RerunWeighting:$false `
    -SkipWeighting:$false `
    -RerunAlignment:$false `
    -IntegratedImageOutputDirectory $target `
    -AlignmentReference $alignmentReference `
    -GenerateDrizzleData `
    -CfaPattern "RGGB" `
    -ApprovalExpression "Median<100 && FWHM<4.5" `
    -WeightingExpression "(15*(1-(FWHM-FWHMMin)/(FWHMMax-FWHMMin))
    +  5*(1-(Eccentricity-EccentricityMin)/(EccentricityMax-EccentricityMin))
    + 05*(SNRWeight-SNRWeightMin)/(SNRWeightMax-SNRWeightMin)
    + 20*(1-(Median-MedianMin)/(MedianMax-MedianMin))
    + 40*(Stars-StarsMin)/(StarsMax-StarsMin))
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

if((Read-Host -Prompt "Cleanup intermediate files (debayered, weighted, aligned, drizzle)?") -eq "Y"){
    $data|foreach-object{
        $_.RemoveAlignedAndDrizzleFiles()
        $_.RemoveWeightedFiles()
        $_.RemoveDebayeredFiles()
        #$_.RemoveCorrectedFiles()
    }
}

$mostRecent=
    $data.Stats|
    Sort-Object LocalDate -desc |
    Select-Object -first 1

$data | Export-Clixml -Path (Join-Path $target "Stats.$($mostRecent.LocalDate.ToString('yyyyMMdd HHmmss')).clixml")