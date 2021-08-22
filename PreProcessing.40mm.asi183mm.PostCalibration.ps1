if (-not (get-module psxisfreader)){import-module psxisfreader}
$ErrorActionPreference="STOP"
$VerbosePreference="Continue"
$target="E:\Astrophotography\40mm\Greater Cepheus"
$alignmentReference=$null
#$alignmentReference="PatchworkCygnus_0_3.BHS_Ha.11x240s.BHS_Ha.6x300s.BHS_Ha.5x600s.xisf"
#$alignmentReference="PatchworkCygnus_0_4.BHS_Ha.36x300s.Adaptive.xisf"
#$alignmentReference="PatchworkCygnus_1_1.BHS_Ha.15x600s.Ha.17x600s.Ha.21x240s.xisf"
#$alignmentReference="PatchworkCygnus_1_2.BHS_Ha.40x240s.Ha.6x240s.BHS_Ha.11x600s.Ha.8x600s.xisf"
#$alignmentReference="PatchworkCygnus_1_3.BHS_Ha.14x600s.xisf"
#$alignmentReference="PatchworkCygnus_1_4.BHS_Ha.8x300s.Adaptive.BHS_Ha.26x360s.Adaptive.BHS_Ha.19x600s.Adaptive.xisf"
#$alignmentReference="PatchworkCygnus_2_0.BHS_Ha.23x600s.xisf"
#$alignmentReference="PatchworkCygnus_2_1.BHS_Ha.22x360s.Adaptive.BHS_Ha.15x600s.Adaptive.xisf"
#$alignmentReference="PatchworkCygnus_2_2.BHS_Ha.15x600s.xisf"
#$alignmentReference="PatchworkCygnus_2_4.BHS_Ha.10x300s.BHS_Ha.2x360s.Adaptive.xisf"
#$alignmentReference="Soul Nebula.BHS_Ha.13x600s.Adaptive.ESD.xisf"
#$alignmentReference="Cassiopeia.BHS_Ha.10x600s.Adaptive.ESD.xisf"
#$alignmentReference=Join-Path $target "Taurus to Pleiades 2 50mm.BHS_Ha.33x360s.ESD.xisf"
#$alignmentReference=Join-Path $target "PatchworkCygnus_2_2.BHS_Ha.20x360s.BHS_Ha.19x600s.Adaptive.ESD.xisf"
#$alignmentReference=Join-Path $target "Patchwork Cygnus Center.Ha.21x360s.LF.xisf"
Clear-Host

$rawSubs =
    Get-XisfLightFrames -Path $target -Recurse -UseCache -SkipOnError |
    Where-Object {-not $_.HasTokensInPath(@("reject","process","testing","clouds","draft","cloudy","_ez_LS_"))} |
    #Where-Object Filter -ne "Oiii" |
    #Where-Object Filter -eq "BHS_Sii" |
    Where-Object {-not $_.IsIntegratedFile()} 

$data=Invoke-XisfPostCalibrationMonochromeImageWorkflow `
    -RawSubs $rawSubs `
    -CalibrationPath "F:\PixInsightLT\Calibrated" `
    -CorrectedOutputPath "S:\PixInsight\Corrected" `
    -WeightedOutputPath "S:\PixInsight\Weighted" `
    -DarkLibraryPath "E:\Astrophotography\DarkLibrary\ZWO ASI183MM Pro" `
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
    -ApprovalExpression "Median<120 && FWHM<4.5" `
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
