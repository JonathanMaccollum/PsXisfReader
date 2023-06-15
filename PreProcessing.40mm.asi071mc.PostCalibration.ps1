if (-not (get-module psxisfreader)){import-module psxisfreader}
$ErrorActionPreference="STOP"

#$target="E:\Astrophotography\40mm\Lobster Claw and Cave"
#$target="E:\Astrophotography\40mm\Anglerfish Dark Shark Rosebud Panel 1"
#$target="E:\Astrophotography\40mm\Cepheus near Alrai"
#$target="E:\Astrophotography\40mm\LDN 1472 in Perseus"
#$target="E:\Astrophotography\40mm\Crescent Oxygen River"
#$target="E:\Astrophotography\40mm\Cepheus on Barnard 170"
#$target="E:\Astrophotography\40mm\Cygnus near DWB111"
#$target="E:\Astrophotography\40mm\Heart and Soul"
#$target="E:\Astrophotography\40mm\Flaming Star Region"
$target="E:\Astrophotography\40mm\NGC1333 Region"
#$target="E:\Astrophotography\40mm\Cepheus near CTB-1"
#$target="E:\Astrophotography\40mm\Cygnus near Sh2-115"

$alignmentReference=$null


#$alignmentReference="E:\Astrophotography\50mm\Orion 50mm\Orion 50mm.Ha3nm.30x6min.ESD.Drizzle2x.xisf"
#$alignmentReference=Join-Path $target "Omega 40mm OSC.D1.63x180s.ESD.Drizzled.xisf"
#$alignmentReference=Join-Path $target "Cepheus on Barnard 170.L3.27x360s.ESD.LN.xisf"
#$alignmentReference=Join-Path $target "LDN 1472 in Perseus.L3.83x360s.PSFSW.ESD.LN.xisf"
#$alignmentReference=Join-Path $target "Heart and Soul.L3.61x360s.ESD.xisf"
#$alignmentReference=Join-Path $target "Crescent Oxygen River.L3.37x360s.ESD.LN.xisf"
#$alignmentReference=Join-Path $target "Cygnus near DWB111.L3.88x360s.ESD.xisf"
#$alignmentReference=Join-Path $target "Anglerfish Dark Shark Rosebud.L3.49x360s.ESD.LN.xisf"
$alignmentReference=Join-Path $target "NGC1333 Region.L3.205x360s.ESD.LN.xisf"


Clear-Host

$rawSubs =
    Get-XisfLightFrames -Path $target -Recurse |
    Where-Object {-not $_.HasTokensInPath(@("reject","process","testing","clouds","draft","cloudy","_ez_LS_"))} |
    Where-Object {-not $_.IsIntegratedFile()} |
    #where-object FocalRatio -eq 2.8 |
    where-object Instrument -eq "ZWO ASI071MC Pro"

#    -CalibrationPath "F:\PixInsightLT\Calibrated\M81 M82 50mm\SuperFlatted" `

$data = Invoke-XisfPostCalibrationColorImageWorkflow `
    -RawSubs $rawSubs `
    -CorrectedOutputPath "S:\PixInsight\Corrected" `
    -CalibrationPath "E:\Calibrated\40mm" `
    -WeightedOutputPath "S:\PixInsight\Weighted" `
    -DebayeredOutputPath "S:\PixInsight\Debayered" `
    -DarkLibraryPath "E:\Astrophotography\DarkLibrary\ZWO ASI071MC Pro" `
    -AlignedOutputPath "S:\PixInsight\Aligned" `
    -BackupCalibrationPaths @("M:\PixInsightLT\Calibrated","S:\PixInsightLT\Calibrated") `
    -PixInsightSlot 201 `
    -RerunCosmeticCorrection:$false `
    -SkipCosmeticCorrection:$false `
    -RerunWeighting:$false `
    -SkipWeighting:$false `
    -RerunAlignment:$false `
    -IntegratedImageOutputDirectory $target `
    -AlignmentReference $alignmentReference `
    -GenerateDrizzleData `
    -Rejection "Rejection_ESD" `
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
        $_.RemoveCorrectedFiles()
    }
}

$mostRecent=
    $data.Stats|
    Sort-Object LocalDate -desc |
    Select-Object -first 1

$data | Export-Clixml -Path (Join-Path $target "Stats.$($mostRecent.LocalDate.ToString('yyyyMMdd HHmmss')).clixml")