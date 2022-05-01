if (-not (get-module psxisfreader)){import-module psxisfreader}
$ErrorActionPreference="STOP"

$target=""
$alignmentReference=$null
$targets = @(
    #"E:\Astrophotography\90mm\Hyades on HD28124"
    #"E:\Astrophotography\90mm\Orion on HD37903"
    #"E:\Astrophotography\90mm\Abell 31"
    #"E:\Astrophotography\90mm\sh2-202 on HD19536"
    #"E:\Astrophotography\90mm\LBN691 on HD83126"
    #"E:\Astrophotography\90mm\LBN691 Panel 2"
    #"E:\Astrophotography\90mm\Perseus on HD 232763"
    #"E:\Astrophotography\90mm\Cali to NGC1333"
    #"E:\Astrophotography\90mm\Pleiades"
    #"E:\Astrophotography\90mm\Orion Northern Panel"
    #"E:\Astrophotography\90mm\Orion Middle Panel"
    #"E:\Astrophotography\90mm\Orion Southern Panel"
    #"E:\Astrophotography\90mm\NGC3344 Widefield"
    #"E:\Astrophotography\90mm\M64 - Black Eye Galaxy Widefield"
    #"E:\Astrophotography\90mm\sh2-73"
    #"E:\Astrophotography\90mm\Lyra"
    #"E:\Astrophotography\90mm\Cygnus on 31Cyg"
    #"E:\Astrophotography\90mm\M101 Region on HD117449"
    #"E:\Astrophotography\90mm\Beehive and Massalia in Cancer"
    #"E:\Astrophotography\90mm\Whirlpool and Sunflower"
    #"E:\Astrophotography\90mm\Cygnus on Crescent"
    "E:\Astrophotography\90mm\Between LBN691 and Polaris"
    #"E:\Astrophotography\90mm\Scorpius on HD145468"
    #"E:\Astrophotography\90mm\Scorpius on HD145468 Plus 1hr RA"
)
$referenceImages = @(
    "Hyades on HD28124.D1.13x180s.LF.xisf"
    "LBN691 on HD83126.L3.77x180s.LF.xisf"
    "LBN691 Panel 2.L3.65x180s.LF.xisf"
    "Orion on HD37903.D1.26x180s.LF.xisf"
    "Pleiades.L3.83x180s.LF.xisf"
    "Cali to NGC1333.L3.87x180s.LF.xisf"
    "Perseus on HD 232763.L3.60x180s.LF.xisf"
    "Orion Northern Panel.L3.49x180s.LF.xisf"
    "Orion Middle Panel.L3.53x180s.LF.xisf"
    "Orion Southern Panel.L3.59x180s.LF.xisf"
    "Lyra.L3.96x180s.LF.xisf"
    "Whirlpool and Sunflower.L3.119x180s.LF.xisf"
    "Cygnus on 31Cyg.L3.100x180s.LF.xisf"
    "Cygnus on Crescent.L3.75x180s.LF.xisf"
    "Between LBN691 and Polaris.L3.166x180s.LF.xisf"
    "M64 - Black Eye Galaxy Widefield.L3.189x180s.LF.xisf"
)

$targets | foreach-object {
    $target = $_

    $alignmentReference = $null 
    $alignmentReference =
        $referenceImages | 
        foreach-object {
            Join-Path $target $_
        } |
        where-object {test-path $_} |
        Select-Object -First 1
    if(-not $alignmentReference){
        Write-Warning "$($target): No alignment reference was specified... a new reference will automatically be selected."
        Wait-Event -Timeout 1
    }

    $rawSubs =
        Get-XisfLightFrames -Path $target -Recurse |
        Where-Object {-not $_.HasTokensInPath(@("reject","process","testing","clouds","draft","cloudy","_ez_LS_"))} |
        Where-Object {-not $_.IsIntegratedFile()} 

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
        -Verbose `
        -GenerateThumbnail

    if($data){
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
    }

}