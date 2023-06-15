Clear-Host
#update-module PsXisfReader
if (-not (get-module psxisfreader)){import-module psxisfreader}
$ErrorActionPreference="STOP"
$VerbosePreference="Continue"
#$target="E:\Astrophotography\1000mm\Eye of Smaug in Cygnus Panel 2"
$target="E:\Astrophotography\730mm\IC410 - Tadpoles"

#Get-XisfFitsStats -Path "E:\Astrophotography\1000mm\LDN 1657 in Seagull Panel 1\360.00s\LDN 1657 in Seagull Panel 1_Ha3nm_LIGHT_2021-03-21_21-22-16_0000_360.00s_-15.00_0.49.xisf" |
#Get-XisfCalibrationState -CalibratedPath "F:\PixInsightLT\Calibrated" -Verbose -AdditionalSearchPaths @("M:\PixInsightLT\Calibrated","S:\PixInsightLT\Calibrated") -Recurse

$referenceImages = @(
    "IC410 - Tadpoles.Ha.77x180s.ESD.xisf"
)
$alignmentReference = $null 
$alignmentReference = 
    $referenceImages | 
    foreach-object {
        Join-Path $target $_
    } |
    where-object {test-path $_} |
    Select-Object -First 1
if(-not $alignmentReference){
    Write-Warning "No alignment reference was specified... a new reference will automatically be selected."
    Wait-Event -Timeout 5
}
$rawSubs = 
    Get-XisfLightFrames -Path $target -Recurse -UseCache -SkipOnError -PathTokensToIgnore @("reject","process","testing","clouds","draft","cloudy","_ez_LS_","drizzle","quick") |
    #where-object Instrument -eq "QHYCCD-Cameras-Capture (ASCOM)" |
    #where-object Instrument -eq "QHY268M" |
    #Where-Object Filter -NotIn @("Ha","L3") |
    #Where-Object {-not $_.Filter.Contains("Oiii")} |
    #Where-Object Filter -eq "Oiii" |
    #Where-Object Filter -eq "Ha" |
    #Where-Object Exposure -eq 180 |
    #Where-object ObsDateMinus12hr -gt ([DateTime]"2022-02-18") |
    Where-Object {-not $_.IsIntegratedFile()} #|
    #select-object -First 30
#$rawSubs|Format-Table Path,*
$createSuperLum=$false
$data=Invoke-XisfPostCalibrationMonochromeImageWorkflow `
    -RawSubs $rawSubs `
    -CalibrationPath "F:\PixInsightLT\Calibrated" `
    -CorrectedOutputPath "S:\PixInsight\Corrected" `
    -WeightedOutputPath "S:\PixInsight\Weighted" `
    <#-DarkLibraryPath "E:\Astrophotography\DarkLibrary\QHY268M"#> `
    -DarkLibraryPath "E:\Astrophotography\DarkLibrary\QHYCCD-Cameras-Capture (ASCOM)" `
    -AlignedOutputPath "S:\PixInsight\Aligned" `
    -BackupCalibrationPaths @("M:\PixInsightLT\Calibrated","S:\PixInsightLT\Calibrated") `
    -PixInsightSlot 200 `
    -RerunCosmeticCorrection:$false `
    -SkipCosmeticCorrection:$false `
    -RerunWeighting:$false `
    -SkipWeighting:$false `
    -RerunAlignment:$false `
    -IntegratedImageOutputDirectory $target `
    -AlignmentReference $alignmentReference `
    -GenerateDrizzleData `
    -ApprovalExpression "Median<400 && FWHM<2.7" `
    -WeightingExpression "(15*(1-(FWHM-FWHMMin)/(FWHMMax-FWHMMin))
    +  5*(1-(Eccentricity-EccentricityMin)/(EccentricityMax-EccentricityMin))
    + 15*(SNRWeight-SNRWeightMin)/(SNRWeightMax-SNRWeightMin)
    + 30*(1-(Median-MedianMin)/(MedianMax-MedianMin))
    + 20*(Stars-StarsMin)/(StarsMax-StarsMin))
    + 20" `
    -Rejection "Rejection_ESD" `
    -GenerateThumbnail `
    -Verbose
if($data){

    $stacked = $data | where-object {$_.Aligned -and (Test-Path $_.Aligned)}
    $toReject = $data | where-object {-not $_.Aligned -or (-not (Test-Path $_.Aligned))}
    Write-Host "Stacked: $($stacked.Stats | Measure-ExposureTime -TotalMinutes)"
    Write-Host "Rejected: $($toReject.Stats | Measure-ExposureTime -TotalMinutes)"
    $stacked.Aligned |
        Get-XisfFitsStats | 
        group-object Filter | foreach-object{
            $group = $_.Group
            $filter = $group[0].Filter
            " $filter - $($group | Measure-ExposureTime -TotalMinutes)"
        }

    if($createSuperLum){
    <#Super Luminance#>
        $approved = $stacked.Aligned |
            Get-XisfFitsStats | 
            Where-Object Filter -ne "IR742"
        $reference =  $approved |
            Sort-Object SSWeight -Descending |
            Select-Object -First 1
        $outputFileName = $reference.Object
        $approved | group-object Filter | foreach-object{
            $filter = $_.Group[0].Filter
            $_.Group | group-object Exposure | foreach-object {
                $exposure=$_.Group[0].Exposure;
                $outputFileName+=".$filter.$($_.Group.Count)x$($exposure)s"
            }
        }
        $outputFileName
        $outputFileName+=".SuperLum.xisf"
        
        $outputFile = Join-Path $target $outputFileName
        if(-not (test-path $outputFile)) {
            write-host ("Integrating  "+ $outputFileName)
            $toStack = $approved | sort-object SSWeight -Descending
            $toStack | 
            Group-Object Filter | 
            foreach-object {$dur=$_.Group|Measure-ExposureTime -TotalSeconds; new-object psobject -Property @{Filter=$_.Name; ExposureTime=$dur}} |
            foreach-object {
                write-host "$($_.Filter): $($_.Exposure)"
            }
            try {
            Invoke-PiLightIntegration `
                -Images ($toStack|foreach-object {$_.Path}) `
                -OutputFile $outputFile `
                -KeepOpen `
                -PixInsightSlot 200 `
                -WeightKeyword:"SSWEIGHT"
            }
            catch {
                write-warning $_.ToString()
                throw
            }
        }
    }

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
}
