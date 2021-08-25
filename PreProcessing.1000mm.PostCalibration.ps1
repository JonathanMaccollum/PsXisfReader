Clear-Host
if (-not (get-module psxisfreader)){import-module psxisfreader}
$ErrorActionPreference="STOP"
$VerbosePreference="Continue"
$target="E:\Astrophotography\1000mm\NGC7000 Framing 2"
$alignmentReference=$null 

#$alignmentReference = join-path $target "Eye of Smaug in Cygnus Panel -1.Ha.24x360s.ESD.xisf"
#$alignmentReference = join-path $target "Eye of Smaug in Cygnus.Ha.30x360s.ESD.xisf"
#$alignmentReference = join-path $target "Eye of Smaug in Cygnus Panel 0.Ha.33x360s.ESD.xisf"
#$alignmentReference = join-path $target "Eye of Smaug Panel 2.Ha.50x360s.ESD.xisf"
#$alignmentReference = join-path $target "Eye of Smaug Panel 3.Ha.18x360s.ESD.xisf"
#$alignmentReference = join-path $target "Eye of Smaug Panel 4.Ha.21x360s.ESD.xisf"
#$alignmentReference = join-path $target "Eye of Smaug in Cygnus Panel 5.Ha.27x360s.ESD.xisf"
#$alignmentReference = join-path $target "Western Smaug in Cygnus Panel 5.Ha.33x180s.ESD.xisf"
#$alignmentReference = join-path $target "IC417.Ha.113x360s.ESD.xisf"
#$alignmentReference = join-path $target "M38.Ha.18x360s.ESD.xisf"
#$alignmentReference = join-path $target "Jellyfish Exhaust.Ha.15x360s.ESD.xisf"
#$alignmentReference = join-path $target "Ring Nebula.Ha.138x360s.ESD.xisf"
#$alignmentReference = join-path $target "NGC5905-5908 Panel 1.D1.44x120s.ESD.xisf"
#$alignmentReference = join-path $target "Abell 31.Ha.130x360s.Ha.6x600s.Adaptive.ESD.xisf"
#$alignmentReference = join-path $target "NGC4725.L3.61x240s.ESD.xisf"
#$alignmentReference = join-path $target "Owl Nebula.Ha.71x360s.ESD.xisf"
#$alignmentReference = join-path $target "Abell 39.Oiii6nm.21x240s.ESD.xisf"
#$alignmentReference=Join-Path $target "LBN691 Panel 1.L3.3x120s.L3.38x240s.ESD.xisf"
#$alignmentReference=Join-Path $target "Sh2-73.B.7x240s.G.8x240s.L3.6x240s.R.7x240s.SuperLum.Adaptive.xisf"
#$alignmentReference = Join-Path $target "Sh 2-108.Oiii6nm.54x240s.Oiii6nm.31x360s.ESD.xisf"
#$alignmentReference=Join-Path $target "vdb131 vdb132 Panel2.Ha.42x360s.ESD.xisf"
#$alignmentReference=Join-Path $target "vdb131 vdb132.Ha.48x360s.ESD.xisf"
#$alignmentReference=Join-Path $target "wr134.Sii6nm.47x180s.ESD.xisf"
#$alignmentReference=Join-Path $target "sh2-86.Ha.47x180s.ESD.xisf"

#$alignmentReference=Join-Path $target "Tulip Panel 4.Ha.49x180s.ESD.xisf"
#$alignmentReference=Join-Path $target "Tulip Panel 3.Ha.47x180s.ESD.xisf"
#$alignmentReference=Join-Path $target "Tulip Panel 2.Ha.44x180s.ESD.xisf"
#$alignmentReference=Join-Path $target "Tulip Panel 1.Ha.49x180s.ESD.xisf"
#$alignmentReference = join-path $target "Cave Nebula OSC.Ha.21x180s.ESD.xisf"
#$alignmentReference = join-path $target "IC1396 Elephants Trunk Nebula.Ha.34x180s.Ha.16x360s.ESD.xisf"
#$alignmentReference = join-path $target "Lobster Claw in Cepheus.Ha.38x180s.Ha.16x360s.ESD.xisf"
$alignmentReference = join-path $target "NGC7000 Framing 2.Ha.142x180s.ESD.xisf"
$rawSubs = 
    Get-XisfLightFrames -Path $target -Recurse -UseCache -SkipOnError |
    #where-object Instrument -eq "QHYCCD-Cameras-Capture (ASCOM)" |
    #where-object Instrument -eq "QHY268M" |
    Where-Object {-not $_.HasTokensInPath(@("reject","process","testing","clouds","draft","cloudy","_ez_LS_","drizzle"))} |
    #Where-Object Filter -in @("Ha","R","G","B") |
    Where-Object Filter -ne "L3" |
    #Where-Object Filter -ne "Oiii6nm" |
    #Where-Object Filter -eq "Sii6nm" |
    #Where-Object Exposure -eq 360 |
    #Where-object ObsDateMinus12hr -eq ([DateTime]"2021-05-05")
    Where-Object {-not $_.IsIntegratedFile()} #|
    #select-object -First 30
#$rawSubs|Format-Table Path,*
$createSuperLum=$false
$data=Invoke-XisfPostCalibrationMonochromeImageWorkflow `
    -RawSubs $rawSubs `
    -CalibrationPath "F:\PixInsightLT\Calibrated" `
    -CorrectedOutputPath "S:\PixInsight\Corrected" `
    -WeightedOutputPath "S:\PixInsight\Weighted" `
    -DarkLibraryPath "E:\Astrophotography\DarkLibrary\QHY268M" `
    <#-DarkLibraryPath "E:\Astrophotography\DarkLibrary\QHYCCD-Cameras-Capture (ASCOM)"#> `
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
    -ApprovalExpression "Median<50 && FWHM<2.9" `
    -WeightingExpression "(15*(1-(FWHM-FWHMMin)/(FWHMMax-FWHMMin))
    +  5*(1-(Eccentricity-EccentricityMin)/(EccentricityMax-EccentricityMin))
    + 15*(SNRWeight-SNRWeightMin)/(SNRWeightMax-SNRWeightMin)
    + 30*(1-(Median-MedianMin)/(MedianMax-MedianMin))
    + 20*(Stars-StarsMin)/(StarsMax-StarsMin))
    + 20" `
    -Rejection "Rejection_ESD" `
    -GenerateThumbnail `
    -Verbose

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
        #$_.RemoveCorrectedFiles()
    }
}

$mostRecent=
    $data.Stats|
    Sort-Object LocalDate -desc |
    Select-Object -first 1
if($data){
    $data | Export-Clixml -Path (Join-Path $target "Stats.$($mostRecent.LocalDate.ToString('yyyyMMdd HHmmss')).clixml")
}
