Clear-Host
#update-module PsXisfReader
if (-not (get-module psxisfreader)){import-module psxisfreader}
$ErrorActionPreference="STOP"
$VerbosePreference="Continue"
#$target="E:\Astrophotography\1000mm\Eye of Smaug in Cygnus Panel 2"
#$target="E:\Astrophotography\1000mm\Flame and Horsehead Nebula P1"
#$target="E:\Astrophotography\1000mm\NGC 3344"
#$target="E:\Astrophotography\1000mm\Rosette Nebula"
#$target="E:\Astrophotography\1000mm\M64 - Black Eye Galaxy"
#$target="E:\Astrophotography\1000mm\LDN1622"
#$target="E:\Astrophotography\1000mm\NGC 4535 and NGC 4560"
#$target="E:\Astrophotography\1000mm\Muppet in Auriga"
#$target="E:\Astrophotography\1000mm\Tulip Panel 4"
#$target="E:\Astrophotography\1000mm\Abell 39"
#$target="E:\Astrophotography\1000mm\M60 sn2022hrs in NGC4647"
#$target="E:\Astrophotography\1000mm\Abell 35"
#$target="E:\Astrophotography\1000mm\Abell 31"
#$target="E:\Astrophotography\1000mm\Sadr Take 3"
#$target="E:\Astrophotography\1000mm\Eye of Smaug Take 2"
$target="E:\Astrophotography\1000mm\Eye of Smaug Take 2P2"
#$target="E:\Astrophotography\1000mm\vdb131 vdb132 Take 2"

$createSuperLum=$false

#Get-XisfFitsStats -Path "E:\Astrophotography\1000mm\LDN 1657 in Seagull Panel 1\360.00s\LDN 1657 in Seagull Panel 1_Ha3nm_LIGHT_2021-03-21_21-22-16_0000_360.00s_-15.00_0.49.xisf" |
#Get-XisfCalibrationState -CalibratedPath "F:\PixInsightLT\Calibrated" -Verbose -AdditionalSearchPaths @("M:\PixInsightLT\Calibrated","S:\PixInsightLT\Calibrated") -Recurse

$referenceImages = @(
    "Jellyfish Exhaust.Ha.31x180s.Ha.9x240s.Ha.15x360s.ESD.xisf"
"LDN1622.Ha.26x240s.Ha.46x360s.ESD.xisf"
"M13.B.16x120s.G.15x120s.IR742.76x120s.R.16x120s.SuperLum.xisf"
"Eye of Smaug in Cygnus Panel -1.Ha.24x360s.ESD.xisf"
"Eye of Smaug in Cygnus Panel 0.Ha.33x360s.ESD.xisf"
"Eye of Smaug in Cygnus.Ha.30x360s.ESD.xisf"
"Eye of Smaug Panel 2.Ha.50x360s.ESD.xisf"
"Eye of Smaug Panel 3.Ha.18x360s.ESD.xisf"
"Eye of Smaug Panel 4.Ha.21x360s.ESD.xisf"
"Eye of Smaug in Cygnus Panel 5.Ha.27x360s.ESD.xisf"

"Western Smaug in Cygnus Panel -1.Ha.50x180s.ESD.xisf"
"Western Smaug in Cygnus Panel 0.Ha.59x180s.ESD.xisf"
"Western Smaug in Cygnus Panel 1.Ha.52x180s.ESD.xisf"
"Western Smaug in Cygnus Panel 2.Ha.50x180s.ESD.xisf"
"Western Smaug in Cygnus Panel 4.Ha.73x180s.ESD.xisf"
"Western Smaug in Cygnus Panel 3.Ha.49x180s.ESD.xisf"
"Western Smaug in Cygnus Panel 5.Ha.33x180s.ESD.xisf"

"Eastern Smaug in Cygnus Panel -1.Ha.18x180s.ESD.xisf"
"Eastern Smaug in Cygnus Panel 0.Ha.18x180s.ESD.xisf"
"Eastern Smaug in Cygnus Panel 1.Ha.40x180s.ESD.xisf"
"Eastern Smaug in Cygnus Panel 2.Ha.52x180s.Ha.20x360s.ESD.xisf"
"Eastern Smaug in Cygnus Panel 3.Ha.75x180s.ESD.xisf"
"Eastern Smaug in Cygnus Panel 4.Ha.75x180s.ESD.xisf"
"Eastern Smaug in Cygnus Panel 5.Ha.52x180s.ESD.xisf"

"IC417.Ha.113x360s.ESD.xisf"
"M38.Ha.18x360s.ESD.xisf"
"Jellyfish Exhaust.Ha.15x360s.ESD.xisf"
"Ring Nebula.Ha.138x360s.ESD.xisf"
"NGC5905-5908 Panel 1.D1.44x120s.ESD.xisf"
"Abell 31.Ha.130x360s.Ha.6x600s.Adaptive.ESD.xisf"
"NGC4725.L3.61x240s.ESD.xisf"
"Owl Nebula.Ha.71x360s.ESD.xisf"
"Abell 39.Oiii6nm.21x240s.ESD.xisf"
"LBN691 Panel 1.L3.3x120s.L3.38x240s.ESD.xisf"
"Sh2-73.B.7x240s.G.8x240s.L3.6x240s.R.7x240s.SuperLum.Adaptive.xisf"
"_Sh 2-108.Oiii6nm.54x240s.Oiii6nm.31x360s.ESD.xisf"
"vdb131 vdb132 Panel2.Ha.42x360s.ESD.xisf"
"vdb131 vdb132.Ha.48x360s.ESD.xisf"
"wr134.Sii6nm.47x180s.ESD.xisf"
"sh2-86.Ha.47x180s.ESD.xisf"

"Tulip Panel 4.Ha.49x180s.ESD.xisf"
"Tulip Panel 3.Ha.47x180s.ESD.xisf"
"Tulip Panel 2.Ha.44x180s.ESD.xisf"
"Tulip Panel 1.Ha.49x180s.ESD.xisf"
"Cave Nebula OSC.Ha.21x180s.ESD.xisf"
"IC1396 Elephants Trunk Nebula.Ha.34x180s.Ha.16x360s.ESD.xisf"
"Lobster Claw in Cepheus.Ha.38x180s.Ha.16x360s.ESD.xisf"
"NGC7000 Framing 2.Ha.142x180s.ESD.xisf"
"IC1871.Ha.31x360s.LF.xisf"
"Sh2-132 - Lion Nebula.Ha.27x180s.Ha.14x360s.ESD.xisf"
"sh2-115.Oiii.4x180s.Oiii.4x360s.ESD.xisf"
"NGC7129 NGC7142.Ha.20x360s.ESD.xisf"
"Heart Nebula.Ha.14x360s.ESD.xisf"
"NGC 896.Ha.11x360s.ESD.xisf"

"Flame and Horsehead Nebula P1.Ha.31x180s.ESD.xisf"
"Flame and Horsehead Nebula P2.Ha.26x180s.ESD.xisf"
"NGC 3344.L.29x90s.ESD.xisf"
"IC417.Ha.88x360s.LF.xisf"
"vdb31.B.34x180s.G.37x180s.L.70x90s.R.36x180s.SuperLum.xisf"
"The Coma Cluster.B.20x180s.G.19x180s.L.78x90s.R.21x180s.SuperLum.xisf"
"Rosette Nebula.B.27x90s.G.31x90s.Ha.7x180s.Ha.17x240s.Oiii.25x180s.R.25x90s.Sii3.26x180s.SuperLum.xisf"
"M64 - Black Eye Galaxy.L.109x90s.L.1x180s.ESD.xisf"
"_NGC 2112 - Take 2.B.14x90s.G.15x90s.Ha.20x180s.R.15x90s.Sii3.13x360s.SuperLum.xisf"
"_NGC 4535 and NGC 4560.L.76x90s.ESD.xisf"
"Muppet in Auriga.Ha.51x180s.Ha.15x360s.ESD.xisf"
"M60 sn2022hrs in NGC4647.LRGB.L.41x90s.R.35x90s.G.42x90s.B.44x90s.SynthLum.xisf"
"Sadr Take 3.Ha6nmMaxFR.187x90s.ESD.xisf"
"Eye of Smaug Take 2.R.10x45s.ESD.xisf"
"vdb131 vdb132 Take 2.L.38x45s.PSFSW.ESD.xisf"
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
    Get-XisfLightFrames -Path $target -Recurse -UseCache -SkipOnError |
    #where-object Instrument -eq "QHYCCD-Cameras-Capture (ASCOM)" |
    #where-object Instrument -eq "QHY268M" |
    Where-Object {-not $_.HasTokensInPath(@("reject","process","testing","clouds","draft","cloudy","_ez_LS_","drizzle","quick"))} |
    #Where-Object Filter -NotIn @("Ha","L3") |
    #Where-Object {-not $_.Filter.Contains("Oiii")} |
    Where-Object Filter -ne "V4" |
    Where-Object Instrument -ne "QHY294PROM" |
    #Where-Object Filter -in @( "R","G","B" )|
    #Where-Object Filter -ne "L" |
    #Where-Object Exposure -ne 30 |
    #Where-object ObsDateMinus12hr -ge ([DateTime]"2022-03-20") |
    #Where-object ObsDateMinus12hr -eq ([DateTime]"2022-05-04") |
    Where-Object {-not $_.IsIntegratedFile()} #|
    #select-object -First 30
#$rawSubs|Format-Table Path,*
#$rawSubs = $rawSubs | Where-Object ObsDateMinus12hr -eq "2022-04-27"
#$rawSubs|Group-Object ObsDateMinus12hr
$data=Invoke-XisfPostCalibrationMonochromeImageWorkflow `
    -RawSubs $rawSubs `
    -CalibrationPath "F:\PixInsightLT\Calibrated" `
    -CorrectedOutputPath "S:\PixInsight\Corrected" `
    -WeightedOutputPath "S:\PixInsight\Weighted" `
    -DarkLibraryPath "E:\Astrophotography\DarkLibrary\QHY600M" `
    <#-DarkLibraryPath "E:\Astrophotography\DarkLibrary\QHYCCD-Cameras-Capture (ASCOM)"#> `
    -AlignedOutputPath "S:\PixInsight\Aligned" `
    -BackupCalibrationPaths @("M:\PixInsightLT\Calibrated","S:\PixInsightLT\Calibrated") `
    -PixInsightSlot 200 -PSFSignalWeightWeighting `
    -RerunCosmeticCorrection:$false `
    -SkipCosmeticCorrection:$false `
    -RerunWeighting:$false `
    -SkipWeighting:$false `
    -RerunAlignment:$false `
    -IntegratedImageOutputDirectory $target `
    -AlignmentReference $alignmentReference `
    -GenerateDrizzleData `
    -ApprovalExpression "Median<150 && FWHM<4.3" `
    -WeightingExpression "(15*(1-(FWHM-FWHMMin)/(FWHMMax-FWHMMin))
    +  5*(1-(Eccentricity-EccentricityMin)/(EccentricityMax-EccentricityMin))
    + 15*(SNRWeight-SNRWeightMin)/(SNRWeightMax-SNRWeightMin)
    + 20*(1-(Median-MedianMin)/(MedianMax-MedianMin))
    + 30*(Stars-StarsMin)/(StarsMax-StarsMin))
    + 20" `
      `
    -GenerateThumbnail -Rejection "Rejection_ESD" `
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
}
