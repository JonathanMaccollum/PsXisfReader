Clear-Host
#update-module PsXisfReader
if (-not (get-module psxisfreader)){import-module psxisfreader}
$ErrorActionPreference="STOP"
$VerbosePreference="Continue"
$targets = @(
     #"W:\Astrophotography\950mm\M45 Take 2"
     #"W:\Astrophotography\950mm\Jellyfish Exhaust Take 2"
     #"W:\Astrophotography\950mm\LDN 1551 1546 LBN 821"
     #"W:\Astrophotography\950mm\C2022 E3 ZTF Night8"
     #"W:\Astrophotography\950mm\C2022 E3 ZTF Night11a"
     #"W:\Astrophotography\950mm\C2022 E3 ZTF Night11b"
     #"W:\Astrophotography\950mm\C2022 E3 ZTF Night11c"
     #"W:\Astrophotography\950mm\C2022 E3 ZTF Night12a"
     #"W:\Astrophotography\950mm\C2022 E3 ZTF Night12b"
     #"W:\Astrophotography\950mm\C2022 E3 ZTF Night12c"
     #"W:\Astrophotography\950mm\C2022 E3 ZTF Night12d"
     #"W:\Astrophotography\950mm\C2022 E3 ZTF Night12e"
     #"W:\Astrophotography\950mm\C2022 E3 ZTF Night12f"
     #"W:\Astrophotography\950mm\The Coma Cluster"
     #"W:\Astrophotography\950mm\Sh2-261 - Lowers Nebula"
     #"W:\Astrophotography\950mm\T Coronae Borealis"
     #"W:\Astrophotography\950mm\Arcturus"
     #"W:\Astrophotography\950mm\M35 and NGC2158"
     #"W:\Astrophotography\950mm\NGC 3344 Take 2"
     #"W:\Astrophotography\950mm\M51"
     #"W:\Astrophotography\950mm\NGC 3628 - Hamburger Galaxy"
     #"W:\Astrophotography\950mm\Dolphin Nebula in Cygnus"
     #"W:\Astrophotography\950mm\M106 Take 3"
     #"W:\Astrophotography\950mm\Sadr Take 3h"
     #"W:\Astrophotography\950mm\Bubble Nebula"
     #"W:\Astrophotography\950mm\vdB 15"
     #"W:\Astrophotography\950mm\LDN 1235 - Dark Shark Nebula Take 2"
     #"W:\Astrophotography\950mm\vdB 10"
     #"W:\Astrophotography\950mm\Flaming Star Nebula"
     #"W:\Astrophotography\950mm\IC410 - Tadpoles"
     #"W:\Astrophotography\950mm\M 33"
     #"W:\Astrophotography\950mm\vdb 13 and LDN 1448 and LDN 1451"
     #"W:\Astrophotography\950mm\M81 and M82"
     #"W:\Astrophotography\950mm\Barnard 3 and 4"
     #"W:\Astrophotography\950mm\Cone and Fox Fur"
     #"W:\Astrophotography\950mm\Ursa Major on HD 83489"
     #"W:\Astrophotography\950mm\M63 Take 2"
     #"W:\Astrophotography\950mm\PK 164 31.1"
     #"W:\Astrophotography\950mm\NGC 2633 and NGC 2634"
     #"W:\Astrophotography\950mm\LBN415 feature in Draco"
     #"W:\Astrophotography\950mm\Barnard 343 Region in Cygnus"
     #"W:\Astrophotography\950mm\Sh2-108 Take 3"
     #"W:\Astrophotography\950mm\NGC6979 - Pickerings Triangular Wisp"
     #"W:\Astrophotography\950mm\NGC 281 - Pacman"
     "W:\Astrophotography\950mm\vdb133"
)
$referenceImages = @(
    "M45 Take 2.L.31x90s.PSFSW.ESD.xisf"
    "Jellyfish Exhaust Take 2.Ha.72x360s.ESD.xisf"
    "LDN 1551 1546 LBN 821.SL.L.32x180s.R.20x180s.G.20x180s.B.25x180s.xisf"
    "ThreePanelRef.xisf"
    "FullRef.ABCDEF.xisf"
    "The Coma Cluster.L.Best.18.of.63x180s.PSFSW.ESD.xisf"
    "Sh2-261 - Lowers Nebula.Ha.47x360s.PSFSW.ESD.LSPR.xisf"
    "M35 and NGC2158.Ha.11x180s.Ha.61x360s.ESD.LN.LSPR.xisf"
    "NGC 3344 Take 2.L.65x180s.ESD.LN.LSPR.xisf"
    "M51.L.7x360s.ESD.xisf"
    "NGC 3628 - Hamburger Galaxy.L.38x360s.ESD.LSPR.xisf"
    "Dolphin Nebula in Cygnus.Ha.19x360s.ESD.xisf"
    "M106 Take 3.L.36x180s.ESD.LSPR.xisf"
    "Sadr Take 3g.L.57x30s.ESD.LSPR.xisf"
    "Bubble Nebula.L.22x180s.ESD.xisf"
    "vdB 15.L.95x30s.L.43x180s.ESD.xisf"
    "LDN 1235 - Dark Shark Nebula Take 2.L.89x180s.ESD.LSPR.LN.xisf"
    "M 33.B.24x180s.ESD.LSPR.xisf"
    "Flaming Star Nebula.L.65x180s.ESD.xisf"
    "M81 and M82.L.120x180s.ESD.xisf"
    "_Barnard 3 and 4.L.128x180s.L.27x360s.ESD.NewFlats.xisf"
    "Cone and Fox Fur.Ha3nm.16x360s.ESD.xisf"
    "PK 164 31.1.Ha3nm.155x360s.ESD.LSPR.LN.xisf"
    "NGC 2633 and NGC 2634.R.120x360s.LF.LN.LSPR.xisf"
    "LBN415 feature in Draco.B.153x360s.LF.LSPR.xisf"
    "Barnard 343 Region in Cygnus.R.18x360s.LF.xisf"
    "Sh2-108 Take 3.R.29x360s.LF.LSPR.xisf"
    "NGC6979 - Pickerings Triangular Wisp.R.23x360s.ESD.LSPR.xisf"
    "vdb133.L.57x360s.LF.LSPR.xisf"
    "NGC 281 - Pacman.R.60x360s.LF.LSPR.xisf"
    "vdB 10.L.68x180s.ESD.LSPR.xisf"
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
        Write-Warning "No alignment reference was specified... a new reference will automatically be selected."
        Wait-Event -Timeout 5
    }


    

    $rawSubs = 
        Get-XisfLightFrames -Path $target -Recurse -UseCache -SkipOnError -TruncateFilterBandpass:$false |
        where-object Instrument -eq "QHY268M" |
        Where-Object {-not $_.HasTokensInPath(@("reject","process","planning","testing","clouds","draft","cloudy","_ez_LS_","drizzle","quick","abandoned"))} |
        #Where-Object Filter -In @("Ha") |
        #Where-Object {-not $_.Filter.Contains("Oiii")} |
        #Where-Object Exposure -eq 10 |
        #Where-Object FocalRatio -eq "5.6" |
        #Where-Object Filter -ne "Oiii" |
        #Where-Object Filter -eq "L" |
        #Where-Object Exposure -eq 180 |
        #Where-Object Filter -ne "G" |
        #Where-Object Filter -ne "B" |
        #Where-object ObsDateMinus12hr -eq ([DateTime]"2022-11-09")
        Where-Object {-not $_.IsIntegratedFile()} #|
        #select-object -First 30
    #$rawSubs|Format-Table Path,*
    $rawSubs | Group-Object {[decimal]::Round($_.Rotator,1,([System.MidpointRounding]::AwayFromZero))} | sort-object Count

    <#
    Get-XisfFile -Path  D:\Backups\Camera\Dropoff\NINA |
        where-object ImageType -eq "FLAT" |
        Group-Object {[decimal]::Round($_.Rotator,1,([System.MidpointRounding]::AwayFromZero))} | sort-object Count
        #>

    #return
    $uncalibrated = 
        $rawSubs |
        Get-XisfCalibrationState `
            -CalibratedPath "E:\Calibrated\950mm" `
            -Verbose -ShowProgress -ProgressTotalCount ($rawSubs.Count) |
        foreach-object {
            $x = $_
            if(-not $x.IsCalibrated()){
                $x
            }
            else {
                #$x
            }
        } 

    if($uncalibrated){
        if((Read-Host -Prompt "Found $($uncalibrated.Count) uncalibrated files. Relocate to dropoff?") -eq "Y"){
            $uncalibrated |
                foreach-object {
                    Move-Item $_.Path "D:\Backups\Camera\Dropoff\NINA" -verbose
                }
        }
        exit
    }

    $createSuperLum=$false
    $data=Invoke-XisfPostCalibrationMonochromeImageWorkflow `
        -RawSubs $rawSubs `
        -CalibrationPath "E:\Calibrated\950mm" `
        -CorrectedOutputPath "S:\PixInsight\Corrected" `
        -WeightedOutputPath "S:\PixInsight\Weighted" `
        -DarkLibraryPath "W:\Astrophotography\DarkLibrary\QHY268M" `
        -AlignedOutputPath "S:\PixInsight\Aligned" `
        -BackupCalibrationPaths @("M:\PixInsightLT\Calibrated","S:\PixInsightLT\Calibrated") `
        -PixInsightSlot 200 `
        -RerunCosmeticCorrection:$false `
        -SkipCosmeticCorrection:$false `
        -RerunWeighting:$false `
        -SkipWeighting:$false `
        -PSFSignalWeightWeighting:$true `
        -RerunAlignment:$false `
        -IntegratedImageOutputDirectory $target `
        -AlignmentReference $alignmentReference `
        -GenerateDrizzleData `
        -ApprovalExpression "Median<42 && FWHM<5.5 && Stars > 2200" `
        -WeightingExpression "(15*(1-(FWHM-FWHMMin)/(FWHMMax-FWHMMin))
        +  5*(1-(Eccentricity-EccentricityMin)/(EccentricityMax-EccentricityMin))
        + 15*(SNRWeight-SNRWeightMin)/(SNRWeightMax-SNRWeightMin)
        + 30*(1-(Median-MedianMin)/(MedianMax-MedianMin))
        + 20*(Stars-StarsMin)/(StarsMax-StarsMin))
        + 20" `
        -Rejection "LinearFit" `
        -GenerateThumbnail `
        -Verbose -HotAutoSigma 4.0 #-KeepOpen 
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
                    -PixInsightSlot 201 `
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

}