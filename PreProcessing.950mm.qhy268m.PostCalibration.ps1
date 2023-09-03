Clear-Host
#update-module PsXisfReader
if (-not (get-module psxisfreader)){import-module psxisfreader}
$ErrorActionPreference="STOP"
$VerbosePreference="Continue"
$targets = @(
     #"E:\Astrophotography\950mm\M45 Take 2"
     #"E:\Astrophotography\950mm\Jellyfish Exhaust Take 2"
     #"E:\Astrophotography\950mm\LDN 1551 1546 LBN 821"
     #"E:\Astrophotography\950mm\C2022 E3 ZTF Night8"
     #"E:\Astrophotography\950mm\C2022 E3 ZTF Night11a"
     #"E:\Astrophotography\950mm\C2022 E3 ZTF Night11b"
     #"E:\Astrophotography\950mm\C2022 E3 ZTF Night11c"
     #"E:\Astrophotography\950mm\C2022 E3 ZTF Night12a"
     #"E:\Astrophotography\950mm\C2022 E3 ZTF Night12b"
     #"E:\Astrophotography\950mm\C2022 E3 ZTF Night12c"
     #"E:\Astrophotography\950mm\C2022 E3 ZTF Night12d"
     #"E:\Astrophotography\950mm\C2022 E3 ZTF Night12e"
     #"E:\Astrophotography\950mm\C2022 E3 ZTF Night12f"
     #"E:\Astrophotography\950mm\The Coma Cluster"
     #"E:\Astrophotography\950mm\Sh2-261 - Lowers Nebula"
     #"E:\Astrophotography\950mm\T Coronae Borealis"
     #"E:\Astrophotography\950mm\Arcturus"
     #"E:\Astrophotography\950mm\M35 and NGC2158"
     #"E:\Astrophotography\950mm\NGC 3344 Take 2"
     #"E:\Astrophotography\950mm\M51"
     #"E:\Astrophotography\950mm\NGC 3628 - Hamburger Galaxy"
     "E:\Astrophotography\950mm\Dolphin Nebula in Cygnus"
     "E:\Astrophotography\950mm\M106 Take 3"
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
        Get-XisfLightFrames -Path $target -Recurse -UseCache -SkipOnError |
        where-object Instrument -eq "QHY268M" |
        Where-Object {-not $_.HasTokensInPath(@("reject","process","planning","testing","clouds","draft","cloudy","_ez_LS_","drizzle","quick"))} |
        #Where-Object Filter -In @("Ha") |
        #Where-Object {-not $_.Filter.Contains("Oiii")} |
        #Where-Object Exposure -eq 10 |
        #Where-Object FocalRatio -eq "5.6" |
        #Where-Object Filter -eq "L" |
        #Where-Object Filter -ne "R" |
        #Where-Object Filter -ne "G" |
        #Where-Object Filter -ne "B" |
        #Where-object ObsDateMinus12hr -eq ([DateTime]"2022-11-09")
        Where-Object {-not $_.IsIntegratedFile()} #|
        #select-object -First 30
    #$rawSubs|Format-Table Path,*

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
        -DarkLibraryPath "E:\Astrophotography\DarkLibrary\QHY268M" `
        -AlignedOutputPath "S:\PixInsight\Aligned" `
        -BackupCalibrationPaths @("M:\PixInsightLT\Calibrated","S:\PixInsightLT\Calibrated") `
        -PixInsightSlot 200 `
        -RerunCosmeticCorrection:$false `
        -SkipCosmeticCorrection:$false `
        -RerunWeighting:$false `
        -SkipWeighting:$false `
        -PSFSignalWeightWeighting `
        -RerunAlignment:$false `
        -IntegratedImageOutputDirectory $target `
        -AlignmentReference $alignmentReference `
        -GenerateDrizzleData `
        -ApprovalExpression "Median<42 && FWHM<5.5 && Stars > 400" `
        -WeightingExpression "PSFSignalWeight" `
        -Rejection "Rejection_ESD" `
        -GenerateThumbnail `
        -Verbose #-KeepOpen
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