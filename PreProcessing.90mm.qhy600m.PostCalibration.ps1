Clear-Host
#update-module PsXisfReader
if (-not (get-module psxisfreader)){import-module psxisfreader}
$ErrorActionPreference="STOP"
$VerbosePreference="Continue"
$targets = @(
     #"E:\Astrophotography\90mm\Lyra on Lyr A"
     #"E:\Astrophotography\90mm\Cygnus on HD192143"
     #"E:\Astrophotography\90mm\Scorpius on TYC 6202-0266-1"
     #"E:\Astrophotography\90mm\Ursa Minor"
     #"E:\Astrophotography\90mm\Sag near GSC 6265-2391"
     #"E:\Astrophotography\90mm\Cygnus to Vulpecula"
     #"E:\Astrophotography\90mm\Cygnus on HD195405"
     #"E:\Astrophotography\90mm\Cepheus near CTB-1"
     #"E:\Astrophotography\90mm\Orion on 48 Ori"
     #"E:\Astrophotography\90mm\String of Pearls in Monoceros"
     #"E:\Astrophotography\90mm\Beehive Cluster Widefield"
     #"E:\Astrophotography\90mm\LDN 1472 in Perseus"
     #"E:\Astrophotography\90mm\C2022 E3 ZTF Widefield Take 9 (Hyades)"
     "E:\Astrophotography\90mm\M81 M82 Region"
)
$referenceImages = @(
    "Lyra on Lyr A.R.20x180s.ESD.OvercorrectedFlats.xisf"
    "Cygnus on HD192143.Ha6nmMaxFR.14x180s.ESD.xisf"
    "Scorpius on TYC 6202-0266-1.Ha6nmMaxFR.55x180s.ESD.xisf"
    "Sag near GSC 6265-2391.Ha6nmMaxFR.8x180s.ESD.xisf"
    "Cygnus on HD195405.Ha6nmMaxFR.16x180s.PSFSW.LF.xisf"
    "Cygnus to Vulpecula.R.11x90s.noweights.ESD.xisf"
    "Cepheus near CTB-1.Ha6nmMaxFR.23x720s.ESD.LN.xisf"
    "Orion on 48 Ori.Ha6nmMaxFR.29x720s.ESD.LN.xisf"
    "LDN 1472 in Perseus.L3.36x360s.ESD.xisf"
    "String of Pearls in Monoceros.Ha6nmMaxFR.18x180s.Ha6nmMaxFR.20x720s.ESD.LN.xisf"
    "M81 M82 Region.L3.29x360s.ESD.xisf"
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
        where-object Instrument -eq "QHY600M" |
        Where-Object {-not $_.HasTokensInPath(@("reject","process","planning","testing","clouds","draft","cloudy","_ez_LS_","drizzle","quick"))} |
        where-object Geometry -eq "9576:6388:1" |
        #Where-Object Filter -In @("R","B","G") |
        #Where-Object {-not $_.Filter.Contains("Oiii")} |
        #Where-Object Exposure -eq 180 |
        #Where-Object FocalRatio -eq "5.6" |
        Where-Object Filter -eq "L3" |
        Where-Object Gain -eq 26 |
        #Where-object ObsDateMinus12hr -eq ([DateTime]"2021-05-05")
        Where-Object {-not $_.IsIntegratedFile()} #|
        #select-object -First 30
    #$rawSubs|Format-Table Path,*

    $uncalibrated = 
        $rawSubs |
        Get-XisfCalibrationState `
            -CalibratedPath "E:\Recalibrated\90mm" `
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

    # if($uncalibrated){
    #     if((Read-Host -Prompt "Found $($uncalibrated.Count) uncalibrated files. Relocate to dropoff?") -eq "Y"){
    #         $uncalibrated |
    #             foreach-object {
    #                 Move-Item $_.Path "D:\Backups\Camera\Dropoff\NINA" -verbose
    #             }
    #     }
    #     exit
    # }

    $createSuperLum=$false
    $data=Invoke-XisfPostCalibrationMonochromeImageWorkflow `
        -RawSubs $rawSubs `
        -CalibrationPath "E:\Recalibrated\90mm" `
        -CorrectedOutputPath "S:\PixInsight\Corrected" `
        -WeightedOutputPath "S:\PixInsight\Weighted" `
        -DarkLibraryPath "E:\Astrophotography\DarkLibrary\QHY600M" `
        -AlignedOutputPath "S:\PixInsight\Aligned" `
        -BackupCalibrationPaths @("M:\PixInsightLT\Calibrated","S:\PixInsightLT\Calibrated") `
        -PixInsightSlot 201 `
        -RerunCosmeticCorrection:$false `
        -SkipCosmeticCorrection:$false `
        -RerunWeighting:$false `
        -SkipWeighting:$false `
        -PSFSignalWeightWeighting `
        -RerunAlignment:$false `
        -IntegratedImageOutputDirectory $target `
        -AlignmentReference $alignmentReference `
        -GenerateDrizzleData `
        -ApprovalExpression "Median<42 && FWHM<5.5 && Stars > 8000" `
        -WeightingExpression "PSFSignalWeight" `
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

$data=Import-Clixml -Path "E:\Astrophotography\90mm\M81 M82 Region\Stats.20230316 030934.clixml" 
$data|where-object Aligned -ne $null | Foreach-Object {$_.Stats}