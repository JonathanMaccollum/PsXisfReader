Clear-Host
#update-module PsXisfReader
if (-not (get-module psxisfreader)){import-module psxisfreader}
$ErrorActionPreference="STOP"
$VerbosePreference="Continue"
$targets = @(
     #"E:\Astrophotography\350mm\Cygnus on HD192985"
     "E:\Astrophotography\350mm\Ring Nebula Widefield Take 3"
     #"E:\Astrophotography\350mm\Leo Triplet Widefield"
)
$referenceImages = @(
     "Cygnus on HD192985.Ha6nmMaxFR.46x360s.LF.LSPR.xisf"
     "Leo Triplet Widefield.L3.108x360s.ESD.xisf"
     "Ring Nebula Widefield Take 3.R.71x60s.ESD.xisf"
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
        #Where-Object Filter -in @("Oiii6nm") |
        #Where-Object Gain -eq 26 |
        #Where-object ObsDateMinus12hr -eq ([DateTime]"2021-05-05")
        Where-Object {-not $_.IsIntegratedFile()} #|
        #select-object -First 30
    #$rawSubs|Format-Table Path,*

    $uncalibrated = 
        $rawSubs |
        Get-XisfCalibrationState `
            -CalibratedPath "E:\Calibrated\350mm" `
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
        -CalibrationPath "E:\Calibrated\350mm" `
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
        -ApprovalExpression "Median<142 && FWHM<5.5 && Stars > 1000" `
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
