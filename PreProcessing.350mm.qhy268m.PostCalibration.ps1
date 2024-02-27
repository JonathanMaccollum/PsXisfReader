Clear-Host
#update-module PsXisfReader
if (-not (get-module psxisfreader)){import-module psxisfreader}
$ErrorActionPreference="STOP"
$VerbosePreference="Continue"
$targets = @(
    #"E:\Astrophotography\350mm\Draco near SAO HD 133229"
    #"E:\Astrophotography\350mm\Cygnus - wr134 to Tulip"
    #"E:\Astrophotography\350mm\Cygnus - wr134 to Tulip South 1"
    #"E:\Astrophotography\350mm\Cygnus - wr134 to Tulip North 1"
    #"E:\Astrophotography\350mm\M 101"
    #"E:\Astrophotography\350mm\Crescent and sh2-108"
    #"E:\Astrophotography\350mm\Melotte 111"
    #"E:\Astrophotography\350mm\Draco Trio Widefield"
    #"E:\Astrophotography\350mm\sh2-73"
    #"E:\Astrophotography\350mm\Smaug Mosaic - 01"
    #"E:\Astrophotography\350mm\Smaug Mosaic - 02"
    #"E:\Astrophotography\350mm\Smaug Mosaic - 03"
    #"E:\Astrophotography\350mm\Smaug Mosaic NW1 - 01"
    #"E:\Astrophotography\350mm\Smaug Mosaic NW1 - 02"
    #"E:\Astrophotography\350mm\Smaug Mosaic NW1 - 03"
    #"E:\Astrophotography\350mm\Smaug Mosaic SE1 - 01"
    #"E:\Astrophotography\350mm\Smaug Mosaic SE1 - 02"
    #"E:\Astrophotography\350mm\Smaug Mosaic SE1 - 03"
    #"E:\Astrophotography\350mm\LBN 437 - Gecko Nebula"
    #"E:\Astrophotography\350mm\Abell 85"
    "E:\Astrophotography\350mm\NGC7129 Region Take 2"
    #"E:\Astrophotography\350mm\Spider and Fly Region"
    #"E:\Astrophotography\350mm\NGC 7000 Region"
    #"E:\Astrophotography\350mm\Flaming Star Nebula"
    #"E:\Astrophotography\350mm\NGC 1333 Region"
)
$referenceImages = @(
    "Draco near SAO HD 133229.L3.21x180s.PSFSW.ESD.xisf"
    "Cygnus - wr134 to Tulip.L3.24x180s.ESD.LSPR.xisf"
    "Cygnus - wr134 to Tulip South 1.Ha6nmMaxFR.21x360s.ESD.xisf"
    "Cygnus - wr134 to Tulip North 1.Ha.28x360s.ESD.xisf"
    "M 101.L3.24x180s.PSFSW.ESD.LSPR.xisf"
    "Crescent and sh2-108.Ha6nmMaxFR.65x360s.PSFSW.ESD.xisf"
    "sh2-73.L.14x180s.L3.67x180s.ESD.LSPR.LN.xisf" #Meade #2 with larger stars
    "Smaug Mosaic - 01.Ha.16x360s.ESD_MosaicRef.xisf"
    "Smaug Mosaic - 02.Ha.16x360s.ESD_MosaicRef.xisf"
    "Smaug Mosaic - 03.Ha.12x360s.ESD_MosaicRef.xisf"
    "Smaug Mosaic SE1 - 01.Ha6nmMaxFR.73x360s.Ha3nm.76x360s.LF.LSPR.xisf"
    "Smaug Mosaic SE1 - 02.Ha6nmMaxFR.39x360s.Ha3nm.43x360s.ESD.LSPR.xisf"
    "Smaug Mosaic SE1 - 03.Ha6nmMaxFR.45x360s.Ha3nm.48x360s.LF.LSPR.xisf"
    "LBN 437 - Gecko Nebula.L.24x180s.ESD.xisf"
    "Ref.xisf"
    "Smaug Mosaic NW1 - 03.Ha6nmMaxFR.2x180s.Ha6nmMaxFR.88x360s.Ha3nm.15x360s.ESD.xisf"
    "Smaug Mosaic NW1 - 02.Ha6nmMaxFR.93x360s.ESD.LSPR.xisf"
    "Abell 85.Ha6nmMaxFR.18x360s.ESD.xisf"
    "NGC7129 Region Take 2.Oiii5nm.18x720s.Oiii6nm.32x360s.ESD.xisf"
    "NGC 1333 Region.L.78x180s.ESD.xisf"
    "Spider and Fly Region.Ha3nm.54x360s.ESD.xisf"
    "NGC 7000 Region.Ha3nm.21x360s.ESD.xisf"
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
        Get-XisfLightFrames -Path $target -Recurse -UseCache:$false -SkipOnError -TruncateFilterBandpass:$false |
        where-object Instrument -eq "QHY268M" |
        Where-Object {-not $_.HasTokensInPath(@("reject","process","planning","testing","clouds","draft","cloudy","_ez_LS_","drizzle","quick"))} |
        where-object Geometry -eq "6252:4176:1" |
        #where-object ObsDateMinus12hr -ge "2023-05-03" |
        #where-object Filter -eq "L" |
        #where-object Filter -in @("Ha3nm","Ha6nm") |
        #where-object Filter -ne "Sii6nm" |
        Where-Object {-not $_.IsIntegratedFile()} #|

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
        -DarkLibraryPath "E:\Astrophotography\DarkLibrary\QHY268M" `
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
