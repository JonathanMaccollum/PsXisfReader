Clear-Host
#update-module PsXisfReader
if (-not (get-module psxisfreader)){import-module psxisfreader -force}
$ErrorActionPreference="STOP"
$VerbosePreference="Continue"
$targets = @(
     #"E:\Astrophotography\90mm\Crescent Oxygen River"
     #"E:\Astrophotography\90mm\Veil Widefield"
     #"E:\Astrophotography\90mm\Lobster Claw and Cave"     
     #"E:\Astrophotography\90mm\LDN 1472 in Perseus"
     #"E:\Astrophotography\90mm\Anglerfish Dark Shark Rosebud Panel 2"
     #"E:\Astrophotography\90mm\Anglerfish Dark Shark Rosebud Panel 1"
     #"E:\Astrophotography\90mm\LBN 437 - Gecko Nebula Widefield"
     #"E:\Astrophotography\90mm\Cepheus on Barnard 170"
     #"E:\Astrophotography\90mm\Cygnus near Sh2-115"
     "E:\Astrophotography\90mm\Cygnus near DWB111"
     #"E:\Astrophotography\90mm\Heart and Soul"
     #"E:\Astrophotography\90mm\Flaming Star Region"
     #"E:\Astrophotography\90mm\Cepheus near CTB-1"
     #"E:\Astrophotography\90mm\NGC1333 Region"
     #"E:\Astrophotography\90mm\M81 M82 Region"
     #"E:\Astrophotography\90mm\Cygnus Crack and NA Nebula Region"
     #"E:\Astrophotography\90mm\C2022 E3 ZTF Widefield"

)
$referenceImages = @(
    "Crescent Oxygen River.Ha.9x360s.ESD.xisf"
    "Veil Widefield.Ha.30x360s.ESD.xisf"
    "Lobster Claw and Cave.Ha.13x360s.ESD.xisf"
    "Anglerfish Dark Shark Rosebud Panel 2.L.30x180s.ESD.xisf"
    "Anglerfish Dark Shark Rosebud Panel 1.L.40x180s.ESD.LN.xisf"
    "LDN 1472 in Perseus.L.16x90s.L.21x180s.ESD.xisf"
    "Cepheus on Barnard 170.L.14x180s.ESD.xisf"
    "Cygnus near Sh2-115.Oiii.84x360s.ESD.LN.xisf"
    "Cygnus near DWB111.Ha.2x360s.Ha.20x900s.ESD.LN.xisf"
    "Heart and Soul.Ha.37x360s.Ha.18x900s.ESD.LN.xisf"
    "Cepheus near CTB-1.R.19x180s.ESD.xisf"
    "Flaming Star Region.L.67x180s.ESD.xisf"
    "NGC1333 Region.L.15x360s.ESD.xisf"
    "Cygnus Crack and NA Nebula Region.R.17x360s.ESD.Drizzled.1x.xisf"
    "M81 M82 Region.L.46x360s.ESD.xisf"
    "C2022 E3 ZTF Widefield.L.30x180s.ESD.xisf"
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
        where-object Instrument -eq "ZWO ASI533MM Pro" |
        Where-Object {-not $_.HasTokensInPath(@("reject","process","planning","testing","clouds","draft","cloudy","_ez_LS_","drizzle","quick"))} |
        #Where-Object Filter -In @("L","Sii3") |
        #Where-Object {-not $_.Filter.Contains("Sii")} |
        #Where-Object Exposure -eq 180 |
        #Where-Object FocalRatio -eq "5.6" |
        #Where-Object Filter -eq "L" |
        #Where-object ObsDateMinus12hr -eq ([DateTime]"2021-05-05")
        Where-Object {-not $_.IsIntegratedFile()} #|
        #select-object -First 30
    #$rawSubs|Format-Table Path,*

    $uncalibrated = 
        $rawSubs |
        Get-XisfCalibrationState `
            -CalibratedPath "E:\Calibrated\90mm" `
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
            exit
        }
        
    }

    $createSuperLum=$false
    $data=Invoke-XisfPostCalibrationMonochromeImageWorkflow `
        -RawSubs $rawSubs `
        -CalibrationPath "E:\Calibrated\90mm" `
        -CorrectedOutputPath "S:\PixInsight\Corrected" `
        -WeightedOutputPath "S:\PixInsight\Weighted" `
        -DarkLibraryPath "E:\Astrophotography\DarkLibrary\ZWO ASI533MM Pro" `
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
        -ApprovalExpression "Median<175 && FWHM<5.5 && Stars > 2200" `
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
            $outputFileName = $reference.Object.Trim()
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

}