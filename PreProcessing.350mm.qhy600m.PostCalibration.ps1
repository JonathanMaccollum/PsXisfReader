Clear-Host
#update-module PsXisfReader
if (-not (get-module psxisfreader)){import-module psxisfreader}
$ErrorActionPreference="STOP"
$VerbosePreference="Continue"
$targets = @(
     #"E:\Astrophotography\350mm\Cygnus on HD192985"
     #"E:\Astrophotography\350mm\Ring Nebula Widefield Take 3"
     #"E:\Astrophotography\350mm\Leo Triplet Widefield"
     #"E:\Astrophotography\350mm\Veil Region"
     #"E:\Astrophotography\350mm\Lobster Claw Region"
     #"E:\Astrophotography\350mm\Kembles Cascade"
     #"E:\Astrophotography\350mm\Angler and Dark Shark"
     #"E:\Astrophotography\350mm\vdB 14 and vdb 15"
     #"E:\Astrophotography\350mm\Barnard 22"
     "E:\Astrophotography\350mm\Crescent Region"
     #"E:\Astrophotography\350mm\Flaming Star Region"
     #"E:\Astrophotography\350mm\Lynx Panel 1"
     #"E:\Astrophotography\350mm\Lynx Panel 2"
     #"E:\Astrophotography\350mm\Lynx Panel 3"
     #"E:\Astrophotography\350mm\Lynx Panel 4"
     #"E:\Astrophotography\350mm\Lynx Panel 5"
     #"E:\Astrophotography\350mm\Lynx Panel 6"
     #"E:\Astrophotography\350mm\Lynx Panel 7"
     #"E:\Astrophotography\350mm\Lynx Panel 8"
     #"E:\Astrophotography\350mm\Lynx Panel 9"
     #"E:\Astrophotography\350mm\Lynx Panel 10"
     #"E:\Astrophotography\350mm\NGC 7000 Region Take 2"
     #"E:\Astrophotography\350mm\M 33"
     #"E:\Astrophotography\350mm\Barnard 23 24 and 26"
     #"E:\Astrophotography\350mm\LDN 1472 Barnard 3 and NGC1333"
     #"E:\Astrophotography\350mm\M45 Region"
     #"E:\Astrophotography\350mm\Sh2-240"
     #"E:\Astrophotography\350mm\Volcano Nebula and M81 M82 Panel 1"
     #"E:\Astrophotography\350mm\Volcano Nebula and M81 M82 Panel 2"
     #"E:\Astrophotography\350mm\Volcano Nebula and M81 M82 Panel 3"
     #"E:\Astrophotography\350mm\Melotte 20"
     #"E:\Astrophotography\350mm\Melotte 111 Take 2"
     #"E:\Astrophotography\350mm\LBN 406 in Draco"
)
$referenceImages = @(
     "Cygnus on HD192985.Ha6nmMaxFR.46x360s.LF.LSPR.xisf"
     "Leo Triplet Widefield.L3.108x360s.ESD.xisf"
     "Ring Nebula Widefield Take 3.R.71x60s.ESD.xisf"
     "Veil Region.Ha5nm.30x360s.ESD.xisf"
     "Lobster Claw Region.Ha5nm.31x360s.ESD.xisf"
     "Angler and Dark Shark.L.50x180s.ESD.xisf"
     "vdB 14 and vdb 15.L.53x180s.ESD.LSPR.xisf"
     "Barnard 22.L.80x180s.ESD.LSPR.xisf"
     "Crescent Region.Ha3nm.10x360s.ESD.xisf"
     "Flaming Star Region.L.23x180s.ESD.LSPR.xisf"
     "Reference.xisf"
     "NGC 7000 Region Take 2.Ha3nm.29x360s.nocc.PSFSW.ESD.LSPR.xisf"
     "Barnard 23 24 and 26.L.138x180s.ESD.LN.LSPR.xisf"
     "LDN 1472 Barnard 3 and NGC1333.L.109x180s.ESD.LSPR.xisf"
     "M 33.Oiii5nm.22x360s.ESD.xisf"
     "M45 Region.R.48x180s.ESD.LSPR.xisf"
     "Sh2-240.Ha3nm.22x360s.ESD.xisf"
     #"Volcano Nebula and M81 M82 Panel 1.L.145x180s.ESD.LSPR.xisf"
     #"Volcano Nebula and M81 M82 Panel 2.L.142x180s.ESD.LSPR.xisf"
     #"Volcano Nebula and M81 M82 Panel 3.L.87x180s.ESD.LSPR.xisf"
     "MosaicByCoords.Framing.P1.xisf"
     "MosaicByCoords.Framing.P2.xisf"
     "MosaicByCoords.Framing.P3.xisf"
     "Melotte 111 Take 2.G.2x180s.G.4x360s.LF.xisf"
     "LBN 406 in Draco.G.11x360s.LF.xisf"
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
        #Where-Object Filter -eq "Oiii5nm" |
        #Where-Object {-not $_.Filter.Contains("Oiii")} |
        #Where-Object Exposure -eq 360 |
        #Where-Object FocalRatio -eq "5.6" |
        #Where-Object Filter -in @("R") |
        #Where-Object Gain -eq 26 |
        #Where-object ObsDateMinus12hr -eq ([DateTime]"2024-02-20")
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
    $rejectionMethod="Rejection_ESD"
    #$rejectionMethod="LinearFit"
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
        -Rejection $rejectionMethod `
        -GenerateThumbnail `
        -Verbose -HotAutoSigma 4.0
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
