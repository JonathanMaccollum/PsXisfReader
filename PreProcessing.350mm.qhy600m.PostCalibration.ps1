Clear-Host
#update-module PsXisfReader
if (-not (get-module psxisfreader)){import-module psxisfreader}
$ErrorActionPreference="STOP"
$VerbosePreference="Continue"
$targets = @(
     #"W:\Astrophotography\350mm\Cygnus on HD192985"
     #"W:\Astrophotography\350mm\Ring Nebula Widefield Take 3"
     #"W:\Astrophotography\350mm\Leo Triplet Widefield"
     #"W:\Astrophotography\350mm\Veil Region"
     #"W:\Astrophotography\350mm\Lobster Claw Region"
     #"W:\Astrophotography\350mm\Kembles Cascade"
     #"W:\Astrophotography\350mm\Angler and Dark Shark"
     #"W:\Astrophotography\350mm\vdB 14 and vdb 15"
     #"W:\Astrophotography\350mm\Barnard 22"
     #"W:\Astrophotography\350mm\Crescent Region"
     #"W:\Astrophotography\350mm\Flaming Star Region"
     #"W:\Astrophotography\350mm\Lynx Panel 1"
     #"W:\Astrophotography\350mm\Lynx Panel 2"
     #"W:\Astrophotography\350mm\Lynx Panel 3"
     #"W:\Astrophotography\350mm\Lynx Panel 4"
     #"W:\Astrophotography\350mm\Lynx Panel 5"
     #"W:\Astrophotography\350mm\Lynx Panel 6"
     #"W:\Astrophotography\350mm\Lynx Panel 7"
     #"W:\Astrophotography\350mm\Lynx Panel 8"
     #"W:\Astrophotography\350mm\Lynx Panel 9"
     #"W:\Astrophotography\350mm\Lynx Panel 10"
     #"W:\Astrophotography\350mm\NGC 7000 Region Take 2"
     #"W:\Astrophotography\350mm\M 33"
     #"W:\Astrophotography\350mm\Barnard 23 24 and 26"
     #"W:\Astrophotography\350mm\LDN 1472 Barnard 3 and NGC1333"
     #"W:\Astrophotography\350mm\M45 Region"
     #"W:\Astrophotography\350mm\Sh2-240"
     #"W:\Astrophotography\350mm\Volcano Nebula and M81 M82 Panel 1"
     #"W:\Astrophotography\350mm\Volcano Nebula and M81 M82 Panel 2"
     #"W:\Astrophotography\350mm\Volcano Nebula and M81 M82 Panel 3"
     #"W:\Astrophotography\350mm\Melotte 20"
     #"W:\Astrophotography\350mm\Melotte 111 Take 2"
     #"W:\Astrophotography\350mm\LBN 406 in Draco"
     #"W:\Astrophotography\350mm\Sh2-129 Ou4 Barnard 148 354"
     #"W:\Astrophotography\350mm\Kappa Borealis"
     #"W:\Astrophotography\350mm\Hercules - Nebula and DolDzim7"
     #"W:\Astrophotography\350mm\Dirty Cygnus"
     #"W:\Astrophotography\350mm\Vulpecula - LDN 792 LDN 784 LDN 768 and Stock 1"
     #"W:\Astrophotography\350mm\Vulpecula - LDN 741 and Coat Hangar"
     #"W:\Astrophotography\350mm\Smaug Take 5"
     #"W:\Astrophotography\350mm\vdb133 Region"
     #"W:\Astrophotography\350mm\Pacman Region in Cassiopeia"
     #"W:\Astrophotography\350mm\Cepheus Region with Barnard 169-174"
     #"W:\Astrophotography\350mm\Cone Nebula Region"
    "W:\Astrophotography\350mm\Rosette Region"
    #"W:\Astrophotography\350mm\Soul Nebula"
)
$referenceImages = @(
     "Cygnus on HD192985.Ha6nmMaxFR.46x360s.LF.LSPR.xisf"
     "Leo Triplet Widefield.L3.108x360s.ESD.xisf"
     "Ring Nebula Widefield Take 3.R.71x60s.ESD.xisf"
     "Veil Region.Ha5nm.30x360s.ESD.xisf"
     "Lobster Claw Region.Ha5nm.31x360s.ESD.xisf"
     "Angler and Dark Shark.L.50x180s.ESD.xisf"
     #"vdB 14 and vdb 15.L.53x180s.ESD.LSPR.xisf"
     "Barnard 22.L.80x180s.ESD.LSPR.xisf"
     "Crescent Region.Ha3nm.10x360s.ESD.xisf"
     "Flaming Star Region.L.23x180s.ESD.LSPR.xisf"
     "Reference.xisf"
     "NGC 7000 Region Take 2.Ha3nm.29x360s.nocc.PSFSW.ESD.LSPR.xisf"
     "Barnard 23 24 and 26.L.138x180s.ESD.LN.LSPR.xisf"
     #"LDN 1472 Barnard 3 and NGC1333.L.109x180s.ESD.LSPR.xisf"
     "LDN 1472 Barnard 3 and NGC1333.L.117x720s.ESD.LSPR.xisf"
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
     "Melotte 20.G.12x360s.PC.LSPR.xisf"
     "Sh2-129 Ou4 Barnard 148 354.Ha3nm.8x360s.Ha3nm.18x720s.LF.xisf"
     "Kappa Borealis.R.33x360s.LF.LSPR.xisf"
     "Hercules - Nebula and DolDzim7.G.71x360s.ESD.xisf"
     "Dirty Cygnus.R.12x360s.ESD2.xisf"
     "Vulpecula - LDN 792 LDN 784 LDN 768 and Stock 1.R.27x360s.ESD.LSPR.xisf"
     "Vulpecula - LDN 741 and Coat Hangar.R.19x360s.ESD.LSPR.xisf"
     "Smaug Take 5.R.34x360s.ESD.LSPR.xisf"
     "Pacman Region in Cassiopeia.Ha3nm.74x720s.ESD.xisf"
     "vdb133 Region.G.24x720s.ESD.LSPR.xisf"
     "Cepheus Region with Barnard 169-174.Ha3nm.18x720s.ESD.xisf"
     "Cone Nebula Region.Ha3nm.28x720s.ESD.xisf"
     "vdB 14 and vdb 15.L.11x720s.ESD.xisf"
     "Soul Nebula.Ha3nm.61x720s.ESD.LSPR.xisf"
     "Rosette Region.L.18x720s.ESD.LSPR.xisf"
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
        Get-XisfLightFrames -Path $target -Recurse -SkipOnError -TruncateFilterBandpass:$false -UseCache:$true |
        where-object Instrument -eq "QHY600M" |
        Where-Object {-not $_.HasTokensInPath(@("reject","process","planning","testing","clouds","draft","cloudy","_ez_LS_","drizzle","quick"))} |
        where-object Geometry -eq "9576:6388:1" |
        #where-object ReadoutMode -eq "2CMS-0" |
        #Where-Object {-not $_.Filter.Contains("Oiii")} |
        #Where-Object Exposure -ge 360 |
        #Where-Object FocalRatio -eq "5.6" |
        #Where-Object Filter -ne "L" |
        #Where-Object Filter -in @('L','R','G','B') |
        #Where-Object Filter -in @('Ha3nm') |
        #Where-Object Gain -eq 26 |
        #Where-object ObsDateMinus12hr -eq ([DateTime]"2024-03-09")
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
                    Move-Item $_.Path "D:\Backups\Camera\Dropoff\NINACS" -verbose
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
        -DarkLibraryPath "W:\Astrophotography\DarkLibrary\QHY600M" `
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
        -ApprovalExpression "Median<40 && FWHM<1.65 && Stars > 8000" `
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
