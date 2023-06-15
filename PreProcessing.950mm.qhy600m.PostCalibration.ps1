Clear-Host
#update-module PsXisfReader
if (-not (get-module psxisfreader)){import-module psxisfreader}
$ErrorActionPreference="STOP"
$VerbosePreference="Continue"
$targets = @(
     #"E:\Astrophotography\950mm\Sadr Take 2b"
     #"E:\Astrophotography\950mm\Eye of Smaug Take 2"
     #"E:\Astrophotography\950mm\Eye of Smaug Take 2P2"
     #"E:\Astrophotography\950mm\NGC7129 NGC7142 Take 2b"
     #"E:\Astrophotography\950mm\vdb131 vdb132 Take 2"
     #"E:\Astrophotography\950mm\sh2-115 Take 2"
     #"E:\Astrophotography\950mm\NGC1333 Take 2"
     #"E:\Astrophotography\950mm\LDN 1251 Take 2"
     #"E:\Astrophotography\950mm\LBN 437 - Gecko Nebula"
     #"E:\Astrophotography\950mm\Abell 85"
     #"E:\Astrophotography\950mm\IC5068 - Crack in Cygnus Panel 1"
     #"E:\Astrophotography\950mm\IC5068 - Crack in Cygnus Panel 2"
     #"E:\Astrophotography\950mm\Spider and Fly Take 3"
     #"E:\Astrophotography\950mm\Properller and Smaug Panel 1"
     #"E:\Astrophotography\950mm\Properller and Smaug Panel 2"
     #"E:\Astrophotography\950mm\2022 SN Candidate in Cygnus"
     #"E:\Astrophotography\950mm\Soul Take 3 Panel 1"
     #"E:\Astrophotography\950mm\Soul Take 3 Panel 2"
     #"E:\Astrophotography\950mm\C2022 E3 ZTF"
     #"E:\Astrophotography\950mm\Abell 31 Take 2"
     #"E:\Astrophotography\950mm\NGC2787 Take 2"
     "E:\Astrophotography\950mm\Crescent Nebula"
     #"E:\Astrophotography\950mm\M 101"
     #"E:\Astrophotography\950mm\ARP 84 (NGC 5394 and NGC 5395)"
     #"E:\Astrophotography\950mm\Bright Stars in Draco"
     #"E:\Astrophotography\950mm\NGC 5935 NGC 5945 NGC 5943"
     #"E:\Astrophotography\950mm\Melotte 111"
     #"E:\Astrophotography\950mm\wr134 Take 3"
)
$referenceImages = @(
    "Sadr Take 2b.Ha.64x180s.ESD.xisf"
    "vdb131 vdb132 Take 2.Ha.49x180s.ESD.xisf"
    "Eye of Smaug Take 2.Ha.29x180s.ESD.xisf"
    "Eye of Smaug Take 2P2.Ha.10x180s.ESD.xisf"
    "Properller and Smaug Panel 1.Oiii.21x360s.PSFSW.ESD.xisf"
    "Properller and Smaug Panel 2.Oiii.18x360s.PSFSW.ESD.xisf"
    "sh2-115 Take 2.Oiii.32x360s.ESD.xisf"
    "LDN 1251 Take 2.Superlum.L.191x180s.R.127x180s.G.114x180s.B.121x180s.LF.45degFlats.xisf"
    "NGC7129 NGC7142 Take 2b.Ha.63x360s.PSFSW.ESD.xisf"
    "LBN 437 - Gecko Nebula.L.13x180s.PSFSW.ESD.xisf"
    "IC5068 - Crack in Cygnus Panel 1.Ha.25x360s.ESD.xisf"
    "IC5068 - Crack in Cygnus Panel 2.Ha.29x360s.ESD.xisf"
    "NGC1333 Take 2.L.180x180s.ESD.xisf"
    "Abell 85.Ha.113x360s.PSFSW.ESD.xisf"
    "Spider and Fly Take 3.L.68x180s.PSFSW.ESD.xisf"
    "2022 SN Candidate in Cygnus.R.96x10s.nocc.ESD.xisf"
    "Soul Take 3 Panel 1.Ha.95x360s.PSFSW.ESD.xisf"
    "Soul Take 3 Panel 2.R.49x180s.PSFSW.ESD.xisf"
    "Crescent Nebula.Ha.53x360s.LF.xisf"
    "M 101.SL.LRGB.L.6x180s.R.22x90s.R.8x180s.G.16x90s.G.8x180s.B.20x90s.B.8x180s.AllSubs.LF.LSPR.xisf"
    "ARP 84 (NGC 5394 and NGC 5395).SL.L.21x180s.R.44x180s.G.32x180s.B.29x180s.ESD.LSPR.xisf"
    "NGC 5935 NGC 5945 NGC 5943.R.16x180s.ESD.LSPR.xisf"
    "wr134 Take 3.Ha.Best25.of.89x360s.ESD.xisf"
    "Melotte 111.R.15x90s.ESD.xisf"
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
        #Where-Object Filter -In @("Ha") |
        #Where-Object {-not $_.Filter.Contains("Oiii")} |
        #Where-Object Exposure -eq 10 |
        #Where-Object FocalRatio -eq "5.6" |
        Where-Object Filter -ne "L" |
        #Where-Object Filter -eq "R" |
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
        -DarkLibraryPath "E:\Astrophotography\DarkLibrary\QHY600M" `
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
        -ApprovalExpression "Median<42 && FWHM<5.5 && Stars > 2200" `
        -WeightingExpression "(15*(1-(FWHM-FWHMMin)/(FWHMMax-FWHMMin))
        +  5*(1-(Eccentricity-EccentricityMin)/(EccentricityMax-EccentricityMin))
        + 15*(SNRWeight-SNRWeightMin)/(SNRWeightMax-SNRWeightMin)
        + 30*(1-(Median-MedianMin)/(MedianMax-MedianMin))
        + 20*(Stars-StarsMin)/(StarsMax-StarsMin))
        + 20" `
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