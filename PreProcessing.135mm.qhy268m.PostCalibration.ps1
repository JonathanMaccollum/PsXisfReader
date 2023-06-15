Clear-Host
#update-module PsXisfReader
if (-not (get-module psxisfreader)){import-module psxisfreader}
$ErrorActionPreference="STOP"
$VerbosePreference="Continue"
$targets = @(
     "E:\Astrophotography\135mm\Cygnus on HD193701"
     #"E:\Astrophotography\135mm\Cassiopeia on HD 443 with CED 214"
     #"E:\Astrophotography\135mm\Heart and Double Cluster in Cassiopeia"
     #"E:\Astrophotography\135mm\sh2-240"
     #"E:\Astrophotography\135mm\sh2-202 in Cassiopeia"
     #"E:\Astrophotography\135mm\DL Camelopardalis"
     #"E:\Astrophotography\135mm\Taurus Molecular Cloud with vdb31"
     #"E:\Astrophotography\135mm\Comet C2021 A1 Leonard Framing 4"
     #"E:\Astrophotography\135mm\sh2-202 in Cassiopeia"
     #"E:\Astrophotography\135mm\Cepheus on HD204211"
     #"E:\Astrophotography\135mm\Abell 31"
     #"E:\Astrophotography\135mm\Cepheus on HD218724"
     #"E:\Astrophotography\135mm\Cone to Sh2-284 Panel 1"
     #"E:\Astrophotography\135mm\Cone to Sh2-284 Panel 2"
     #"E:\Astrophotography\135mm\Orion on HD290890"
     #"E:\Astrophotography\135mm\NGC3344 at 135mm"
     #"E:\Astrophotography\135mm\m101 at 135mm P1"
     #"E:\Astrophotography\135mm\m101 at 135mm P2"
     #"E:\Astrophotography\135mm\m101 at 135mm P3"
)
$referenceImages = @(
    "Cygnus on HD193701.Ha.80x180s.ESD.xisf"
    "Cassiopeia on HD 443 with CED 214.Ha.39x180s.ESD.xisf"
    "Heart and Double Cluster in Cassiopeia.Ha.69x180s.ESD.xisf"
    "Sh2-240.Ha.111x180s.ESD.xisf"
    "DL Camelopardalis.Ha.27x180s.ESD.xisf"
    "sh2-202 in Cassiopeia.Ha.64x180s.ESD.xisf"
    "Taurus Molecular Cloud with vdb31.L.62x90s.ESD.xisf"
    "Comet C2021 A1 Leonard with M3.G.84x45s.ESD.xisf"
    "Cepheus on HD218724.Sii3.100x180s.ESD.xisf"
    "Cepheus on HD204211.Ha.112x180s.ESD.xisf"
    "Cone to Sh2-284 Panel 1.Ha.44x180s.ESD.xisf"
    "Cone to Sh2-284 Panel 2.Ha.55x180s.ESD.xisf"
    "Abell 31.R.90x90s.ESD.xisf"
    "_NGC3344 at 135mm.L.276x90s.ESD.xisf"
    "m101 at 135mm P1.L.42x90s.ESD.xisf"
    "m101 at 135mm P2.L.128x90s.ESD.xisf"
    "m101 at 135mm P3.L.106x90s.ESD.xisf"
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
        #where-object Instrument -eq "QHYCCD-Cameras-Capture (ASCOM)" |
        where-object Instrument -eq "QHY268M" |
        Where-Object {-not $_.HasTokensInPath(@("reject","process","planning","testing","clouds","draft","cloudy","_ez_LS_","drizzle","quick"))} |
        #Where-Object Filter -In @("R","B","G") |
        #Where-Object {-not $_.Filter.Contains("Oiii")} |
        #Where-Object Filter -ne "V4" |
        #Where-Object Filter -eq "R" |
        #Where-Object Filter -ne "Ha" |
        #Where-Object Exposure -eq 180 |
        #Where-object ObsDateMinus12hr -eq ([DateTime]"2021-05-05")
        Where-Object {-not $_.IsIntegratedFile()} #|
        #select-object -First 30
    #$rawSubs|Format-Table Path,*
    $createSuperLum=$false
    $data=Invoke-XisfPostCalibrationMonochromeImageWorkflow `
        -RawSubs $rawSubs `
        -CalibrationPath "F:\PixInsightLT\Calibrated" `
        -CorrectedOutputPath "S:\PixInsight\Corrected" `
        -WeightedOutputPath "S:\PixInsight\Weighted" `
        <#-DarkLibraryPath "E:\Astrophotography\DarkLibrary\QHY268M"#> `
        -DarkLibraryPath "E:\Astrophotography\DarkLibrary\QHYCCD-Cameras-Capture (ASCOM)" `
        -AlignedOutputPath "S:\PixInsight\Aligned" `
        -BackupCalibrationPaths @("M:\PixInsightLT\Calibrated","S:\PixInsightLT\Calibrated") `
        -PixInsightSlot 201 `
        -RerunCosmeticCorrection:$false `
        -SkipCosmeticCorrection:$false `
        -RerunWeighting:$false `
        -SkipWeighting:$false `
        -RerunAlignment:$false `
        -IntegratedImageOutputDirectory $target `
        -AlignmentReference $alignmentReference `
        -GenerateDrizzleData `
        -ApprovalExpression "Median<42 && FWHM<1.6 && Stars > 8000" `
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