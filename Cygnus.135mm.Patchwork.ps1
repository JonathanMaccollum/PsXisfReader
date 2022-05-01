import-module $PSScriptRoot/PsXisfReader.psd1 -Force
$ErrorActionPreference="STOP"

$target="E:\Astrophotography\135mm\Auriga Widefield Panel 2"
$CalibrationPath = "F:\PixInsightLT\Calibrated"
$CorrectedOutputPath = "S:\PixInsight\Corrected"
$WeightedOutputPath = "S:\PixInsight\Weighted"
$AlignedOutputPath = "S:\PixInsight\Aligned"
$BackupCalibrationPaths = @("T:\PixInsightLT\Calibrated","N:\PixInsightLT\Calibrated","S:\PixInsightLT\Calibrated")
$RerunWeighting=$true
$RerunAlignment=$true
$CreateSuperLuminance=$true
$alignmentReference=$null
#$alignmentReference="PatchworkCygnus_0_3.BHS_Ha.11x240s.BHS_Ha.6x300s.BHS_Ha.5x600s.xisf"
Clear-Host
$data =
    Get-XisfLightFrames -Path $target -Recurse |
    Where-Object {-not $_.HasTokensInPath(@("reject","process","testing","clouds","draft","cloudy"))} |
    Where-Object {-not $_.IsIntegratedFile()} |
    Get-XisfCalibrationState `
        -CalibratedPath $CalibrationPath `
        -AdditionalSearchPaths $BackupCalibrationPaths `
        -Verbose |
    foreach-object {
        $x = $_
        if(-not $x.IsCalibrated()){
            Write-Warning "Uncalibrated frame detected... $($x.Path)"
        }
        else {
            $x
        }
    } |
    Get-XisfCosmeticCorrectionState -CosmeticCorrectionPath $CorrectedOutputPath |
    Group-Object {$_.IsCorrected()} |
    ForEach-Object {
        $group=$_.Group
        if( $group[0].IsCorrected() ){
            $group
        }
        else {
            
            $group|group-object{
                ($_.Calibrated | Get-XisfFitsStats | Get-XisfCalibrationState -CalibratedPath $CalibrationPath).MasterDark
            } | foreach-object {
                $masterDarkFileName = $_.Name
                $masterDark = $DarkLibraryFiles | where-object {$_.Path.Name -eq $masterDarkFileName} | Select-Object -First 1
                if(-not $masterDark){
                    $masterDark = get-childitem "D:\Backups\Camera\2019\Dark Library" *.xisf | Get-XisfFitsStats | where-object {$_.Path.Name -eq $masterDarkFileName} | Select-Object -First 1
                }
                if(-not $masterDark){
                    $masterDark = get-childitem "E:\Astrophotography\DarkLibrary\ZWO ASI183MM Pro" *.xisf | Get-XisfFitsStats | where-object {$_.Path.Name -eq $masterDarkFileName} | Select-Object -First 1                    
                }
                $images = $_.Group
                Write-Host "Correcting $($images.Count) Images"
                if(-not $masterDark) {
                    write-warning "Skipping $($images.Count) files... unable to locate master dark: $masterDarkFileName"
                }
                else{
                    Invoke-PiCosmeticCorrection `
                        -Images ($images.Calibrated) `
                        -HotDarkLevel 0.4 `
                        -MasterDark ($masterDark.Path) `
                        -OutputPath $CorrectedOutputPath `
                        -PixInsightSlot 200
                }
            }
            $group |
                Get-XisfCosmeticCorrectionState `
                    -CosmeticCorrectionPath $CorrectedOutputPath
        }
    } |
    Get-XisfSubframeSelectorState -SubframeSelectorPath $WeightedOutputPath |
    foreach-object {
        $x = $_
        if($RerunWeighting -and $x.IsWeighted()) {
            Remove-Item $x.Weighted -Verbose
            $x.Weighted.Refresh()
        }
        $x
    } |
    Group-Object { $_.IsWeighted() } |
    ForEach-Object {
        $group=$_.Group
        if($group | Where-Object {$_.IsWeighted()}) {
            $group
        }
        else{
            $group | group-object {$_.Stats.Filter } | foreach-object {
                $byFilter = $_.Group
                $filter=$byFilter[0].Stats.Filter
                Write-Host "Weighting $($byFilter.Count) Images for filter $filter"
                Start-PiSubframeSelectorWeighting `
                    -PixInsightSlot 200 `
                    -OutputPath $WeightedOutputPath `
                    -Images ($byFilter.Corrected) `
                    -ApprovalExpression "Median<100 && FWHM<4.5" `
                    -WeightingExpression "(15*(1-(FWHM-FWHMMin)/(FWHMMax-FWHMMin))
                    +  5*(1-(Eccentricity-EccentricityMin)/(EccentricityMax-EccentricityMin))
                    + 05*(SNRWeight-SNRWeightMin)/(SNRWeightMax-SNRWeightMin)
                    + 20*(1-(Median-MedianMin)/(MedianMax-MedianMin))
                    + 40*(Stars-StarsMin)/(StarsMax-StarsMin))
                    + 20"
        
            }
            $group |
                Get-XisfSubframeSelectorState `
                    -SubframeSelectorPath $WeightedOutputPath
        }
    } |
    Get-XisfAlignedState -AlignedPath $AlignedOutputPath |
    foreach-object {
        $x = $_
        if(($RerunAlignment -or $RerunWeighting) -and $x.IsAligned()) {
            Remove-Item $x.Aligned -Verbose
            $x.Aligned.Refresh()
        }
        $x
    } |
    Group-Object {$_.IsAligned()} |
    ForEach-Object {
        $group=$_.Group
        if( $group[0].IsAligned() ){
            $group
        }
        else {
            $approved = $group |
                Where-object {$_.IsWeighted()} |
                foreach-object{$_.Weighted} |
                Get-XisfFitsStats
            if(-not $approved){
                $group
            }
            else {
                $reference = join-path $target $alignmentReference
                if(-not ($alignmentReference -and (test-path $reference))){
                    $reference =  ($approved |
                        Sort-Object SSWeight -Descending |
                        Select-Object -First 1).Path
                }

                Write-Host "Aligning $($approved.Count) Images"
                Invoke-PiStarAlignment `
                    -PixInsightSlot 200 `
                    -Images ($approved.Path) `
                    -ReferencePath ($reference) `
                    -OutputPath $AlignedOutputPath `
                    -Interpolation Lanczos4 `
                    -ClampingThreshold 0.2
                $group |
                    Get-XisfAlignedState `
                        -AlignedPath $AlignedOutputPath
                
            }
        }
    } |
    Group-Object {$_.IsAligned()} |
    ForEach-Object {
        $group=$_.Group
        if( $group[0].IsAligned() ){
            $approved = $group |
                foreach-object{$_.Aligned} |
                Get-XisfFitsStats
            $reference =  $approved |
                Sort-Object SSWeight -Descending |
                Select-Object -First 1
            
            $outputFileName = $reference.Object
            $approved | group-object Exposure,Filter | foreach-object {
                $exposure=$_.Group[0].Exposure;
                $filter=$_.Group[0].Filter;
                $outputFileName+=".$filter.$($_.Group.Count)x$($exposure)s"
            }
            $outputFileName+=".Adaptive.ESD.xisf"
            $outputFile = Join-Path $target $outputFileName


            if((-not (test-path $outputFile)) -and $CreateSuperLuminance) {
                write-host ("Creating super luminance "+ $outputFileName)
                $toStack = $approved | sort-object SSWeight -Descending
                Invoke-PiLightIntegration `
                    -Images ($toStack|foreach-object {$_.Path}) `
                    -OutputFile $outputFile `
                    -KeepOpen `
                    -GenerateDrizzleData `
                    -Rejection "Rejection_ESD" `
                    -Normalization "AdaptiveNormalization" `
                    -RejectionNormalization "AdaptiveRejectionNormalization" `
                    -LinearFitLow 5 `
                    -LinearFitHigh 4 `
                    -PixInsightSlot 200
            }

            $group|group-object {$_.Stats.Filter}|foreach-object {
                $byFilter=$_.Group |
                    foreach-object{$_.Aligned} |
                    Get-XisfFitsStats
                $filter=$byFilter[0].Filter
                    
                $outputFileName = $reference.Object
                $byFilter | group-object Exposure | foreach-object {
                    $exposure=$_.Group[0].Exposure;
                    $outputFileName+=".$filter.$($_.Group.Count)x$($exposure)s"
                }
                $outputFileName+=".Adaptive.ESD.xisf"
                $outputFile = Join-Path $target $outputFileName
                if(-not (test-path $outputFile)) {
                    write-host ("Integrating  "+ $outputFileName)
                    $toStack = $byFilter | sort-object SSWeight -Descending
                    try {
                        Invoke-PiLightIntegration `
                        -Images ($toStack|foreach-object {$_.Path}) `
                        -OutputFile $outputFile `
                        -KeepOpen `
                        -Rejection "Rejection_ESD" `
                        -Normalization "AdaptiveNormalization" `
                        -RejectionNormalization "AdaptiveRejectionNormalization" `
                        -LinearFitLow 5 `
                        -LinearFitHigh 4 `
                        -PixInsightSlot 200
                    }
                    catch {
                        write-warning $_.ToString()
                    }
                }
            }

            $group
        }
        else{
            $group
        }
    }

$mostRecent=
    $data.Stats|
    Sort-Object LocalDate -desc |
    Select-Object -first 1

$data | Export-Clixml -Path (Join-Path $target "Stats.$($mostRecent.LocalDate.ToString('yyyyMMdd HHmmss')).clixml")


