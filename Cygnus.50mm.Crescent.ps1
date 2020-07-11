import-module $PSScriptRoot/PsXisfReader.psd1 -Force
$ErrorActionPreference="STOP"

$target="E:\Astrophotography\50mm\PatchworkCygnus_1_0"
$CalibrationPath = "E:\PixInsightLT\Calibrated"
$CorrectedOutputPath = "S:\PixInsight\Corrected"
$WeightedOutputPath = "S:\PixInsight\Weighted"
$AlignedOutputPath = "S:\PixInsight\Aligned"
$BackupCalibrationPaths = @("T:\PixInsightLT\Calibrated","N:\PixInsightLT\Calibrated","S:\PixInsightLT\Calibrated")
Clear-Host

$data =
    Get-XisfLightFrames -Path $target -Recurse |
    Where-Object {-not $_.HasTokensInPath(@("reject","process","testing","clouds","draft"))} |
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
            Write-Host "Correcting $($group.Count) Images"
            $group|group-object{
                ($_.Calibrated | Get-XisfFitsStats | Get-XisfCalibrationState -CalibratedPath $CalibratedOutput).MasterDark
            } | foreach-object {
                $masterDarkFileName = $_.Name
                $masterDark = $DarkLibraryFiles | where-object {$_.Path.Name -eq $masterDarkFileName} | Select-Object -First 1
                if(-not $masterDark){
                    $masterDark = get-childitem "D:\Backups\Camera\2019\Dark Library" *.xisf | Get-XisfFitsStats | where-object {$_.Path.Name -eq $masterDarkFileName} | Select-Object -First 1
                }
                $images = $_.Group
                Invoke-PiCosmeticCorrection `
                    -Images ($images.Calibrated) `
                    -HotDarkLevel 0.4 `
                    -MasterDark ($masterDark.Path) `
                    -OutputPath $CorrectedOutputPath `
                    -PixInsightSlot 200
            }
            $group |
                Get-XisfCosmeticCorrectionState `
                    -CosmeticCorrectionPath $CorrectedOutputPath
        }
    } |
    Get-XisfSubframeSelectorState -SubframeSelectorPath $WeightedOutputPath |
    Group-Object {""} |
    ForEach-Object {
        $group=$_.Group
        if($group | Where-Object {$_.IsWeighted()}) {
            $group
        }
        else{
            Write-Host "Weighting $($group.Count) Images"

            Start-PiSubframeSelectorWeighting `
                -PixInsightSlot 200 `
                -OutputPath $WeightedOutputPath `
                -Images ($group.Corrected) `
                -ApprovalExpression "Median<203 && FWHM<1.2" `
                -WeightingExpression "(15*(1-(FWHM-FWHMMin)/(FWHMMax-FWHMMin))
                +  5*(1-(Eccentricity-EccentricityMin)/(EccentricityMax-EccentricityMin))
                + 15*(SNRWeight-SNRWeightMin)/(SNRWeightMax-SNRWeightMin)
                + 30*(1-(Median-MedianMin)/(MedianMax-MedianMin))
                + 20*(Stars-StarsMin)/(StarsMax-StarsMin))
                + 20"
    
            $group |
                Get-XisfSubframeSelectorState `
                    -SubframeSelectorPath $WeightedOutputPath
        }
    } |
    Get-XisfAlignedState -AlignedPath $AlignedOutputPath |
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
                $reference =  $approved |
                    Sort-Object SSWeight -Descending |
                    Select-Object -First 1
                Write-Host "Aligning $($group.Count) Images"
                Invoke-PiStarAlignment `
                    -PixInsightSlot 200 `
                    -Images ($approved.Path) `
                    -ReferencePath ($reference.Path) `
                    -OutputPath $AlignedOutputPath
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
            $approved | group-object Exposure | foreach-object {
                $exposure=$_.Group[0].Exposure;
                $outputFileName+=".$($_.Group.Count)x$($exposure)s"
            }
            $outputFileName+=".xisf"
            write-host $outputFileName
            $toStack = $approved | sort-object SSWeight -Descending
            $outputFile = Join-Path $target $outputFileName
            if(-not (test-path $outputFile)) {
                Invoke-PiLightIntegration `
                    -Images ($toStack|foreach-object {$_.Path}) `
                    -OutputFile $outputFile `
                    -KeepOpen `
                    -GenerateDrizzleData `
                    -PixInsightSlot 200
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
