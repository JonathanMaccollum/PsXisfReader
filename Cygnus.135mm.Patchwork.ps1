import-module $PSScriptRoot/PsXisfReader.psd1 -Force
$ErrorActionPreference="STOP"

$target="E:\Astrophotography\135mm\PatchworkCygnus135_0_0"
$CalibrationPath = "E:\PixInsightLT\Calibrated"
$CorrectedOutputPath = "S:\PixInsight\Corrected"
$WeightedOutputPath = "S:\PixInsight\Weighted"
$DebayeredOutputPath = "S:\PixInsight\Debayered"
$AlignedOutputPath = "S:\PixInsight\Aligned"
$BackupCalibrationPaths = @("T:\PixInsightLT\Calibrated","N:\PixInsightLT\Calibrated","S:\PixInsightLT\Calibrated")
Clear-Host

$data =
    Get-XisfLightFrames -Path $target -Recurse |
    Where-Object {-not $_.HasTokensInPath(@("reject","process","testing","clouds"))} |
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
                ($group.Calibrated | Get-XisfFitsStats | Get-XisfCalibrationState -CalibratedPath $CalibratedOutput).MasterDark
            } | foreach-object {
                $masterDarkFileName = $_.Name
                $masterDark = $DarkLibrary | where-object {$_.Path.Name -eq $masterDarkFileName} | Select-Object -First 1
                $images = $_.Group
                Invoke-PiCosmeticCorrection `
                    -Images ($images.Calibrated) `
                    -CFAImages `
                    -HotDarkLevel 0.5 `
                    -MasterDark ($masterDark.Path) `
                    -OutputPath $CorrectedOutputPath `
                    -PixInsightSlot 200
            }
            $group |
                Get-XisfCosmeticCorrectionState `
                    -CosmeticCorrectionPath $CorrectedOutputPath
        }
    } |
    Get-XisfDebayerState -DebayerPath $DebayeredOutputPath |
    Group-Object {$_.IsDebayered()} |
    ForEach-Object {
        $group=$_.Group
        if( $group[0].IsDebayered() ){
            $group
        }
        else {
            Write-Host "Debayering $($group.Count) Images"
            Invoke-PiDebayer `
                -PixInsightSlot 200 `
                -Images ($group.Corrected) `
                -OutputPath $DebayeredOutputPath

            $group |
                Get-XisfDebayerState `
                    -DebayerPath $DebayeredOutputPath
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
                -Images ($group.Debayered) `
                -ApprovalExpression "Median<203 && FWHM<1.2" `
                -WeightingExpression "(15*(1-(FWHM-FWHMMin)/(FWHMMax-FWHMMin))
                + 05*(1-(Eccentricity-EccentricityMin)/(EccentricityMax-EccentricityMin))
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
                    -PixInsightSlot 200
            }


        }
        else{
            $group
        }
    }
    










exit
$data = 
    Get-XisfLightFrames -Path $target -Recurse |
    where-object {-not $_.Path.FullName.ToLower().Contains("reject")} |
    where-object {-not $_.Path.FullName.ToLower().Contains("process")} |
    where-object {-not $_.Path.FullName.ToLower().Contains("testing")} |
    where-object {-not $_.Path.FullName.ToLower().Contains("clouds")} |
    Where-Object {-not $_.History} |
    foreach-object {
        $x=$_
        $y = Get-CalibrationFile -Path ($x.Path) `
            -CalibratedPath (Join-Path $CalibrationPath ($x.Object)) `
            -AdditionalSearchPaths $BackupCalibrationPaths
        Add-Member -InputObject $x -Name "Calibrated" -MemberType NoteProperty -Value $y -Force
        $x
    }
$uncalibrated = $data | where-object {-not $_.Calibrated}
if($uncalibrated){
    $uncalibrated|foreach-object {
        Write-Host ($_.Path.Name)
    }
    Write-Warning "$($uncalibrated.Length) uncalibrated frames detected"
    if(-not (Read-Host "Continue?").ToLower().StartsWith("y")){
        break;
    }
}
$data|group-object Filter,Exposure|foreach-object {
    $filter=$_.Group[0].Filter
    $exposure=$_.Group[0].Exposure
        new-object psobject -Property @{
            Filter=$filter
            Exposure=$exposure
            Exposures=$_.Group.Count
            ExposureTime=([TimeSpan]::FromSeconds($_.Group.Count*$exposure))
            Images=$_.Group
        }
} | Sort-Object Filter|Format-Table Filter,Exposures,Exposure,ExposureTime
exit
$data|group-object Instrument,Gain,Offset,Exposure,SetTemp|foreach-object {
    $masterDark = ($DarkLibrary|group-object Instrument,Gain,Offset,Exposure,SetTemp|? Name -eq $_.Name).Group | select-object -First 1
    $images=$_.Group
    Invoke-PiCosmeticCorrection `
        -Images ($images.Calibrated) `
        -CFAImages `
        -HotDarkLevel 0.5 `
        -MasterDark ($masterDark.Path) `
        -OutputPath $CorrectedOutputPath `
        -PixInsightSlot 200 `
        -KeepOpen
}
exit
$data |
    where-object { $uncalibrated -notcontains $_ } |
    group-object Filter |
    foreach-object {
        $images = $_.Group
        $filter=$images[0].Filter.Trim()
        $toWeigh = $_.Group.Calibrated

        $resultCsv="$target\Subframe.$filter.csv"
        $resultData="$target\Subframe.$filter.Data.csv"
        if(-not (test-path $resultCsv)) {
            set-clipboard -Value $resultCsv
            write-host "Starting Subframe Selector... Save the resulting output CSV file to $resultCsv before exiting. (Path copied to clipboard.)"

            Start-PiSubframeSelectorWeighting `
            -PixInsightSlot 200 `
            -OutputPath $WeightedOutputPath `
            -Images $toWeigh `
            -ApprovalExpression "Median<203 && FWHM<0.6"
            -WeightingExpression "(15*(1-(FWHM-FWHMMin)/(FWHMMax-FWHMMin))
            + 05*(1-(Eccentricity-EccentricityMin)/(EccentricityMax-EccentricityMin))
            + 15*(SNRWeight-SNRWeightMin)/(SNRWeightMax-SNRWeightMin)
            + 30*(1-(Median-MedianMin)/(MedianMax-MedianMin))
            + 20*(Stars-StarsMin)/(StarsMax-StarsMin))
            + 20"
        }
        
        $subframeResults = Get-Content -Path $resultCsv
        $indexOfCSV=0
        $csvFound=$false
        $subframeResults | foreach-object {
            if($_.StartsWith("Index,Approved,Locked")){                
                $csvFound=$true;
            }
            if(-not $csvFound){
                $indexOfCSV+=1;
            }            
        }
        $subframeResults | Select-Object -Skip $indexOfCSV|Out-File -Path $resultData -Force
        $subframeResults | Select-Object -First $indexOfCSV
        $subframeData    = Import-Csv -Path "$resultData"|sort-object {[double]$_.Weight} -Descending
        $subframeData|Format-Table
    }
exit
$stats = get-childitem $target *.data.csv |
        import-csv |
        foreach-object {
            $x=$_
            $y = Get-XisfFitsStats -Path ($x.File.Replace("/","\"))
            $w=$null
            if($x.Approved -eq "true") {
                $w = Get-Item (join-path $WeightedOutputPath ($y.Path.Name.TrimEnd(".xisf")+"_a.xisf"))
            }
            Add-Member -InputObject $y -Name "Approved" -MemberType NoteProperty -Value ([bool]::Parse($x.Approved))
            Add-Member -InputObject $y -Name "Weight" -MemberType NoteProperty -Value ([decimal]::Parse($x.Weight))
            Add-Member -InputObject $y -Name "Weighted" -MemberType NoteProperty -Value ($w)
            $y
        }
$summary = $stats |
    group-object Approved,Filter,Exposure |
    foreach-object {
        $approved=$_.Values[0]
        $filter=$_.Values[1]
        $exposure=$_.Values[2]
        $topWeight = $_.Group | sort-object Weight -Descending | select-object -first 1
        new-object psobject -Property @{
            Approved=$approved
            Filter=$filter
            Exposure=$exposure
            Exposures=$_.Group.Count
            ExposureTime=([TimeSpan]::FromSeconds($_.Group.Count*$exposure))
            Images=$_.Group
            TopWeight = $topWeight.Path
        } } 
$summary |
        Sort-Object Approved,Filter |
        Format-Table Approved,Filter,Exposures,Exposure,ExposureTime,TopWeight

write-host "Rejected:"
[TimeSpan]::FromSeconds((
    $summary|where-Object Approved -eq $false | foreach-object {$_.ExposureTime.TotalSeconds} |Measure-Object  -Sum).Sum).ToString()
write-host "Approved:"
[TimeSpan]::FromSeconds((
        $summary|where-Object Approved -eq $true|foreach-object {$_.ExposureTime.TotalSeconds} |Measure-Object  -Sum).Sum).ToString()

#$referenceFrame = $stats | where-object Approved -eq "true" | sort-object "Weight" -Descending | select-object -first 1
$referenceFrame = $stats | where-object Approved -eq "true" | where-object Filter -eq "'L3'" | sort-object "Weight" -Descending | select-object -first 1

<#
$stats | where-object Approved -eq "true" | group-object Filter | foreach-object {
$filter = $_.Group[0].Filter

write-host "Aligning $($_.Group.Count) for filter $filter"
$images = $_.Group | foreach-object {$_.Weighted}
Invoke-PiStarAlignment `
    -PixInsightSlot 200 `
    -Images $images `
    -ReferencePath "E:\Astrophotography\135mm\NGC2403\NGC2403.071mc.135mm.180x2min.integrated.xisf" `
    -OutputPath $AlignedOutputPath `
    -KeepOpen
}
#>


$aligned = Get-XisfFile -Path $AlignedOutputPath |
    where-object object -eq "'M109'" 
write-host "Aligned Stats"
$aligned |
    group-object Filter,Exposure | 
    foreach-object {
        $filter=$_.Values[0]
        $exposure=$_.Values[1]
        new-object psobject -Property @{
            Filter=$filter
            Exposure=$exposure
            Exposures=$_.Group.Count
            ExposureTime=([TimeSpan]::FromSeconds($_.Group.Count*$exposure))
            Images=$_.Group
        } } |
    Sort-Object Filter |
    Format-Table Filter,Exposures,Exposure,ExposureTime,TopWeight

$aligned |
    group-object Filter | 
    foreach-object {
        $filter=$_.Values[0]
        new-object psobject -Property @{
            Filter=$filter
            Images=$_.Group
        } } |
    ForEach-Object {
        $filter = $_.Filter
        $outputFileName = $_.Images[0].Object.Trim("'")
        $outputFileName+=".$($filter.Trim("'"))"            
        $_.Images | group-object Exposure | foreach-object {
            $exposure=$_.Group[0].Exposure;
            $outputFileName+=".$($_.Group.Count)x$($exposure)s"
        }
        $outputFileName+=".xisf"
        write-host $outputFileName
        $ref = $stats |
            where-object Approved -eq "true" |
            where-object Filter -eq $filter |
            sort-object "Weight" -Descending | 
            select-object -first 1
        write-host ($ref.Path.Name.Replace(".xisf","_a_r.xisf"))
        $toStack = $_.Images | sort-object {
            $x = $_
            ($x.Path.Name) -ne ($ref.Path.Name.Replace(".xisf","_a_r.xisf"))
        }
        $outputFile = Join-Path $target $outputFileName
        if(-not (test-path $outputFile)) {
            Invoke-PiLightIntegration `
                -Images ($toStack|foreach-object {$_.Path}) `
                -OutputFile $outputFile `
                -KeepOpen `
                -PixInsightSlot 200
        }
    }

