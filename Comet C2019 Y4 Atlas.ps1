import-module $PSScriptRoot/pixinsightpreprocessing.psm1 -Force
import-module $PSScriptRoot/PsXisfReader.psm1 -Force

$target="E:\Astrophotography\1000mm\Comet C2019 Y4 Atlas"
$CalibrationPath = "E:\PixInsightLT\Calibrated"
$WeightedOutputPath = "S:\PixInsight\Weighted"
$AlignedOutputPath = "S:\PixInsight\Aligned"
$CometAlignedOutputPath = "S:\PixInsight\CometAligned"
$BackupCalibrationPaths = @("T:\PixInsightLT\Calibrated")



$data = 
    Get-XisfLightFrames -Path $target -Recurse |
    where-object {-not $_.Path.FullName.ToLower().Contains("reject")} |
    where-object {-not $_.Path.FullName.ToLower().Contains("testing")} |
    where-object {-not $_.Path.FullName.ToLower().Contains("clouds")} |
    foreach-object {
        $x=$_
        $y = Get-CalibrationFile -Path ($x.Path) `
            -CalibratedPath $CalibrationPath `
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

$data|group-object ObsDateMinus12hr,Filter|foreach-object {
    $images = $_.Group
    $filter=$images[0].Filter.Trim()
    $obsDate=$images[0].ObsDateMinus12hr
    $toWeigh = $_.Group.Calibrated

    $resultCsv="$target\Subframe.$($obsDate.ToString('yyyyMMdd')).$filter.csv"
    $resultData="$target\Subframe.$($obsDate.ToString('yyyyMMdd')).$filter.Data.csv"

    if(-not (test-path $resultCsv)) {
        set-clipboard -Value $resultCsv
        write-host "Starting Subframe Selector... Save the resulting output CSV file to $resultCsv before exiting."
        Start-PiSubframeSelectorWeighting `
        -PixInsightSlot 200 `
        -OutputPath $WeightedOutputPath `
        -Images $toWeigh `
        -WeightingExpression "(10*(1-(FWHM-FWHMMin)/(FWHMMax-FWHMMin))
        + 10*(1-(Eccentricity-EccentricityMin)/(EccentricityMax-EccentricityMin))
        + 10*(SNRWeight-SNRWeightMin)/(SNRWeightMax-SNRWeightMin)
        + 20*(1-(Median-MedianMin)/(MedianMax-MedianMin))
        + 30*(Stars-StarsMin)/(StarsMax-StarsMin))
        + 30" `
        -ApprovalExpression "Stars > 25 && NoiseRatio > 0.94 && Median < 75"
    }    
    $subframeResults = Get-Content -Path $resultCsv
    $subframeResults | Select-Object -Skip 29|Out-File -Path $resultData -Force
    #$subframeResults | Select-Object -First 29
    $subframeData    = Import-Csv -Path "$resultData"|sort-object {[double]$_.Weight} -Descending
    $subframeData|Format-Table
}


$data|group-object ObsDateMinus12hr,Exposure|foreach-object {
    $images = $_.Group
    $exposure=$images[0].Exposure.Trim()
    $obsDate=$images[0].ObsDateMinus12hr
    $x=$images|foreach-object{@{UtcDate=([DateTime]$_.ObsDate).ToUniversalTime()}}|measure-object UtcDate -Minimum -Maximum
    new-object psobject -Property @{
        ObsDate=$obsDate.ToString('yyyyMMdd')
        Exposures="$($images.Count.ToSTring('000'))x$($exposure)s"
        StartingUTC=($x.Minimum.ToString('HH:mm:ss'))
        EndingUTC=($x.Maximum.ToString('HH:mm:ss'))
    }
} |Sort-Object ObsDate | Format-Table ObsDate,Exposures,StartingUTC,EndingUTC


$data|group-object ObsDateMinus12hr|foreach-object {
    $images = $_.Group
    $obsDate=$images[0].ObsDateMinus12hr

    $all = Get-ChildItem -Path $target "Subframe.$($obsDate.ToString('yyyyMMdd')).*.Data.csv" | 
        Import-CSV |
        foreach-object {
            $x=$_
            $y = Get-XisfFitsStats -Path ($x.File.Replace("/","\"))
            $r=$w=$null
            if($x.Approved -eq "true") {
                $w = Get-Item (join-path $WeightedOutputPath ($y.Path.Name.TrimEnd(".xisf")+"_a.xisf"))
                $r = join-path $AlignedOutputPath ($y.Path.Name.TrimEnd(".xisf")+"_a_r.xisf")
                $ca = join-path $CometAlignedOutputPath ($y.Path.Name.TrimEnd(".xisf")+"_a_r_a.xisf")
            }
            Add-Member -InputObject $y -Name "Approved" -MemberType NoteProperty -Value ([bool]::Parse($x.Approved))
            Add-Member -InputObject $y -Name "Weight" -MemberType NoteProperty -Value ([decimal]::Parse($x.Weight))
            Add-Member -InputObject $y -Name "Weighted" -MemberType NoteProperty -Value ($w)
            Add-Member -InputObject $y -Name "Aligned" -MemberType NoteProperty -Value ($r)
            Add-Member -InputObject $y -Name "CometAligned" -MemberType NoteProperty -Value ($ca)
            if($y) {
                $y
            }
        }
    $referenceFrame = $all|sort-object {[double]$_.Weight} -Descending | select-object -First 1
    <#
    $toAlign = $all | where-object {$_.Aligned -and -not (Test-Path $_.Aligned) -and ($_.Weighted) } | foreach-object {$_.Weighted}
    if($toAlign){
        write-host "Aligning $($toAlign.Count) for night $obsDate"
        $x = @{
            PixInsightSlot = 200
            Images = $toAlign
            ReferencePath = ($referenceFrame.Weighted)
            OutputPath = $AlignedOutputPath
        }    
        Invoke-PiStarAlignment @x
    }
    #>
    $aligned = $all | where-object {$_.Aligned -and (Test-Path $_.Aligned)}
    if($aligned.Count-lt 3){
        return;
    }

    write-host "$($obsDate.ToString('yyyyMMdd'))"
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
            $outputFileName+=".$($obsDate.ToString('yyyyMMdd')).$($filter.Trim("'"))"
            $_.Images | group-object Exposure | foreach-object {
                $exposure=$_.Group[0].Exposure;
                $outputFileName+=".$($_.Group.Count)x$($exposure)s"
            }
            $outputFileName+=".xisf"
            write-host $outputFileName
            $ref = $all |
                where-object Approved -eq "true" |
                where-object Filter -eq $filter |
                sort-object "Weight" -Descending | 
                select-object -first 1
            write-host ($ref.Aligned)
            $toStack = $_.Images | sort-object {
                $x = $_
                ($x.Aligned) -ne ($ref.Aligned)
            }
            $outputFile = Join-Path $target $outputFileName
            if(-not (test-path $outputFile) -and $toStack.Count -gt 3) {
                write-host "Integrating $outputFile"
                Invoke-PiLightIntegration `
                    -Images ($toStack|foreach-object {$_.Aligned}) `
                    -OutputFile $outputFile `
                    -PixInsightSlot 200 `
                    -Rejection "PercentileClip" `
                    -LinearFitHigh 7 `
                    -LinearFitLow 8;
            }
        }

<#

    #Comet Alignment
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
            $ref = $all |
                where-object Approved -eq "true" |
                where-object Filter -eq $filter |
                sort-object "Weight" -Descending | 
                select-object -first 1
            $toAlign = $_.Images | sort-object {
                $x = $_
                ($x.Aligned) -ne ($ref.Aligned)
            } | foreach-object {$_.Aligned}
            if(-not $toAlign.Contains($referenceFrame.Aligned)){
                $toAlign = @($referenceFrame.Aligned)+$toAlign
            }
            Start-PiCometAlignment `
                -Images $toAlign `
                -OutputPath $CometAlignedOutputPath `
                -Verbose `
                -PixInsightSlot 200
        }
        #>

    $cometAligned = $all | where-object {$_.CometAligned -and (Test-Path $_.CometAligned)}
    if($cometAligned.Count-lt 3){
        return;
    }


    $cometAligned |
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
            $outputFileName+=".$($obsDate.ToString('yyyyMMdd')).$($filter.Trim("'")).CometAligned.PC"
            $_.Images | group-object Exposure | foreach-object {
                $exposure=$_.Group[0].Exposure;
                $outputFileName+=".$($_.Group.Count)x$($exposure)s"
            }
            $outputFileName+=".xisf"
            write-host $outputFileName
            $ref = $all |
                where-object Approved -eq "true" |
                where-object Filter -eq $filter |
                sort-object "Weight" -Descending | 
                select-object -first 1
            write-host ($ref.CometAligned)
            $toStack = $_.Images | sort-object {
                $x = $_
                ($x.CometAligned) -ne ($ref.CometAligned)
            }
            $outputFile = Join-Path $target $outputFileName
            if(-not (test-path $outputFile) -and $toStack.Count -gt 3) {
                write-host "Integrating $outputFile"
                Invoke-PiLightIntegration `
                    -Images ($toStack|foreach-object {$_.CometAligned}) `
                    -OutputFile $outputFile `
                    -PixInsightSlot 200 `
                    -Rejection "PercentileClip"
            }
        }
    
}





