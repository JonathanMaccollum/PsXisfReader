import-module $PSScriptRoot/pixinsightpreprocessing.psm1 -Force
import-module $PSScriptRoot/PsXisfReader.psm1

$target="E:\Astrophotography\135mm\NGC2403"
$CalibrationPath = "E:\PixInsightLT\Calibrated"
$WeightedOutputPath = "S:\PixInsight\Weighted"
$AlignedOutputPath = "S:\PixInsight\Aligned"

$BackupCalibrationPaths = @("T:\PixInsightLT\Calibrated","N:\PixInsightLT\Calibrated","S:\PixInsightLT\Calibrated")

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
$cometData =    $data | where-object {[DateTime]($_.ObsDate)-gt [DateTime]("2020-04-01") }

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

$cometData|group-object Filter,Exposure|foreach-object {
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

$calibrated = $cometData |
    where-object { $uncalibrated -notcontains $_ } |
    group-object Filter,Exposure |
    foreach-object {
        $images = $_.Group
        $filter=$images[0].Filter.Trim()
        $exposure=$images[0].Exposure.Trim()
        $toWeigh = $_.Group.Calibrated

        $resultCsv="$target\Subframe.$filter.$exposure.csv"
        $resultData="$target\Subframe.$filter.$exposure.Data.csv"
        if(-not (test-path $resultCsv)) {
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
        $subframeResults | Select-Object -Skip 29|Out-File -Path $resultData -Force
        $subframeResults | Select-Object -First 29
        $subframeData    = Import-Csv -Path "$resultData"|sort-object {[double]$_.Weight} -Descending
        $subframeData|FT
    }

$stats = get-childitem $target *60.data.csv |
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

