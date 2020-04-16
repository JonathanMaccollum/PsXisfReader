import-module $PSScriptRoot/pixinsightpreprocessing.psm1 -Force
import-module $PSScriptRoot/PsXisfReader.psm1

$target="E:\Astrophotography\1000mm\NGC3718"
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

$calibrated = $data |
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
            -WeightingExpression "(5*(1-(FWHM-FWHMMin)/(FWHMMax-FWHMMin))
            + 15*(1-(Eccentricity-EccentricityMin)/(EccentricityMax-EccentricityMin))
            + 15*(SNRWeight-SNRWeightMin)/(SNRWeightMax-SNRWeightMin)
            + 20*(Stars-StarsMin)/(StarsMax-StarsMin))
            + 30" `
            -ApprovalExpression "FWHM<4.5 && Eccentricity<0.65"
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
        $subframeData|FT
    }

$stats = get-childitem $target *.data.csv |
        import-csv |
        foreach-object {
            $x=$_
            $y = Get-XisfFitsStats -Path ($x.File.Replace("/","\"))
            $r=$w=$null
            if($x.Approved -eq "true") {
                $w = Get-Item (join-path $WeightedOutputPath ($y.Path.Name.TrimEnd(".xisf")+"_a.xisf"))
                $r = join-path $AlignedOutputPath ($y.Path.Name.TrimEnd(".xisf")+"_a_r.xisf")
            }
            Add-Member -InputObject $y -Name "Approved" -MemberType NoteProperty -Value ([bool]::Parse($x.Approved))
            Add-Member -InputObject $y -Name "Weight" -MemberType NoteProperty -Value ([decimal]::Parse($x.Weight))
            Add-Member -InputObject $y -Name "Weighted" -MemberType NoteProperty -Value ($w)
            Add-Member -InputObject $y -Name "Aligned" -MemberType NoteProperty -Value ($r)
            if($y) {
                $y
            }
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
write-host "Approved Total:"
[TimeSpan]::FromSeconds((
        $summary|where-Object Approved -eq $true|foreach-object {$_.ExposureTime.TotalSeconds} |Measure-Object  -Sum).Sum).ToString()
Write-Host
write-host "Approved L:"
[TimeSpan]::FromSeconds((
        $summary|where-Object Approved -eq $true | where-object Filter -eq "'L'"|foreach-object {$_.ExposureTime.TotalSeconds} |Measure-Object  -Sum).Sum).ToString()
write-host "Approved R:"
[TimeSpan]::FromSeconds((
        $summary|where-Object Approved -eq $true | where-object Filter -eq "'R'"|foreach-object {$_.ExposureTime.TotalSeconds} |Measure-Object  -Sum).Sum).ToString()
write-host "Approved G:"
[TimeSpan]::FromSeconds((
        $summary|where-Object Approved -eq $true | where-object Filter -eq "'G'"|foreach-object {$_.ExposureTime.TotalSeconds} |Measure-Object  -Sum).Sum).ToString()
write-host "Approved B:"
[TimeSpan]::FromSeconds((
        $summary|where-Object Approved -eq $true | where-object Filter -eq "'B'"|foreach-object {$_.ExposureTime.TotalSeconds} |Measure-Object  -Sum).Sum).ToString()
        

$referenceFrame = $stats | where-object Approved -eq "true" | where-object Filter -eq "'L'" | sort-object "Weight" -Descending | select-object -first 1
$stats | where-object Approved -eq "true" | group-object Filter | foreach-object {
    $filter = $_.Group[0].Filter

    $toAlign = $_.Group | where-object {$_.Aligned -and -not (Test-Path $_.Aligned) -and ($_.Weighted) } | foreach-object {$_.Weighted}
    if($toAlign){
        write-host "Aligning $($toAlign.Count) for filter $filter"
        Invoke-PiStarAlignment `
            -PixInsightSlot 200 `
            -Images $toAlign `
            -ReferencePath ($referenceFrame.Weighted) `
            -OutputPath $AlignedOutputPath
    }
}
$aligned = $stats | where-object {$_.Aligned -and (Test-Path $_.Aligned) }
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
        $ref = $aligned |
            where-object Approved -eq "true" |
            where-object Filter -eq $filter |
            sort-object "Weight" -Descending | 
            select-object -first 1
        write-host ($ref.Aligned)
        $toStack = $_.Images | sort-object {
            $x = $_
            ($x.Aligned) -eq ($ref.Aligned)
        }
        $outputFile = Join-Path $target $outputFileName
        if(-not (test-path $outputFile) -and $toStack.Count -gt 3) {
            write-host "Integrating $outputFile"
            Invoke-PiLightIntegration `
                -Images ($toStack|foreach-object {$_.Aligned}) `
                -OutputFile $outputFile `
                -PixInsightSlot 200 `
                -Rejection "LinearFit" `
                -LinearFitHigh 7 `
                -LinearFitLow 8;
        }
    }
