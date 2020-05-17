import-module $PSScriptRoot/pixinsightpreprocessing.psm1 -force
import-module $PSScriptRoot/PsXisfReader.psm1 -force

$ErrorActionPreference="STOP"
$target="E:\Astrophotography\50mm\Patchwork Cygnus"
$CalibrationPath = "E:\PixInsightLT\Calibrated"
$WeightedOutputPath = "S:\PixInsight\Weighted"
$AlignedOutputPath = "S:\PixInsight\Aligned"

$BackupCalibrationPaths = @("T:\PixInsightLT\Calibrated")

$data = 
    Get-ChildItem -Path "E:\Astrophotography\50mm" -Directory |? {$_.Name.StartsWith("PatchworkCygnus")} |
        Get-ChildItem -File -Filter *.xisf -Recurse |
    Get-XisfFitsStats |
    where-object ImageType -eq "LIGHT" |
    where-object {-not $_.Path.FullName.ToLower().Contains("reject")} |
    where-object {-not $_.Path.FullName.ToLower().Contains("testing")} |
    where-object {-not $_.Path.FullName.ToLower().Contains("clouds")} |
    where-object {-not $_.History} |
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

$data|group-object Object,Filter,Exposure|foreach-object {
    $targetName = $_.Group[0].Object
    $filter=$_.Group[0].Filter
    $exposure=$_.Group[0].Exposure
        new-object psobject -Property @{
            Object=$targetName
            Filter=$filter
            Exposure=$exposure
            Exposures=$_.Group.Count
            ExposureTime=([TimeSpan]::FromSeconds($_.Group.Count*$exposure))
            Images=$_.Group
        }
} | Sort-Object Filter|Format-Table Object,Filter,Exposures,Exposure,ExposureTime

$calibrated = $data |
    where-object { $uncalibrated -notcontains $_ } |
    group-object Object,Filter |
    foreach-object {
        $images = $_.Group
        $targetName=$images[0].Object
        $filter=$images[0].Filter.Trim()
        $toWeigh = $_.Group.Calibrated

        $resultCsv="$target\Subframe.$targetName.$filter.csv"
        $resultData="$target\Subframe.$targetName.$filter.Data.csv"
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
            -ApprovalExpression "Median<40"
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

$stats | where-object Object -eq "Crescent Nebula" | foreach-object {$_.Object="PatchworkCygnus_1_2"}

$summary = $stats |
    group-object Object,Approved,Filter,Exposure |
    foreach-object {
        $targetName=$_.Values[0]
        $approved=$_.Values[1]
        $filter=$_.Values[2]
        $exposure=$_.Values[3]
        $topWeight = $_.Group | sort-object Weight -Descending | select-object -first 1
        new-object psobject -Property @{
            Object=$targetName
            Approved=$approved
            Filter=$filter
            Exposure=$exposure
            Exposures=$_.Group.Count
            ExposureTime=([TimeSpan]::FromSeconds($_.Group.Count*$exposure))
            Images=$_.Group
            TopWeight = $topWeight.Path
        } } 
$summary |
        Sort-Object Object,Approved,Filter |
        Format-Table Object,Approved,Filter,Exposures,Exposure,ExposureTime,TopWeight

write-host "Rejected:"
[TimeSpan]::FromSeconds((
    $summary|where-Object Approved -eq $false | foreach-object {$_.ExposureTime.TotalSeconds} |Measure-Object  -Sum).Sum).ToString()
write-host "Approved Total:"
[TimeSpan]::FromSeconds((
        $summary|where-Object Approved -eq $true|foreach-object {$_.ExposureTime.TotalSeconds} |Measure-Object  -Sum).Sum).ToString()

$stats | where-object Approved -eq "true" | group-object Object,Filter | foreach-object {
    $targetName=$_.Group[0].Object
    $filter = $_.Group[0].Filter
    $referenceFrame = $_.Group | where-object Approved -eq "true" | where-object Filter -eq "'Ha'" | sort-object "Weight" -Descending | select-object -first 1

    $toAlign = $_.Group | where-object {$_.Aligned -and -not (Test-Path $_.Aligned) -and ($_.Weighted) } | foreach-object {$_.Weighted}
    if($toAlign){
        write-host "$targetName - Aligning $($toAlign.Count) for filter $filter"
        Invoke-PiStarAlignment `
            -PixInsightSlot 200 `
            -Images $toAlign `
            -ReferencePath ($referenceFrame.Weighted) `
            -OutputPath $AlignedOutputPath

        $failedAlignemnt = $_.Group | where-object {$_.Aligned -and -not (Test-Path $_.Aligned) -and ($_.Weighted) } | foreach-object {$_.Weighted}
        if($failedAlignemnt){
            Write-Warning "$($failedAlignemnt.Count) $filter subs failed to align... increasing detection scales to 6"
            Invoke-PiStarAlignment `
                -PixInsightSlot 200 `
                -Images $failedAlignemnt `
                -ReferencePath ($referenceFrame.Weighted) `
                -OutputPath $AlignedOutputPath `
                -DetectionScales 6
        }
        $failedAlignemnt = $_.Group | where-object {$_.Aligned -and -not (Test-Path $_.Aligned) -and ($_.Weighted) } | foreach-object {$_.Weighted}
        if($failedAlignemnt){
            Write-Warning "$($failedAlignemnt.Count) $filter subs failed to align... increasing detection scales to 7"
            Invoke-PiStarAlignment `
                -PixInsightSlot 200 `
                -Images $failedAlignemnt `
                -ReferencePath ($referenceFrame.Weighted) `
                -OutputPath $AlignedOutputPath `
                -DetectionScales 7 `
                -KeepOpen
        }
    }
}
$aligned = $stats | where-object {$_.Aligned -and (Test-Path $_.Aligned) }
write-host "Aligned Stats"
$aligned |
    group-object Object,Filter,Exposure | 
    foreach-object {
        $targetName=$_.Values[0]
        $filter=$_.Values[1]
        $exposure=$_.Values[2]
        new-object psobject -Property @{
            Object=$targetName
            Filter=$filter
            Exposure=$exposure
            Exposures=$_.Group.Count
            ExposureTime=([TimeSpan]::FromSeconds($_.Group.Count*$exposure))
            Images=$_.Group
        } } |
    Sort-Object Object,Filter |
    Format-Table Object,Filter,Exposures,Exposure,ExposureTime,TopWeight

$aligned |
    group-object Object,Filter |    
    foreach-object {
        $targetName=$_.Values[0]
        $filter=$_.Values[1]
        new-object psobject -Property @{
            Object=$targetName
            Filter=$filter
            Images=$_.Group
        } } |
    ForEach-Object {
        $filter = $_.Filter
        $targetName=$_.Object
        $outputFileName = $targetName.Trim("'")
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
<#
$outputFile = Join-Path $target "SuperLum.xisf"
if(-not (test-path $outputFile) ) {
    write-host "Integrating $outputFile"
    Invoke-PiLightIntegration `
        -Images ($aligned | sort-object {
            $x = $_
            ($x.Aligned) -ne ($referenceFrame.Aligned)
        }|foreach-object {$_.Aligned}) `
        -OutputFile $outputFile `
        -PixInsightSlot 200 `
        -Rejection "LinearFit" `
        -LinearFitHigh 7 `
        -LinearFitLow 8 -KeepOpen
}

#>