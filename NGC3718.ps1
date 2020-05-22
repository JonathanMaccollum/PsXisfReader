#import-module $PSScriptRoot/pixinsightpreprocessing.psm1 -Force
import-module $PSScriptRoot/PsXisfReader.psd1 -force
$ErrorActionPreference="STOP"
$target="E:\Astrophotography\1000mm\NGC3718"
$CalibrationPath = "E:\PixInsightLT\Calibrated"
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
            if(-not (Read-Host "Continue?").ToLower().StartsWith("y")){
                break;
            }
        }
        else {
            $x
        }
    }

$data.Stats|group-object Filter,Exposure|foreach-object {
    $filter=$_.Group[0].Filter
    $exposure=$_.Group[0].Exposure
        new-object psobject -Property @{
            Filter=$filter
            Exposure=$exposure
            Exposures=$_.Group.Count
            ExposureTime=([TimeSpan]::FromSeconds($_.Group.Count*$exposure))
            Images=$_.Group
        }
} | Sort-Object Filter,ExposureTime|Format-Table Filter,Exposures,Exposure,ExposureTime

get-childitem $target *.data.csv |
        import-csv |
        foreach-object {
            $x = $_
            $y = $data|where-object {$_.Calibrated.Name -eq ([System.IO.FileInfo]$x.File).Name.Replace("/","\")}
            if($y) {
                $x|Get-Member -MemberType NoteProperty|where-object Name -ne File |foreach-object{
                $member=$_
                Add-Member -InputObject ($y.Stats) -Name ($member.Name) -MemberType NoteProperty -Value ($x."$($member.Name)") -Force
                }
            }
        }
exit
$summary = $data.Stats |
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

$data|
    where-object {$_.Stats.Approved -eq "true"}|
    group-object{$_.Stats.Filter} |
    foreach-object {
        $filter=$_.Name
        $images=$_.Group
        $total=$images.Stats|Measure-ExposureTime
        new-object psobject -Property @{
            Filter=$filter
            Total=$total
        }
    } | sort-object Total -descending | FT Filter,Total

$data|
    where-object {$_.Stats.Approved -eq "true"}|
    group-object{$_.Stats.Exposure} |
    foreach-object {
        $exposure=$_.Name
        $images=$_.Group
        $l = ($images.Stats|where-object Filter -eq "L" | Measure-Object).Count
        $r = ($images.Stats|where-object Filter -eq "R" | Measure-Object).Count
        $G = ($images.Stats|where-object Filter -eq "G" | Measure-Object).Count
        $b = ($images.Stats|where-object Filter -eq "B" | Measure-Object).Count
        $total=($images.Stats|Measure-Object).Count
        new-object psobject -Property @{
            Exposure = $exposure
            L=$l
            R=$r
            G=$g
            B=$b
            Total=$total
        }
    } |
    sort-object Exposure |
    Format-Table Exposure,L,R,G,B,Total

    
$data|
    where-object {$_.Stats.Approved -eq "true"}|
    group-object{$_.Stats.ObsDateMinus12hr} |
    #group-object{[DateTime]::Today} |
    foreach-object {
        $date=$_.Name
        $images=$_.Group
        $l = $images.Stats|where-object Filter -eq "L" | Measure-ExposureTime
        $r = $images.Stats|where-object Filter -eq "R" | Measure-ExposureTime
        $G = $images.Stats|where-object Filter -eq "G" | Measure-ExposureTime
        $b = $images.Stats|where-object Filter -eq "B" | Measure-ExposureTime
        $total=$images.Stats|Measure-ExposureTime
        new-object psobject -Property @{
            Date=(([DateTime]$date).ToString('yyyyMMdd'))
            L=$l
            R=$r
            G=$g
            B=$b
            Total=$total
        }
    } |
    sort-object Date |
    Format-Table Date,L,R,G,B,Total

exit


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