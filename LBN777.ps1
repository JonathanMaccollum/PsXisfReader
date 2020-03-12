import-module $PSScriptRoot/pixinsightpreprocessing.psm1 -Force
import-module $PSScriptRoot/PsXisfReader.psm1

$target="D:\Backups\Camera\Astrophotography\1000mm\LBN777"
$CalibrationPath = "S:\PixInsight\Calibrated"
$WeightedOutputPath = "S:\PixInsight\Weighted"
$BackupCalibrationPaths = @("T:\PixInsightLT\Calibrated","N:\PixInsightLT\Calibrated")

$data = 
    Get-XisfLightFrames -Path "D:\Backups\Camera\Astrophotography\1000mm\LBN777" -Recurse |
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
<#
$('L','R','G','B')|foreach-object {
    Invoke-LightFrameSorting `
        -DropoffLocation "D:\Backups\Camera\Dropoff\NINA" -ArchiveDirectory "D:\Backups\Camera\Astrophotography" `
        -MasterDark "D:\Backups\Camera\ASI183mm-Pro\20191220\30x240s.-15C.G111.O8.masterdark.nocalibration.xisf" `
        -MasterFlat "D:\Backups\Camera\2019\20191101.Flats.Newt.Efw\$_.MasterFlat.xisf" `
        -Filter $_ -FocalLength 1000 -Exposure 240 -Gain 111 -Offset 8  `
        -OutputPath "S:\PixInsight\Calibrated" `
        -PixInsightSlot 200
    Invoke-LightFrameSorting `
        -DropoffLocation "D:\Backups\Camera\Dropoff\NINA" -ArchiveDirectory "D:\Backups\Camera\Astrophotography" `
        -MasterDark "D:\Backups\Camera\ASI183mm-Pro\20191220\30x120s.-15C.G111.O8.masterdark.nocalibration.xisf" `
        -MasterFlat "D:\Backups\Camera\2019\20191101.Flats.Newt.Efw\$_.MasterFlat.xisf" `
        -Filter $_ -FocalLength 1000 -Exposure 120 -Gain 111 -Offset 8  `
        -OutputPath "S:\PixInsight\Calibrated" `
        -PixInsightSlot 200
}
#>


$calibrated = $data |
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
            -WeightingExpression "(15*(1-(FWHM-FWHMMin)/(FWHMMax-FWHMMin))
            + 25*(1-(Eccentricity-EccentricityMin)/(EccentricityMax-EccentricityMin))
            + 15*(SNRWeight-SNRWeightMin)/(SNRWeightMax-SNRWeightMin)
            + 20*(1-(Median-MedianMin)/(MedianMax-MedianMin))
            + 10*(Stars-StarsMin)/(StarsMax-StarsMin))
            + 30"
        }
        
        $subframeResults = Get-Content -Path $resultCsv
        $subframeResults | Select-Object -Skip 29|Out-File -Path $resultData -Force
        $subframeResults | Select-Object -First 29
        $subframeData    = Import-Csv -Path "$resultData"|sort-object {[double]$_.Weight} -Descending
        $subframeData|FT
    }

    $stats = get-childitem $target *.data.csv |
        import-csv |
        foreach-object {
            $x=$_
            $y = Get-XisfFitsStats -Path ($x.File.Replace("/","\"))
            Add-Member -InputObject $y -Name "Approved" -MemberType NoteProperty -Value ([bool]$x.Approved)
            $y
        } |
        group-object Approved,Filter,Exposure |
        foreach-object {
            $approved=$_.Values[0]
            $filter=$_.Values[1]
            $exposure=$_.Values[2]
            new-object psobject -Property @{
                Approved=$approved
                Filter=$filter
                Exposure=$exposure
                Exposures=$_.Group.Count
                ExposureTime=([TimeSpan]::FromSeconds($_.Group.Count*$exposure))
                Images=$_.Group
            } } 
    $stats|
            Sort-Object Approved,Filter |
            Format-Table Approved,Filter,Exposures,Exposure,ExposureTime
    
    [TimeSpan]::FromSeconds((
            $stats|where-Object Approved -eq $true|foreach-object {$_.ExposureTime.TotalSeconds} |Measure-Object  -Sum).Sum).ToString()