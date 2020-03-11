import-module $PSScriptRoot/pixinsightpreprocessing.psm1 -Force
import-module $PSScriptRoot/PsXisfReader.psm1

$target="D:\Backups\Camera\Astrophotography\1000mm\IC417"
$CalibrationPath = "S:\PixInsight\Calibrated"
$WeightedOutputPath = "S:\PixInsight\Weighted"
$BackupCalibrationPaths = @("T:\PixInsightLT\Calibrated","N:\PixInsightLT\Calibrated")

$data = get-childitem -path $target *.xisf -Recurse -File |
    where-object {-not $_.FullName.Contains("Rejection")} |
    foreach-object { Get-XisfFitsStats -Path $_ } |
    foreach-object {
        $x=$_
        $y = Get-CalibrationFile -Path ($x.Path) `
            -CalibratedPath $CalibrationPath `
            -AdditionalSearchPaths $BackupCalibrationPaths
        Add-Member -InputObject $x -Name "Calibrated" -MemberType NoteProperty -Value $y -Force
        $x
    }
$uncalibrated = $data | where-object {-not (test-path $_.Calibrated)}
if($uncalibrated){
    Write-Warning "$($uncalibrated.Length) uncalibrated frames detected"
    $uncalibrated|foreach-object {
        Write-Verbose ($_.Calibrated.Name)
    }
}
$calibrated = $data |
    where-object { $uncalibrated -notcontains $_ } |
    foreach-object {$_.Calibrated}

if(-not (test-path $target\Subframe.csv)) {
    Start-PiSubframeSelectorWeighting `
    -PixInsightSlot 200 `
    -OutputPath $WeightedOutputPath `
    -Images $calibrated
}

$subframeResults = Get-Content -Path "$target\Subframe.csv"
$subframeResults | Select-Object -Skip 29|Out-File -Path "$target\SubframeData.csv" -Force
$subframeResults | Select-Object -First 29
$subframeData    = Import-Csv -Path "$target\SubframeData.csv"|sort-object {[double]$_.Weight} -Descending
$subframeData|FT