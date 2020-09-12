$calibrated = 
    Get-XisfLightFrames -Path $target -Recurse |
    Where-Object {-not $_.HasTokensInPath(@("reject","process","testing","clouds","draft","cloudy"))} |
    Where-Object {-not $_.IsIntegratedFile()} |
    Get-XisfCalibrationState `
    -CalibratedPath $CalibrationPath `
    -AdditionalSearchPaths $BackupCalibrationPaths `
    -Verbose

$calibrated|
    foreach-object{
        $x=$_
        $stats = $x.Calibrated|Get-XisfFitsStats
        $flat = $stats.History | where-object {$_.StartsWith("ImageCalibration.masterFlat.fileName: ")}|%{$_.Split(": ")[1]}
        $dark = $stats.History | where-object {$_.StartsWith("ImageCalibration.masterDark.fileName: ")}|%{$_.Split(": ")[1]}
        $bias = $stats.History | where-object {$_.StartsWith("ImageCalibration.masterBias.fileName: ")}|%{$_.Split(": ")[1]}
        $outputPedestal = $stats.History | where-object {$_.StartsWith("ImageCalibration.outputPedestal: ")}|%{$_.Split(": ")[1]}
        new-object psobject -Property @{
            File=$x
            CalibrationStats=$stats
            Flat=$flat
            Dark=$dark
            Bias=$bias
            OutputPedestal = $outputPedestal
        }    
    } | 
    where-object OutputPedestal -ne 200 |
    group-object Bias,Dark,Flat,OutputPedestal |
    foreach-object {
        $x=$_
        $outputPedestal=$x.Group[0].OutputPedestal
        $flat=$x.Group[0].Flat
        $dark=$x.Group[0].Dark
        $bias=$x.Group[0].Bias
        $files = $x.Group|%{$_.File.Path}
        Write-Host "$($files.Count) x $flat $dark $outputPedestal"
        Invoke-PiLightCalibration `
            -KeepOpen `
            -PixInsightSlot 200 `
            -Images $files `
            -OutputPedestal 200 `
            -OutputPath "E:\PixInsightLT\Calibrated\PatchworkCygnus_0_1b" `
            -MasterDark "E:\Astrophotography\DarkLibrary\ZWO ASI183MM Pro\$dark" `
            -MasterFlat "E:\Astrophotography\50mm\Flats\$flat" `
            -Verbose
    }
