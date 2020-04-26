[CmdletBinding]
Function Wait-PixInsightInstance([int]$PixInsightSlot)
{
    $x = Get-Process -Name "pixinsight" 
    if($PixInsightSlot) {
        $x = $x | where-object MainWindowTitle -eq "PixInsight ($PixInsightSlot)"
    }
    if($x) {
        Add-Member -MemberType NoteProperty -Name "PixInsightSlot" -Value $PixInsightSlot -InputObject $x
        while(-not $x.WaitForExit(1000)){
            Write-Verbose "Waiting for PixInsighth slot $PixInsightSlot to exit."
        }
    }
    $x
}
Function Get-PixInsightInstance([int]$PixInsightSlot)
{
    $x = Get-Process -Name "pixinsight" 
    if($PixInsightSlot) {
        $x = $x | where-object MainWindowTitle -eq "PixInsight ($PixInsightSlot)"
    }
    if($x) {
        Add-Member -MemberType NoteProperty -Name "PixInsightSlot" -Value $PixInsightSlot -InputObject $x
    }
    $x
}
Function Invoke-PixInsightScript
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)][int]$PixInsightSlot, 
        [Parameter(Mandatory=$true)][System.IO.FileInfo]$path,
        [Parameter(Mandatory=$false)][Switch]$KeepOpen
    )

    Push-Location "C:\Program Files\PixInsight\bin"
    try
    {
        $arguments = @(
            "--no-startup-scripts"
            "--no-startup-check-updates",
            "--no-startup-gui-messages",
            "--no-splash -n=$PixInsightSlot"
            "--run=$($path.FullName.Replace('\','//'))"
        );
        if(-not ($KeepOpen.IsPresent)) {
            $arguments+="--force-exit"
        }
        Start-Process ".\pixinsight.exe" -PassThru -ArgumentList $arguments >> $null
        while(-not (Get-PixInsightInstance -PixInsightSlot $PixInsightSlot))
        {
            Write-Verbose "Waiting for PixInsight to start slot $PixInsightSlot"
            Wait-Event -Timeout 4            
        }
        Write-Verbose "Waiting for completion of slot $PixInsightSlot"
        Wait-PixInsightInstance -PixInsightSlot $PixInsightSlot >> $null
        Wait-Event -Timeout 4
        Write-Verbose "PixInsight slot $PixInsightSlot completed."
    }
    finally
    {
        Pop-Location
    }
}
Function Format-PiPath
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true,ValueFromPipeline=$true,ParameterSetName="File")][System.IO.FileInfo]$Path,
        [Parameter(Mandatory=$true,ValueFromPipeline=$true,ParameterSetName="Directory")][System.IO.DirectoryInfo]$Directory
    )
    process{
        if($Path) {
            $Path.FullName.Replace('\','/')
        }
        else{
            $Directory.FullName.Replace('\','/')
        }
    
    }    
}
Function Invoke-PIIntegrationScript
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)][int]$PixInsightSlot, 
        [Parameter(Mandatory=$true)][System.IO.FileInfo]$Path, 
        [Parameter(Mandatory=$true)][System.IO.FileInfo]$OutputFile,
        [Parameter(Mandatory=$false)][Switch]$KeepOpen
    )
    $scp = $Path|Format-PiPath
    $output=$OutputFile|Format-PiPath
    $Template="run -x $scp
save integration -p=`"$output`" --nodialog --nomessages --noverify"
    $ScriptToRun = New-TemporaryFile
    $ScriptToRun = Rename-Item ($ScriptToRun.FullName) ($ScriptToRun.FullName+".scp") -PassThru
    $Template|Out-File $ScriptToRun -Force
    try {
        Write-Debug "Invoking Script $(Get-Content $Path)"
        Invoke-PixInsightScript `
        -PixInsightSlot $PixInsightSlot `
        -Path $ScriptToRun `
        -KeepOpen:$KeepOpen

        $OutputFile.Refresh()
        if($OutputFile.Exists) {
            Write-Host "Integration image saved successfully to $($OutputFile.FullName)"
        }
        else{
            Write-Error "Pixinsight exited without producing expected file."

        }
    }
    finally {
        Remove-Item $ScriptToRun -Force
    }
}
Function Invoke-PICalibrationScript
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)][int]$PixInsightSlot, 
        [Parameter(Mandatory=$true)][System.IO.FileInfo]$Path, 
        [Parameter(Mandatory=$false)][Switch]$KeepOpen
    )
    $scp = $Path|Format-PiPath
    $Template="run -x $scp"
    $ScriptToRun = New-TemporaryFile
    $ScriptToRun = Rename-Item ($ScriptToRun.FullName) ($ScriptToRun.FullName+".scp") -PassThru
    $Template|Out-File $ScriptToRun -Force
    try {
        Invoke-PixInsightScript `
        -PixInsightSlot $PixInsightSlot `
        -Path $ScriptToRun `
        -KeepOpen:$KeepOpen
    }
    finally {
        Remove-Item $ScriptToRun -Force
    }
}
Function Invoke-PiDarkIntegration
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)][int]$PixInsightSlot,
        [Parameter(Mandatory=$true)][System.IO.FileInfo[]]$Images,
        [Parameter(Mandatory=$true)][System.IO.FileInfo]$OutputFile,
        [Parameter(Mandatory=$false)][Switch]$KeepOpen
    )
    $ImageDefinition = [string]::Join("`r`n   , ",
    ($Images | ForEach-Object {
        $x=$_|Format-PiPath
        "[true, ""$x"", """", """"]"
    }))
    $IntegrationDefinition = 
    "var P = new ImageIntegration;
    P.combination = ImageIntegration.prototype.Average;
    P.weightMode = ImageIntegration.prototype.DontCare;
    P.weightScale = ImageIntegration.prototype.WeightScale_IKSS;
    P.ignoreNoiseKeywords = false;
    P.normalization = ImageIntegration.prototype.NoNormalization;
    P.rejection = ImageIntegration.prototype.LinearFit;
    P.rejectionNormalization = ImageIntegration.prototype.NoRejectionNormalization;
    P.linearFitLow = 5.000;
    P.linearFitHigh = 2.500;
    P.clipLow = true;
    P.clipHigh = true;
    P.rangeClipLow = false;
    P.rangeLow = 0.000000;
    P.rangeClipHigh = false;
    P.rangeHigh = 0.980000;
    P.mapRangeRejection = true;
    P.reportRangeRejection = false;
    P.generate64BitResult = false;
    P.generateRejectionMaps = false;
    P.generateIntegratedImage = true;
    P.generateDrizzleData = false;
    P.closePreviousImages = true;
    P.bufferSizeMB = 16;
    P.stackSizeMB = 1024;
    P.autoMemorySize = false;
    P.autoMemoryLimit = 0.75;
    P.useROI = false;
    P.useCache = true;
    P.evaluateNoise = false;
    P.mrsMinDataFraction = 0.010;
    P.subtractPedestals = true;
    P.truncateOnOutOfRange = false;
    P.noGUIMessages = true;
    P.useFileThreads = true;
    P.fileThreadOverload = 1.00;
    P.images= [`r`n     $ImageDefinition`r`n   ];
    P.executeGlobal();"
    $executionScript = New-TemporaryFile
    $executionScript = Rename-Item ($executionScript.FullName) ($executionScript.FullName+".js") -PassThru
    try {
        $IntegrationDefinition|Out-File -FilePath $executionScript -Force 
        Invoke-PIIntegrationScript `
            -path $executionScript `
            -PixInsightSlot $PixInsightSlot `
            -OutputFile $OutputFile `
            -KeepOpen:$KeepOpen
    }
    finally {
        Remove-Item $executionScript -Force
    }
} 
Function Invoke-PiFlatIntegration
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)][int]$PixInsightSlot,
        [Parameter(Mandatory=$true)][System.IO.FileInfo[]]$Images,
        [Parameter(Mandatory=$true)][System.IO.FileInfo]$OutputFile,
        [Parameter(Mandatory=$false)][Switch]$KeepOpen
    )
    $ImageDefinition = [string]::Join("`r`n   , ",
    ($Images | ForEach-Object {
        $x=$_|Format-PiPath
        "[true, ""$x"", """", """"]"
    }))
    $IntegrationDefinition = 
    "var P = new ImageIntegration;
    P.combination = ImageIntegration.prototype.Average;
    P.weightMode = ImageIntegration.prototype.DontCare;
    P.weightScale = ImageIntegration.prototype.WeightScale_IKSS;
    P.ignoreNoiseKeywords = false;
    P.normalization = ImageIntegration.prototype.Multiplicative;
    P.rejection = ImageIntegration.prototype.LinearFit;
    P.rejectionNormalization = ImageIntegration.prototype.EqualizeFluxes;
    P.linearFitLow = 5.000;
    P.linearFitHigh = 2.500;
    P.clipLow = true;
    P.clipHigh = true;
    P.rangeClipLow = false;
    P.rangeLow = 0.000000;
    P.rangeClipHigh = false;
    P.rangeHigh = 0.980000;
    P.mapRangeRejection = true;
    P.reportRangeRejection = false;
    P.generate64BitResult = false;
    P.generateRejectionMaps = false;
    P.generateIntegratedImage = true;
    P.generateDrizzleData = false;
    P.closePreviousImages = true;
    P.bufferSizeMB = 16;
    P.stackSizeMB = 1024;
    P.autoMemorySize = false;
    P.autoMemoryLimit = 0.75;
    P.useROI = false;
    P.useCache = true;
    P.evaluateNoise = false;
    P.mrsMinDataFraction = 0.010;
    P.subtractPedestals = true;
    P.truncateOnOutOfRange = false;
    P.noGUIMessages = true;
    P.useFileThreads = true;
    P.fileThreadOverload = 1.00;
    P.images= [`r`n     $ImageDefinition`r`n   ];
    P.executeGlobal();"
    $executionScript = New-TemporaryFile
    $executionScript = Rename-Item ($executionScript.FullName) ($executionScript.FullName+".js") -PassThru
    try {
        $IntegrationDefinition|Out-File -FilePath $executionScript -Force 
        Invoke-PIIntegrationScript `
            -path $executionScript `
            -PixInsightSlot $PixInsightSlot `
            -OutputFile $OutputFile `
            -KeepOpen:$KeepOpen
    }
    finally {
        Remove-Item $executionScript -Force
    }
} 
Function Invoke-PiFlatCalibration
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)][int]$PixInsightSlot,
        [Parameter(Mandatory=$true)][System.IO.FileInfo[]]$Images,
        [Parameter(Mandatory=$true)][System.IO.FileInfo]$MasterDark,
        [Parameter(Mandatory=$true)][System.IO.DirectoryInfo]$OutputPath,
        [Parameter(Mandatory=$false)][Switch]$KeepOpen
    )
    $masterDarkPath = Get-Item $MasterDark | Format-PiPath
    $outputDirectory = Get-Item ($OutputPath.FullName) | Format-PiPath
    $ImageDefinition = [string]::Join("`r`n   , ",
    ($Images | ForEach-Object {
        $x=$_|Format-PiPath
        "[true, ""$x""]"
    }))
    $IntegrationDefinition = 
    "var P = new ImageCalibration;
P.pedestal = 0;
P.pedestalMode = ImageCalibration.prototype.Keyword;
P.masterBiasEnabled = false;
P.masterDarkEnabled = true;
P.masterDarkPath = `"$masterDarkPath`";
P.masterFlatEnabled = false;
P.calibrateBias = false;
P.calibrateDark = false;
P.calibrateFlat = false;
P.optimizeDarks = false;
P.darkOptimizationThreshold = 0.00000;
P.darkOptimizationLow = 3.0000;
P.darkOptimizationWindow = 1024;
P.darkCFADetectionMode = ImageCalibration.prototype.DetectCFA;
P.evaluateNoise = true;
P.noiseEvaluationAlgorithm = ImageCalibration.prototype.NoiseEvaluation_MRS;
P.outputDirectory = `"$outputDirectory`";
P.outputExtension = `".xisf`";
P.outputPostfix = `"_c`";
P.outputSampleFormat = ImageCalibration.prototype.f32;
P.outputPedestal = 0;
P.overwriteExistingFiles = false;
P.onError = ImageCalibration.prototype.Abort;
P.noGUIMessages = true;
    P.targetFrames= [`r`n     $ImageDefinition`r`n   ];
    P.executeGlobal();"
    $executionScript = New-TemporaryFile
    $executionScript = Rename-Item ($executionScript.FullName) ($executionScript.FullName+".js") -PassThru
    try {
        $IntegrationDefinition|Out-File -FilePath $executionScript -Force 
        Invoke-PICalibrationScript `
            -path $executionScript `
            -PixInsightSlot $PixInsightSlot `
            -KeepOpen:$KeepOpen
    }
    finally {
        Remove-Item $executionScript -Force
    }
} 

Function Invoke-PiLightCalibration
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)][int]$PixInsightSlot,
        [Parameter(Mandatory=$true)][System.IO.FileInfo[]]$Images,
        [Parameter(Mandatory=$true)][System.IO.FileInfo]$MasterDark,
        [Parameter(Mandatory=$true)][System.IO.FileInfo]$MasterFlat,
        [Parameter(Mandatory=$true)][System.IO.DirectoryInfo]$OutputPath,
        [Parameter(Mandatory=$false)][Switch]$KeepOpen,
        [Parameter()][int]$OutputPedestal=0
    )
    $masterDarkPath = Get-Item $MasterDark | Format-PiPath
    $masterFlatPath = Get-Item $MasterFlat | Format-PiPath
    $outputDirectory = Get-Item ($OutputPath.FullName) | Format-PiPath
    $ImageDefinition = [string]::Join("`r`n   , ",
    ($Images | ForEach-Object {
        $x=$_|Format-PiPath
        "[true, ""$x""]"
    }))
    $IntegrationDefinition = 
    "var P = new ImageCalibration;
P.pedestal = 0;
P.pedestalMode = ImageCalibration.prototype.Keyword;
P.masterBiasEnabled = false;
P.masterDarkEnabled = true;
P.masterDarkPath = `"$masterDarkPath`";
P.masterFlatEnabled = true;
P.masterFlatPath = `"$masterFlatPath`";
P.calibrateBias = false;
P.calibrateDark = false;
P.calibrateFlat = false;
P.optimizeDarks = false;
P.darkOptimizationThreshold = 0.00000;
P.darkOptimizationLow = 3.0000;
P.darkOptimizationWindow = 1024;
P.darkCFADetectionMode = ImageCalibration.prototype.DetectCFA;
P.evaluateNoise = true;
P.noiseEvaluationAlgorithm = ImageCalibration.prototype.NoiseEvaluation_MRS;
P.outputDirectory = `"$outputDirectory`";
P.outputExtension = `".xisf`";
P.outputPostfix = `"_c`";
P.outputSampleFormat = ImageCalibration.prototype.f32;
P.outputPedestal = $OutputPedestal;
P.overwriteExistingFiles = false;
P.onError = ImageCalibration.prototype.Abort;
P.noGUIMessages = true;
    P.targetFrames= [`r`n     $ImageDefinition`r`n   ];
    P.executeGlobal();"
    $executionScript = New-TemporaryFile
    $executionScript = Rename-Item ($executionScript.FullName) ($executionScript.FullName+".js") -PassThru
    try {
        $IntegrationDefinition|Out-File -FilePath $executionScript -Force 
        Invoke-PICalibrationScript `
            -path $executionScript `
            -PixInsightSlot $PixInsightSlot `
            -KeepOpen:$KeepOpen
    }
    finally {
        Remove-Item $executionScript -Force
    }
} 

Function Invoke-DarkFrameSorting
{
    [CmdletBinding()]
    param (
        [System.IO.DirectoryInfo]$DropoffLocation,
        [System.IO.DirectoryInfo]$ArchiveDirectory,
        [int]$PixInsightSlot
    )

    Get-ChildItem $DropoffLocation "*.xisf" |
    foreach-object { Get-XisfFitsStats -Path $_ } |
    where-object ImageType -eq "DARK" |
    group-object Instrument,Exposure,SetTemp,Gain,Offset |
    foreach-object {
        $images=$_.Group
        $instrument=$images[0].Instrument.Trim().Trim("'")
        $exposure=$images[0].Exposure.Trim()
        $setTemp=($images[0].SetTemp)
        $gain=($images[0].Gain)
        $offset=($images[0].Offset)
        $date = ([DateTime]$images[0].LocalDate).Date.ToString('yyyyMMdd')
        $targetDirectory = "$ArchiveDirectory\DarkLibrary\$instrument"
        [System.IO.Directory]::CreateDirectory($targetDirectory)>>$null
        $masterDark = "$targetDirectory\$($date).MasterDark.Gain.$($gain).Offset.$($offset).$($setTemp)C.$($images.Count)x$($exposure)s.xisf"
        if(-not (Test-Path $masterDark)) {
            Write-Host "Integrating $($images.Count) Darks for $instrument Gain $gain Offset $offset at $setTemp duration $($exposure)s"
            $darkFlats = $images | foreach-object {$_.Path}
            Invoke-PiDarkIntegration `
                -Images $darkFlats `
                -PixInsightSlot $PixInsightSlot `
                -OutputFile $masterDark
        }
        $destinationDirectory=join-path $targetDirectory $date
        [System.IO.Directory]::CreateDirectory($destinationDirectory)>>$null
        $images | Foreach-Object {
            Move-Item -Path $_.Path -Destination $destinationDirectory
        }
    }
}

Function Invoke-DarkFlatFrameSorting
{
    [CmdletBinding()]
    param (
        [System.IO.DirectoryInfo]$DropoffLocation,
        [System.IO.DirectoryInfo]$ArchiveDirectory,
        [int]$PixInsightSlot
    )

    Get-ChildItem $DropoffLocation "*.xisf" |
    foreach-object { Get-XisfFitsStats -Path $_ } |
    where-object ImageType -eq "DARKFLAT" |
    group-object Filter,FocalLength |
    foreach-object {
        $images=$_.Group
        $filter=$images[0].Filter.Trim()
        $focalLength=($images[0].FocalLength).TrimEnd('mm')+'mm'
        $flatDate = ([DateTime]$images[0].LocalDate).Date.ToString('yyyyMMdd')
        $targetDirectory = "$ArchiveDirectory\$focalLength\Flats"
        $masterDark = "$targetDirectory\$($flatDate).MasterDarkFlat.$($filter).xisf"
        if(-not (Test-Path $masterDark)) {
            Write-Host "Integrating $($images.Count) Dark Flats for $filter at $focalLength"
            $darkFlats = $images | foreach-object {$_.Path}
            Invoke-PiDarkIntegration `
                -Images $darkFlats `
                -PixInsightSlot $PixInsightSlot `
                -OutputFile $masterDark
        }
        $destinationDirectory=join-path $targetDirectory $flatDate
        [System.IO.Directory]::CreateDirectory($destinationDirectory)>>$null
        $images | Foreach-Object {
            Move-Item -Path $_.Path -Destination $destinationDirectory
        }
    }
}
Function Get-XisfFile{
    [CmdletBinding()]
    param (
        [System.IO.DirectoryInfo]$Path,
        [Switch]$Recurse
    )
    Get-ChildItem -Path $Path -File -Filter *.xisf -Recurse:$Recurse |
        Get-XisfFitsStats -Path $_
    
}

Function Get-XisfLightFrames{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)][System.IO.DirectoryInfo]$Path,
        [Parameter()][Switch]$Recurse
    )
    Get-ChildItem -Path $Path -File -Filter *.xisf -Recurse:$Recurse |
        Get-XisfFitsStats -Path $_ |
        where-object ImageType -eq "LIGHT"
}
Function Invoke-FlatFrameSorting
{
    [CmdletBinding()]
    param (
        [System.IO.DirectoryInfo]$DropoffLocation,
        [System.IO.DirectoryInfo]$ArchiveDirectory,
        [System.IO.DirectoryInfo]$CalibratedFlatsOutput,
        [int]$PixInsightSlot
    )
    Get-ChildItem $DropoffLocation "*.xisf" |
    foreach-object { Get-XisfFitsStats -Path $_ } |
    where-object ImageType -eq "FLAT" |
    group-object Filter,FocalLength |
    foreach-object {
        $flats=$_.Group
        $filter=$flats[0].Filter.Trim()
        $focalLength=($flats[0].FocalLength).TrimEnd('mm')+'mm'
        $flatDate = ([DateTime]$flats[0].LocalDate).Date.ToString('yyyyMMdd')

        Write-Host "Processing $($flats.Count) Flats for filter '$filter' at $focalLength"

        $targetDirectory = "$ArchiveDirectory\$focalLength\Flats"
        $masterDark = "$targetDirectory\$flatDate.MasterDarkFlat.$filter.xisf"
        if(Test-Path $masterDark) {
            $masterCalibratedFlat = "$targetDirectory\$flatDate.MasterFlatCal.$filter.xisf"
            if(-not (Test-Path $masterCalibratedFlat)) {
                $flats | foreach-object{
                    $input = $_.Path
                    $output = Join-Path ($CalibratedFlatsOutput.FullName) ($input.Name.Replace($input.Extension,"")+"_c.xisf")
                    Add-Member -InputObject $_ -Name "CalibratedFlat" -MemberType NoteProperty -Value $output -Force
                }
                $toCalibrate = $flats | where-object {-not (Test-Path $_.CalibratedFlat)} | foreach-object {$_.Path}
                if($toCalibrate){
                    Invoke-PiFlatCalibration `
                        -Images $toCalibrate `
                        -MasterDark $masterDark `
                        -OutputPath $CalibratedFlatsOutput `
                        -PixInsightSlot $PixInsightSlot
                }
                $calibratedFlats = $flats|foreach-object {Get-Item $_.CalibratedFlat}
                if($flats) {
                    Invoke-PiFlatIntegration `
                        -Images $calibratedFlats `
                        -PixInsightSlot $PixInsightSlot `
                        -OutputFile $masterCalibratedFlat
                }
            }
        }

        $masterNoCalFlat = "$targetDirectory\$FlatDate.MasterFlatNoCal.$($filter).xisf"
        $toIntegrate = $flats|foreach-object {$_.Path}
        if(-not (Test-Path $masterNoCalFlat)) {
            Invoke-PiFlatIntegration `
                -Images $toIntegrate `
                -PixInsightSlot $PixInsightSlot `
                -OutputFile $masterNoCalFlat
        }
        $destinationDirectory=join-path $targetDirectory $flatDate
        [System.IO.Directory]::CreateDirectory($destinationDirectory)>>$null
        $flats | Foreach-Object {Move-Item -Path $_.Path -Destination $destinationDirectory}
    }
}


Function Get-CalibrationFile
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)][System.IO.FileInfo]$Path,
        [Parameter(Mandatory)][System.IO.DirectoryInfo]$CalibratedPath,
        [Parameter()][System.IO.DirectoryInfo[]]$AdditionalSearchPaths
    )
    $calibratedFileName = $Path.Name.Substring(0,$Path.Name.Length-$Path.Extension.Length)+"_c.xisf"
    $calibrated = Join-Path ($CalibratedPath.FullName) ($calibratedFileName)
    if(-not (test-path $calibrated) -and ($AdditionalSearchPaths)){
        $calibrated = $AdditionalSearchPaths | 
            foreach-object { join-path $_ $calibratedFileName } |
            where-object { test-path $_ } | 
            select-object -first 1
    }
    $calibrated
}

Function Invoke-LightFrameSorting
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory,ParameterSetName="ByDirectory")][System.IO.DirectoryInfo]$DropoffLocation,
        [Parameter(Mandatory,ParameterSetName="ByFiles")][Object[]]$XisfStats,
        [Parameter(Mandatory)][System.IO.DirectoryInfo]$ArchiveDirectory,

        [Parameter(Mandatory,ParameterSetName="ByDirectory")][String]$Filter,
        [Parameter(Mandatory,ParameterSetName="ByDirectory")][int]$FocalLength,
        [Parameter(Mandatory,ParameterSetName="ByDirectory")][int]$Exposure,
        [Parameter(Mandatory,ParameterSetName="ByDirectory")][int]$Gain,
        [Parameter(Mandatory,ParameterSetName="ByDirectory")][int]$Offset,
        [Parameter(Mandatory,ParameterSetName="ByDirectory")][int]$SetTemp,

        [Parameter(Mandatory)][System.IO.FileInfo]$MasterDark,
        [Parameter(Mandatory)][System.IO.FileInfo]$MasterFlat,
        [Parameter(Mandatory)][System.IO.DirectoryInfo]$OutputPath,

        [Parameter(Mandatory)][int]$PixInsightSlot,
        [System.IO.DirectoryInfo[]]$AdditionalSearchPaths,
        [Parameter(Mandatory=$false)][int]$OutputPedestal = 0,
        [Parameter(Mandatory=$false)][Switch]$KeepOpen
    )
    if($XisfStats -eq $null)
    {
        $XisfStats = Get-ChildItem $DropoffLocation "*.xisf" -File |
            Get-XisfFitsStats |
            where-object ImageType -eq "Light" |
            where-object SetTemp -eq $SetTemp |
            where-object Filter -eq $Filter |
            where-object Exposure -eq $Exposure |
            where-object CCDGain -eq $CCDGain |
            where-object Offset -eq $Offset |
            where-object FocalLength -eq $FocalLength
    }
    else{
        $Filter=$XisfStats[0].Filter
        $FocalLength=$XisfStats[0].FocalLength
        $Exposure=$XisfStats[0].Exposure
        $Gain=$XisfStats[0].Gain
        $Offset=$XisfStats[0].Offset
        $SetTemp=$XisfStats[0].SetTemp
    }
    $XisfStats |
    Group-Object Object |
    foreach-object {
        $object = $_.Name

        $toCalibrate = $_.Group |
            foreach-object {
                $x = $_
                $calibrated = Get-CalibrationFile -Path $_.Path -CalibratedPath $OutputPath -AdditionalSearchPaths $AdditionalSearchPaths
                Add-Member -InputObject $x -Name "Calibrated" -MemberType NoteProperty -Value $calibrated -Force
                $x
            } |
            where-object {-not (Test-Path $_.Calibrated)} | foreach-object {$_.Path}
        if($toCalibrate){
            Write-Host "Calibrating $($toCalibrate.Count)x$($Exposure)s $Filter Frames for target $object"
            Invoke-PiLightCalibration `
                -Images $toCalibrate `
                -MasterDark $MasterDark `
                -MasterFlat $MasterFlat `
                -OutputPath $OutputPath `
                -OutputPedestal $OutputPedestal `
                -PixInsightSlot $PixInsightSlot `
                -KeepOpen:$KeepOpen
        }

        $archive = Join-Path $ArchiveDirectory "$($FocalLength)mm" 
        $archiveObjectWithoutSpaces = Join-Path $archive ($object.Replace(' ',''))
        if(test-path $archiveObjectWithoutSpaces){
            $archive = $archiveObjectWithoutSpaces
        }else{
            $archive = join-path $archive $object
        }
        $archive = join-path $archive "$($exposure).00s"
        [System.IO.Directory]::CreateDirectory($archive) >> $null

        $_.Group | Foreach-Object {Move-Item -Path $_.Path -Destination $archive}
    }
}
Function Start-PiSubframeSelectorWeighting
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)][int]$PixInsightSlot,
        [Parameter(Mandatory=$true)][System.IO.FileInfo[]]$Images,
        [Parameter(Mandatory=$true)][System.IO.DirectoryInfo]$OutputPath,
        [Parameter(Mandatory=$false)][String]$WeightingExpression="(15*(1-(FWHM-FWHMMin)/(FWHMMax-FWHMMin))
        + 25*(1-(Eccentricity-EccentricityMin)/(EccentricityMax-EccentricityMin))
        + 15*(SNRWeight-SNRWeightMin)/(SNRWeightMax-SNRWeightMin)
        + 20*(1-(Median-MedianMin)/(MedianMax-MedianMin))
        + 10*(Stars-StarsMin)/(StarsMax-StarsMin))
        + 30",
        [Parameter(Mandatory=$false)][String]$ApprovalExpression="Median<70"
    )
    $AE = "`""+
        [String]::Join("\n`"`r`n +`"",
        [Regex]::Split($ApprovalExpression, "`r`n|`r|`n")
        )+"`""
    $WE = "`""+
        [String]::Join("\n`"`r`n +`"",
        [Regex]::Split($WeightingExpression, "`r`n|`r|`n")
        )+"`""
    $outputDirectory = Get-Item ($OutputPath.FullName) | Format-PiPath
    $ImageDefinition = [string]::Join("`r`n   , ",
    ($Images | ForEach-Object {
        $x=$_|Format-PiPath
        "[true, ""$x""]"
    }))
    $CommandDefinition = 
    "var P = new SubframeSelector;
    P.routine = SubframeSelector.prototype.MeasureSubframes;
    P.fileCache = true;
    P.subframeScale = 0.4800;
    P.cameraGain = 1.0000;
    P.cameraResolution = SubframeSelector.prototype.Bits12;
    P.siteLocalMidnight = 24;
    P.scaleUnit = SubframeSelector.prototype.ArcSeconds;
    P.dataUnit = SubframeSelector.prototype.Electron;
    P.structureLayers = 5;
    P.noiseLayers = 0;
    P.hotPixelFilterRadius = 1;
    P.applyHotPixelFilter = false;
    P.noiseReductionFilterRadius = 0;
    P.sensitivity = 0.1000;
    P.peakResponse = 0.8000;
    P.maxDistortion = 0.5000;
    P.upperLimit = 1.0000;
    P.backgroundExpansion = 3;
    P.xyStretch = 1.5000;
    P.psfFit = SubframeSelector.prototype.Gaussian;
    P.psfFitCircular = false;
    P.pedestal = 0;
    P.roiX0 = 0;
    P.roiY0 = 0;
    P.roiX1 = 0;
    P.roiY1 = 0;
    P.inputHints = `"`";
    P.outputHints = `"`";
    P.outputDirectory = `"$outputDirectory`";
    P.outputExtension = `".xisf`";
    P.outputPrefix = `"`";
    P.outputPostfix = `"_a`";
    P.outputKeyword = `"SSWEIGHT`";
    P.overwriteExistingFiles = false;
    P.onError = SubframeSelector.prototype.Continue;
    P.approvalExpression = $AE;
    P.weightingExpression = $WE;
    P.sortProperty = SubframeSelector.prototype.Weight;
    P.graphProperty = SubframeSelector.prototype.Median;
    P.subframes = [`r`n     $ImageDefinition`r`n   ];
    P.measurements = [ // measurementIndex, measurementEnabled, measurementLocked, measurementPath, measurementWeight, measurementFWHM, measurementEccentricity, measurementSNRWeight, measurementMedian, measurementMedianMeanDev, measurementNoise, measurementNoiseRatio, measurementStars, measurementStarResidual, measurementFWHMMeanDev, measurementEccentricityMeanDev, measurementStarResidualMeanDev
    ];
    P.launch();
    P.executeGlobal();"
    $executionScript = New-TemporaryFile
    $executionScript = Rename-Item ($executionScript.FullName) ($executionScript.FullName+".js") -PassThru
    try {
        $CommandDefinition|Out-File -FilePath $executionScript -Force 
        Invoke-PICalibrationScript `
            -path $executionScript `
            -PixInsightSlot $PixInsightSlot `
            -KeepOpen
    }
    finally {
        Remove-Item $executionScript -Force
    }
}

Function Invoke-PiStarAlignment
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)][int]$PixInsightSlot,
        [Parameter(Mandatory=$true)][System.IO.FileInfo[]]$Images,
        [Parameter(Mandatory=$true)][System.IO.FileInfo]$ReferencePath,
        [Parameter(Mandatory=$true)][System.IO.DirectoryInfo]$OutputPath,
        [Parameter(Mandatory=$false)][Switch]$KeepOpen,
        [Parameter(Mandatory=$false)][int]$DetectionScales=5
    )
    $piReferencePath = Get-Item $ReferencePath | Format-PiPath
    $outputDirectory = Get-Item ($OutputPath.FullName) | Format-PiPath
    $ImageDefinition = [string]::Join("`r`n   , ",
    ($Images | ForEach-Object {
        $x=$_|Format-PiPath
        "[true, true, ""$x""]"
    }))
    $CommandDefinition = 
    "
var P = new StarAlignment;
P.structureLayers = $DetectionScales;
P.noiseLayers = 0;
P.hotPixelFilterRadius = 1;
P.noiseReductionFilterRadius = 0;
P.sensitivity = 0.100;
P.peakResponse = 0.80;
P.maxStarDistortion = 0.500;
P.upperLimit = 1.000;
P.invert = false;
P.distortionModel = `"`";
P.undistortedReference = false;
P.distortionCorrection = true;
P.distortionMaxIterations = 20;
P.distortionTolerance = 0.005;
P.distortionAmplitude = 2;
P.localDistortion = true;
P.localDistortionScale = 256;
P.localDistortionTolerance = 0.050;
P.localDistortionRejection = 2.50;
P.localDistortionRejectionWindow = 64;
P.localDistortionRegularization = 0.010;
P.matcherTolerance = 0.0500;
P.ransacTolerance = 2.00;
P.ransacMaxIterations = 2000;
P.ransacMaximizeInliers = 1.00;
P.ransacMaximizeOverlapping = 1.00;
P.ransacMaximizeRegularity = 1.00;
P.ransacMinimizeError = 1.00;
P.maxStars = 0;
P.fitPSF = StarAlignment.prototype.FitPSF_DistortionOnly;
P.psfTolerance = 0.50;
P.useTriangles = false;
P.polygonSides = 5;
P.descriptorsPerStar = 20;
P.restrictToPreviews = false;
P.intersection = StarAlignment.prototype.MosaicOnly;
P.useBrightnessRelations = false;
P.useScaleDifferences = false;
P.scaleTolerance = 0.100;
P.referenceImage = `"$piReferencePath`";
P.referenceIsFile = true;
P.inputHints = `"`";
P.outputHints = `"`";
P.mode = StarAlignment.prototype.RegisterMatch;
P.writeKeywords = true;
P.generateMasks = false;
P.generateDrizzleData = true;
P.generateDistortionMaps = false;
P.frameAdaptation = false;
P.randomizeMosaic = false;
P.noGUIMessages = true;
P.useSurfaceSplines = false;
P.extrapolateLocalDistortion = true;
P.splineSmoothness = 0.250;
P.pixelInterpolation = StarAlignment.prototype.Auto;
P.clampingThreshold = 0.30;
P.outputDirectory = `"S:/PixInsight/Aligned`";
P.outputExtension = `".xisf`";
P.outputPrefix = `"`";
P.outputPostfix = `"_r`";
P.maskPostfix = `"_m`";
P.distortionMapPostfix = `"_dm`";
P.outputSampleFormat = StarAlignment.prototype.SameAsTarget;
P.overwriteExistingFiles = false;
P.onError = StarAlignment.prototype.Continue;
P.useFileThreads = true;
P.fileThreadOverload = 1.20;
P.maxFileReadThreads = 1;
P.maxFileWriteThreads = 1;
P.outputDirectory = `"$outputDirectory`";
P.targets= [`r`n     $ImageDefinition`r`n   ];
P.launch();
P.executeGlobal();"
    $executionScript = New-TemporaryFile
    $executionScript = Rename-Item ($executionScript.FullName) ($executionScript.FullName+".js") -PassThru
    try {
        $CommandDefinition|Out-File -FilePath $executionScript -Force 
        Invoke-PICalibrationScript `
            -path $executionScript `
            -PixInsightSlot $PixInsightSlot `
            -KeepOpen:$KeepOpen
    }
    finally {
        Remove-Item $executionScript -Force
    }
}

Function Start-PiCometAlignment
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)][int]$PixInsightSlot,
        [Parameter(Mandatory=$true)][System.IO.FileInfo[]]$Images,
        [Parameter(Mandatory=$true)][System.IO.DirectoryInfo]$OutputPath
    )
    $outputDirectory = Get-Item ($OutputPath.FullName) | Format-PiPath
    $ImageDefinition = [string]::Join("`r`n   , ",
    ($Images | ForEach-Object {
        $fitsData=Get-XisfFitsStats -Path $_
        $obsDate=$fitsData.ObsDate.Trim().Trim('''')
        $jDate=([DateTime]$obsDate).ToOADate()+2415018.5
        $x=$_|Format-PiPath
        "[true, ""$x"", ""$($obsDate)"", $jDate, 0.00000000, 0.00000000, """"]"
    }))
    $CommandDefinition = 
    "var P = new CometAlignment;
    P.inputHints = `"`";
    P.outputHints = `"`";
    P.outputDir = `"$outputDirectory`";
    P.outputExtension = `".xisf`";
    P.prefix = `"`";
    P.postfix = `"_a`";
    P.overwrite = false;
    P.reference = 0;
    P.subtractFile = `"`";
    P.subtractMode = true;
    P.enableLinearFit = true;
    P.rejectLow = 0.000000;
    P.rejectHigh = 0.920000;
    P.normalize = true;
    P.drzSaveStarsAligned = true;
    P.drzSaveCometAligned = true;
    P.operandIsDI = true;
    P.pixelInterpolation = CometAlignment.prototype.Lanczos3;
    P.linearClampingThreshold = 0.30;
    
    P.targetFrames = [`r`n     $ImageDefinition`r`n   ];
    P.launch();
    "
    $executionScript = New-TemporaryFile
    $executionScript = Rename-Item ($executionScript.FullName) ($executionScript.FullName+".js") -PassThru
    try {
        $CommandDefinition|Out-File -FilePath $executionScript -Force 
        Invoke-PICalibrationScript `
            -path $executionScript `
            -PixInsightSlot $PixInsightSlot `
            -KeepOpen
    }
    finally {
        Remove-Item $executionScript -Force
    }
}




Function Invoke-PiLightIntegration
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)][int]$PixInsightSlot,
        [Parameter(Mandatory=$true)][System.IO.FileInfo[]]$Images,
        [Parameter(Mandatory=$true)][System.IO.FileInfo]$OutputFile,
        [Parameter(Mandatory=$false)][Switch]$KeepOpen,
        [Parameter(Mandatory=$false)][decimal]$LinearFitLow=8,
        [Parameter(Mandatory=$false)][decimal]$LinearFitHigh=7,
        [Parameter(Mandatory=$false)][string]$Rejection = "LinearFit"
    )
    $ImageDefinition = [string]::Join("`r`n   , ",
    ($Images | ForEach-Object {
        $x=$_|Format-PiPath
        "[true, ""$x"", """", """"]"
    }))
    $IntegrationDefinition = 
    "var P = new ImageIntegration;
    P.inputHints = `"`";
    P.combination = ImageIntegration.prototype.Average;
    P.weightMode = ImageIntegration.prototype.KeywordWeight;
    P.weightKeyword = ""SSWEIGHT"";
    P.weightScale = ImageIntegration.prototype.WeightScale_IKSS;
    P.ignoreNoiseKeywords = false;
    P.normalization = ImageIntegration.prototype.AdditiveWithScaling;
    P.rejection = ImageIntegration.prototype.$Rejection;
    P.rejectionNormalization = ImageIntegration.prototype.Scale;
    P.minMaxLow = 1;
    P.minMaxHigh = 1;
    P.pcClipLow = 0.200;
    P.pcClipHigh = 0.100;
    P.sigmaLow = 4.000;
    P.sigmaHigh = 2.000;
    P.winsorizationCutoff = 5.000;
    P.linearFitLow = $LinearFitLow;
    P.linearFitHigh = $LinearFitHigh;
    P.esdOutliersFraction = 0.30;
    P.esdAlpha = 0.05;
    P.ccdGain = 1.00;
    P.ccdReadNoise = 10.00;
    P.ccdScaleNoise = 0.00;
    P.clipLow = true;
    P.clipHigh = true;
    P.rangeClipLow = false;
    P.rangeLow = 0.000000;
    P.rangeClipHigh = false;
    P.rangeHigh = 0.980000;
    P.mapRangeRejection = true;
    P.reportRangeRejection = false;
    P.largeScaleClipLow = false;
    P.largeScaleClipLowProtectedLayers = 2;
    P.largeScaleClipLowGrowth = 2;
    P.largeScaleClipHigh = false;
    P.largeScaleClipHighProtectedLayers = 2;
    P.largeScaleClipHighGrowth = 2;
    P.generate64BitResult = false;
    P.generateRejectionMaps = true;
    P.generateIntegratedImage = true;
    P.generateDrizzleData = false;
    P.closePreviousImages = false;
    P.bufferSizeMB = 16;
    P.stackSizeMB = 1024;
    P.autoMemorySize = false;
    P.autoMemoryLimit = 0.75;
    P.useROI = false;
    P.roiX0 = 2376;
    P.roiY0 = 765;
    P.roiX1 = 3036;
    P.roiY1 = 1293;
    P.useCache = true;
    P.evaluateNoise = true;
    P.mrsMinDataFraction = 0.010;
    P.subtractPedestals = true;
    P.truncateOnOutOfRange = false;
    P.noGUIMessages = true;
    P.useFileThreads = true;
    P.fileThreadOverload = 1.00;
    P.images= [`r`n     $ImageDefinition`r`n   ];
    P.launch();
    P.executeGlobal();"
    $executionScript = New-TemporaryFile
    $executionScript = Rename-Item ($executionScript.FullName) ($executionScript.FullName+".js") -PassThru
    try {
        $IntegrationDefinition|Out-File -FilePath $executionScript -Force 
        Invoke-PIIntegrationScript `
            -path $executionScript `
            -PixInsightSlot $PixInsightSlot `
            -OutputFile $OutputFile `
            -KeepOpen:$KeepOpen
    }
    finally {
        Remove-Item $executionScript -Force
    }
} 