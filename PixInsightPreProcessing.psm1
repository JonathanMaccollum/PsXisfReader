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
            Wait-Event -Timeout 1            
        }
        while((Get-PixInsightInstance -PixInsightSlot $PixInsightSlot))
        {
            Write-Verbose "Waiting for completion of slot $PixInsightSlot"
            Wait-Event -Timeout 1            
        }
        Wait-Event -Timeout 2
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
    if($Path) {
        $Path.FullName.Replace('\','/')
    }
    else{
        $Directory.FullName.Replace('\','/')
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
save integration -p=$output --nodialog --nomessages --noverify"
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
        [Parameter(Mandatory=$false)][Switch]$KeepOpen
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
        $images | Foreach-Object {Move-Item -Path $_.Path -Destination "$targetDirectory\$flatDate"}
    }    
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
        $flats | Foreach-Object {Move-Item -Path $_.Path -Destination "$targetDirectory\$flatDate"}
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
    $calibratedFileName = $Path.Name.Substring($Path.Name.Length-$input.Extension.Length)+"_c.xisf"
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
        [Parameter(Mandatory)][System.IO.DirectoryInfo]$DropoffLocation,
        [Parameter(Mandatory)][System.IO.DirectoryInfo]$ArchiveDirectory,

        [Parameter(Mandatory)][String]$Filter,
        [Parameter(Mandatory)][int]$FocalLength,
        [Parameter(Mandatory)][int]$Exposure,
        [Parameter(Mandatory)][int]$Gain,
        [Parameter(Mandatory)][int]$Offset,

        [Parameter(Mandatory)][System.IO.FileInfo]$MasterDark,
        [Parameter(Mandatory)][System.IO.FileInfo]$MasterFlat,
        [Parameter(Mandatory)][System.IO.DirectoryInfo]$OutputPath,

        [Parameter(Mandatory)][int]$PixInsightSlot,
        [System.IO.DirectoryInfo[]]$AdditionalSearchPaths
    )

    Get-ChildItem $DropoffLocation "*.xisf" |
    foreach-object { Get-XisfFitsStats -Path $_ } |
    where-object ImageType -eq "Light" |
    where-object Filter -eq $Filter |
    where-object Exposure -eq $Exposure |
    where-object CCDGain -eq $CCDGain |
    where-object Offset -eq $Offset |
    where-object FocalLength -eq $FocalLength |
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
            Write-Host "Calibrating $($toCalibrate.Count) Light Frames for target $object"
            Invoke-PiLightCalibration `
                -Images $toCalibrate `
                -MasterDark $MasterDark `
                -MasterFlat $MasterFlat `
                -OutputPath $OutputPath `
                -PixInsightSlot $PixInsightSlot
        }

        $objectNameParts = $object.Split(' ')
        $archive = Join-Path $ArchiveDirectory "$($FocalLength)mm" -AdditionalChildPath $object
        if(-not (test-path $archive)){
            $archive = Join-Path $ArchiveDirectory "$($FocalLength)mm" -AdditionalChildPath ($objectNameParts[0] + ' '+ $objectNameParts[1])
        }
        if(-not (test-path $archive)){
            $archive = Join-Path $ArchiveDirectory "$($FocalLength)mm" -AdditionalChildPath ($objectNameParts[0])
        }
        $archive = Join-Path $archive "$Exposure.00"
        [System.IO.Directory]::CreateDirectory($archive) >> $null

        $_.Group | Foreach-Object {Move-Item -Path $_.Path -Destination $archive}
    }
}
