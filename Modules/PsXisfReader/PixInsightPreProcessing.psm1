[Reflection.Assembly]::Load("System.Text.RegularExpressions") >>$null
Function Wait-PixInsightInstance([int]$PixInsightSlot)
{
    [CmdletBinding]

    $x = Get-Process -Name "pixinsight" 
    if($PixInsightSlot) {
        $x = $x | where-object MainWindowTitle -eq "PixInsight ($PixInsightSlot)"
    }
    if($x) {
        Add-Member -MemberType NoteProperty -Name "PixInsightSlot" -Value $PixInsightSlot -InputObject $x
        while(-not $x.WaitForExit(1000)){
            Write-Debug "Waiting for PixInsight slot $PixInsightSlot to exit."
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
Function Start-PixInsight
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)][int]$PixInsightSlot
    )

    Push-Location "C:\Program Files\PixInsight\bin"
    try
    {
        $arguments = @(
            "--no-startup-scripts"
            "--no-startup-check-updates",
            "--no-startup-gui-messages",
            "--automation-mode",
            "--no-splash -n=$PixInsightSlot"
        );
        Start-Process ".\pixinsight.exe" -PassThru -ArgumentList $arguments >> $null
        while(-not (Get-PixInsightInstance -PixInsightSlot $PixInsightSlot))
        {
            Write-Debug "Waiting for PixInsight to start slot $PixInsightSlot"
            Wait-Event -Timeout 4            
        }
    }
    finally
    {
        Pop-Location
    }
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
        Write-Verbose "Waiting for PixInsight to start slot $PixInsightSlot"
        while(-not (Get-PixInsightInstance -PixInsightSlot $PixInsightSlot))
        {
            Write-Debug "Waiting for PixInsight to start slot $PixInsightSlot"
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
Function ConvertTo-XisfStfThumbnail{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][System.IO.FileInfo]$Path,
        [Parameter(Mandatory=$true)][System.IO.FileInfo]$OutputPath,
        [Parameter(Mandatory)]$PixInsightSlot,
        [Switch]$KeepOpen
    )
    if(-not $Path){
        Write-Error "Unable to locate file $Path"
    }
    else{
        $PiPath = $Path | Format-PiPath
        $PiOutputPath = $OutputPath | Format-PiPath
        $ScriptToRun = New-TemporaryFile
        $ScriptToRun = Rename-Item ($ScriptToRun.FullName) ($ScriptToRun.FullName+".scp") -PassThru
        Write-Debug "creating temporary file $ScriptToRun"
        $Template="
        open `"$PiPath`"
        ImageIdentifier -id=`"Thumbnail`"
        PixelMath -s=B,C,c -x=`"B = 0.25;C = -2.8;c = min( max( 0, med( Thumbnail ) + C*mdev( Thumbnail ) ), 1 );mtf( mtf( B, med( Thumbnail ) - c ), max( 0, (Thumbnail   - c)/~c ) ) `" 
        save Thumbnail -p=`"$PiOutputPath`" --nodialog --nomessages --noverify
        close --force Thumbnail
        "
        $Template|Out-File $ScriptToRun -Force -Append

        try {
            Write-Debug "Invoking Script ..."
            Invoke-PixInsightScript `
                -PixInsightSlot $PixInsightSlot `
                -Path $ScriptToRun `
                -KeepOpen:$KeepOpen
        }
        finally {
            Remove-Item $ScriptToRun -Force
        }    
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

Function Invoke-PIIntegrationBatchScript
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)][int]$PixInsightSlot, 
        [Parameter(Mandatory=$true)][Array]$ScriptsWithOutputFile,
        [Parameter(Mandatory=$false)][Switch]$KeepOpen
    )

    $ScriptToRun = New-TemporaryFile
    $ScriptToRun = Rename-Item ($ScriptToRun.FullName) ($ScriptToRun.FullName+".scp") -PassThru
    Write-Debug "creating temporary file $ScriptToRun"

    $intermediateScripts = $ScriptsWithOutputFile | foreach-object {
        $definition = $_.Script
        $OutputFile = [System.IO.FileInfo]$_.OutputFile

        $executionScript = New-TemporaryFile
        $executionScript = Rename-Item ($executionScript.FullName) ($executionScript.FullName+".js") -PassThru
        Write-Debug "creating temporary file $executionScript for $OutputFile"
        $definition | Out-File $executionScript -Force

        $scp = Format-PiPath -Path $executionScript
        $output=$OutputFile|Format-PiPath
        $Template="run -x $scp
    save integration -p=`"$output`" --nodialog --nomessages --noverify
    close integration
    "
        $Template|Out-File $ScriptToRun -Force -Append

        $executionScript
    }

    try {
        Write-Debug "Invoking Script ..."
        Invoke-PixInsightScript `
            -PixInsightSlot $PixInsightSlot `
            -Path $ScriptToRun `
            -KeepOpen:$KeepOpen
    }
    finally {
        $intermediateScripts | Foreach-Object {Remove-Item -Path $_ -Force}
        Remove-Item $ScriptToRun -Force
    }
}
Function Invoke-PIIntegrationScript
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)][int]$PixInsightSlot, 
        [Parameter(Mandatory=$true)][System.IO.FileInfo]$Path, 
        [Parameter(Mandatory=$true)][System.IO.FileInfo]$OutputFile,
        [Parameter(Mandatory=$false)][Switch]$KeepOpen,
        [Parameter(Mandatory=$false)][Switch]$GenerateThumbnail

    )
    $scp = $Path|Format-PiPath
    $output=$OutputFile|Format-PiPath
    $Template="run -x $scp
save integration -p=`"$output`" --nodialog --nomessages --noverify"
    if($GenerateThumbnail.IsPresent){
        $ThumbnailDir = 
            Join-Path `
                -Path $OutputFile.Directory.FullName `
                -ChildPath "Thumbnails"
        [System.IO.Directory]::CreateDirectory($ThumbnailDir) >> $null
        $thumbnail= Format-PiPath -Path (Join-Path `
                -Path $ThumbnailDir `
                -ChildPath ([System.IO.Path]::GetFileNameWithoutExtension($OutputFile)+".jpeg"))
            
        $Template += "
PixelMath -n+ -id=Thumbnail -s=B,C,c -x=`"B = 0.25;C = -2.8;c = min( max( 0, med( integration ) + C*mdev( integration  ) ), 1 );mtf( mtf( B, med( integration  ) - c ), max( 0, (integration  - c)/~c ) ) `"
save Thumbnail -p=`"$thumbnail`" --nodialog --nomessages --noverify"
    }
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
    P.subtractPedestals = false;
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
    P.rejection = ImageIntegration.prototype.Rejection_ESD;
    P.rejectionNormalization = ImageIntegration.prototype.EqualizeFluxes;
    P.linearFitLow = 5.000;
    P.linearFitHigh = 2.500;
    P.esdOutliersFraction = 0.50;
    P.esdAlpha = 0.40;
    P.esdLowRelaxation = 1.00;
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
    P.subtractPedestals = false;
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
Function Invoke-PiFlatCalibration
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)][int]$PixInsightSlot,
        [Parameter(Mandatory=$true)][System.IO.FileInfo[]]$Images,
        [Parameter(Mandatory=$true,ParameterSetName="DarkOnly")][System.IO.FileInfo]$MasterDark,
        [Parameter(Mandatory=$true,ParameterSetName="BiasOnly")][System.IO.FileInfo]$MasterBias,
        [Parameter(Mandatory=$true)][System.IO.DirectoryInfo]$OutputPath,
        [Parameter(Mandatory=$false)][Switch]$KeepOpen
    )
    if($MasterDark){
        $masterDarkPath = Get-Item $MasterDark | Format-PiPath
    }
    else{
        $masterDarkPath=""
    }
    if($MasterBias){
        $masterBiasPath = Get-Item $MasterBias | Format-PiPath
    }   
    else{
        $masterBiasPath=""
    }
    
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
P.masterBiasEnabled = $((-not [string]::IsNullOrWhiteSpace($masterBiasPath)).ToString().ToLower());
P.masterBiasPath = `"$masterBiasPath`";
P.masterDarkEnabled = $((-not [string]::IsNullOrWhiteSpace($masterDarkPath)).ToString().ToLower());
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
    P.launch();
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
        [Parameter(Mandatory=$false)][System.IO.FileInfo]$MasterBias,
        [Parameter(Mandatory=$false)][System.IO.FileInfo]$MasterDark,
        [Parameter(Mandatory=$false)][Switch]$CalibrateDark,
        [Parameter(Mandatory=$false)][Switch]$OptimizeDark,
        [Parameter(Mandatory=$false)][System.IO.FileInfo]$MasterFlat,
        [Parameter(Mandatory=$true)][System.IO.DirectoryInfo]$OutputPath,
        [Parameter(Mandatory=$false)][Switch]$KeepOpen,
        [Parameter()][int]$OutputPedestal=0,
        [Parameter()][Switch]$EnableCFA
    )

    if($MasterBias -and (-not (Test-Path $MasterBias))){
        write-error "A master Bias was specified, but could not be located: $MasterBias"
        return
    }
    if($MasterDark -and (-not (Test-Path $MasterDark))){
        write-error "A master Dark was specified, but could not be located: $MasterDark"
        return
    }
    if($MasterFlat -and (-not (Test-Path $MasterFlat))){
        write-error "A master Flat was specified, but could not be located: $MasterFlat"
        return
    }

    if($MasterBias -and (Test-Path $MasterBias)){
        $masterBiasPath = Get-Item $MasterBias | Format-PiPath
        $masterBiasEnabled = "true"
    }
    else{
        $masterBiasPath=""
        $masterBiasEnabled = "false"
    }
    
    if($MasterDark -and (Test-Path $MasterDark)){
        $masterDarkPath = Get-Item $MasterDark | Format-PiPath
        $masterDarkEnabled = "true"
    }
    else{
        $masterDarkPath=""
        $masterDarkEnabled = "false"
    }
    
    if($MasterFlat -and (Test-Path $MasterFlat)){
        $masterFlatPath = Get-Item $MasterFlat | Format-PiPath
        $masterFlatEnabled = "true"
    }
    else{
        $masterFlatPath=""
        $masterFlatEnabled = "false"
    }
    if($EnableCFA){
        $EnableCFAFlag="true"
    }
    else{
        $EnableCFAFlag="false"
    }

    if(-not ($OutputPath.Exists)){
        $OutputPath.Create()
    }
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
P.outputHints = `"properties fits-keywords compression-codec lz4+sh no-embedded-data no-resolution`";
P.masterBiasEnabled = $masterBiasEnabled;
P.masterBiasPath = `"$masterBiasPath`";
P.masterDarkEnabled = $masterDarkEnabled;
P.masterDarkPath = `"$masterDarkPath`";
P.masterFlatEnabled = $masterFlatEnabled;
P.masterFlatPath = `"$masterFlatPath`";
P.calibrateBias = false;
P.calibrateDark = $($CalibrateDark.IsPresent.ToString().ToLower());
P.calibrateFlat = false;
P.optimizeDarks = $($OptimizeDark.IsPresent.ToString().ToLower());
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
P.enableCFA = $EnableCFAFlag;
    P.targetFrames= [`r`n     $ImageDefinition`r`n   ];
    P.launch();
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

Function Invoke-PiCosmeticCorrection
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)][int]$PixInsightSlot,
        [Parameter(Mandatory=$true)][System.IO.FileInfo[]]$Images,
        [Parameter(Mandatory=$false)][System.IO.FileInfo]$MasterDark,
        [Parameter(Mandatory=$true)][System.IO.DirectoryInfo]$OutputPath,
        [Parameter(Mandatory=$false)][Switch]$KeepOpen,
        [Parameter()][Switch]$CFAImages,
        [Parameter(Mandatory=$false)][double]$HotDarkLevel=0.666666,

        [Parameter()][Switch]$UseAutoHot,
        [Parameter(Mandatory=$false)][double]$HotAutoSigma=3.0,
        [Parameter()][Switch]$UseAutoCold,
        [Parameter(Mandatory=$false)][double]$ColdAutoSigma=3.0
    )
    if($MasterDark){
        $masterDarkPath = Get-Item $MasterDark | Format-PiPath
        $useMasterDark="true"
    }
    else{
        $masterDarkPath=""
        $useMasterDark="false"
    }
    $outputDirectory = Get-Item ($OutputPath.FullName) | Format-PiPath
    $ImageDefinition = [string]::Join("`r`n   , ",
    ($Images | ForEach-Object {
        $x=$_|Format-PiPath
        "[true, ""$x""]"
    }))
    $useAutoDetect = (($UseAutoHot.IsPresent) -or ($UseAutoCold.IsPresent))
    $IntegrationDefinition = 
    "var P = new CosmeticCorrection;
P.useMasterDark = true;
P.masterDarkPath = `"$masterDarkPath`";
P.outputDir = `"$outputDirectory`";
P.outputExtension = `".xisf`";
P.prefix = `"`";
P.postfix = `"_cc`";
P.overwrite = false;
P.amount = 1.00;
P.cfa = $($CFAImages.IsPresent.ToString().ToLower());
P.useMasterDark = $useMasterDark;
P.hotDarkCheck = $useMasterDark;
P.hotDarkLevel = $HotDarkLevel;
P.coldDarkCheck = false;
P.coldDarkLevel = 0.0000000;
P.useAutoDetect = $($useAutoDetect.ToString().ToLower());
P.hotAutoCheck = $($UseAutoHot.IsPresent.ToString().ToLower());
P.hotAutoValue = $HotAutoSigma;
P.coldAutoCheck = $($UseAutoCold.IsPresent.ToString().ToLower());
P.coldAutoValue = $ColdAutoSigma;
P.useDefectList = false;
P.defects = [ // defectEnabled, defectIsRow, defectAddress, defectIsRange, defectBegin, defectEnd
];


    P.targetFrames= [`r`n     $ImageDefinition`r`n   ];
    P.launch();
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

Function Invoke-PiDebayer
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)][int]$PixInsightSlot,
        [Parameter(Mandatory=$true)][System.IO.FileInfo[]]$Images,
        [Parameter(Mandatory=$true)][System.IO.DirectoryInfo]$OutputPath,
        [Parameter(Mandatory=$false)][Switch]$KeepOpen,
        [Parameter(Mandatory=$false)][string]$CfaPattern = "Auto",
        [Parameter()][Switch]$CFAImages,
        [Parameter()][Switch]$Compress
    )
    if($Compress){
        $outputHints="properties fits-keywords compression-codec lz4+sh"
    }
    else{
        $outputHints=""
    }
    $outputDirectory = Get-Item ($OutputPath.FullName) | Format-PiPath
    $ImageDefinition = [string]::Join("`r`n   , ",
    ($Images | ForEach-Object {
        $x=$_|Format-PiPath
        "[true, ""$x""]"
    }))
    $IntegrationDefinition = 
    "
    var P = new Debayer;
    P.cfaPattern = Debayer.prototype.$($CfaPattern);
    P.debayerMethod = Debayer.prototype.VNG;
    P.fbddNoiseReduction = 0;
    P.evaluateNoise = true;
    P.noiseEvaluationAlgorithm = Debayer.prototype.NoiseEvaluation_MRS;
    P.showImages = true;
    P.cfaSourceFilePath = `"`";
    P.noGUIMessages = true;
    P.inputHints = `"raw cfa`";
    P.outputHints = `"$outputHints`";
    P.outputDirectory = `"$outputDirectory`";
    P.outputExtension = `".xisf`";
    P.outputprefix = `"`";
    P.outputPostfix = `"_d`";
    P.overwriteExistingFiles = false;
    P.onError = Debayer.prototype.OnError_Continue;
    P.useFileThreads = true;
    P.fileThreadOverload = 1.00;
    P.maxFileReadThreads = 1;
    P.maxFileWriteThreads = 1;
    P.targetItems= [`r`n     $ImageDefinition`r`n   ];
    P.launch();
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

Function Invoke-BiasFrameSorting
{
    [CmdletBinding()]
    param (
        [System.IO.DirectoryInfo]$DropoffLocation,
        [System.IO.DirectoryInfo]$ArchiveDirectory,
        [int]$PixInsightSlot,
        [Switch]$KeepOpen,
        [Switch]$DoNotArchive
    )

    Get-ChildItem $DropoffLocation "*.xisf" |
    foreach-object { Get-XisfFitsStats -Path $_ } |
    where-object ImageType -eq "BIAS" |
    group-object Instrument,Exposure,Gain,Offset,Geometry |
    foreach-object {
        $images=$_.Group
        
        $instrument=$images[0].Instrument.Trim().Trim("'")
        $exposure=$images[0].Exposure.ToString().Trim()
        $gain=($images[0].Gain)
        $offset=($images[0].Offset)
        $date = ([DateTime]$images[0].LocalDate).Date.ToString('yyyyMMdd')
        $targetDirectory = "$ArchiveDirectory\BiasLibrary\$instrument"
        [System.IO.Directory]::CreateDirectory($targetDirectory)>>$null
        $masterDark = "$targetDirectory\$($date).MasterBias.Gain.$($gain).Offset.$($offset).$($images.Count)x$($exposure)s.xisf"
        if($images.Count -lt 10){
            Write-Warning "Skipping $($images.Count) Bias for $instrument Gain $gain Offset $offset duration $($exposure)s: Need more than 10 subs for making a bias master."
        }
        elseif(-not (Test-Path $masterDark)) {
            Write-Host "Integrating $($images.Count) Bias for $instrument Gain $gain Offset $offset duration $($exposure)s"
            $darkFlats = $images | foreach-object {$_.Path}
            Invoke-PiDarkIntegration `
                -Images $darkFlats `
                -PixInsightSlot $PixInsightSlot `
                -OutputFile $masterDark `
                -KeepOpen:$KeepOpen
                    
            if(-not $DoNotArchive){
                $destinationDirectory=join-path $targetDirectory $date
                [System.IO.Directory]::CreateDirectory($destinationDirectory)>>$null
                $images | Foreach-Object {
                    Move-Item -Path $_.Path -Destination $destinationDirectory
                }
            }
        }
    }
}

Function Invoke-DarkFrameSorting
{
    [CmdletBinding()]
    param (
        [System.IO.DirectoryInfo]$DropoffLocation,
        [System.IO.DirectoryInfo]$ArchiveDirectory,
        [int]$PixInsightSlot,
        [Switch]$KeepOpen,
        [Switch]$DoNotArchive
    )

    Get-ChildItem $DropoffLocation "*.xisf" |
    foreach-object { Get-XisfFitsStats -Path $_ } |
    where-object ImageType -eq "DARK" |
    group-object Instrument,Exposure,SetTemp,Gain,Offset,Geometry |
    foreach-object {
        $images=$_.Group
        
        $instrument=$images[0].Instrument.Trim().Trim("'")
        $exposure=$images[0].Exposure.ToString().Trim()
        $setTemp=($images[0].SetTemp)
        $gain=($images[0].Gain)
        $offset=($images[0].Offset)
        $date = ([DateTime]$images[0].LocalDate).Date.ToString('yyyyMMdd')
        $targetDirectory = "$ArchiveDirectory\DarkLibrary\$instrument"
        [System.IO.Directory]::CreateDirectory($targetDirectory)>>$null
        $masterDark = "$targetDirectory\$($date).MasterDark.Gain.$($gain).Offset.$($offset).$($setTemp)C.$($images.Count)x$($exposure)s.xisf"
        if($images.Count -lt 10){
            Write-Warning "Skipping $($images.Count) Darks for $instrument Gain $gain Offset $offset at $setTemp duration $($exposure)s: Need more than 10 subs for making a dark master."
        }
        elseif(-not (Test-Path $masterDark)) {
            Write-Host "Integrating $($images.Count) Darks for $instrument Gain $gain Offset $offset at $setTemp duration $($exposure)s"
            $darkFlats = $images | foreach-object {$_.Path}
            Invoke-PiDarkIntegration `
                -Images $darkFlats `
                -PixInsightSlot $PixInsightSlot `
                -OutputFile $masterDark `
                -KeepOpen:$KeepOpen
                    
            if(-not $DoNotArchive){
                $destinationDirectory=join-path $targetDirectory $date
                [System.IO.Directory]::CreateDirectory($destinationDirectory)>>$null
                $images | Foreach-Object {
                    Move-Item -Path $_.Path -Destination $destinationDirectory
                }
            }
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
        $focalLength=($images[0].FocalLength.ToString().TrimEnd('mm'))+'mm'
        $flatDate = ([DateTime]$images[0].LocalDate).Date.ToString('yyyyMMdd')
        $targetDirectory = "$ArchiveDirectory\$focalLength\Flats"
        [System.IO.Directory]::CreateDirectory($targetDirectory)>>$null
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
        Get-XisfFitsStats
    
}
Function Skip-PathsContainingTokens{
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline=$true,Mandatory)][System.IO.FileInfo]$Path,
        [Parameter(ValueFromPipeline=$false)][string[]]$PathTokensToIgnore
    )
    process{
        if(-not $PathTokensToIgnore){
            Write-Output $Path
        }
        else{
            $hasToken=$false
            foreach( $x in $PathTokensToIgnore) {
                if($Path.FullName.ToLower().Contains($x.ToLower())){
                    $hasToken=$true
                    break;
                }
            }
            if(-not $hasToken){
                Write-Output $Path
            }
        }    
    }
}
Function Get-XisfLightFrames{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)][System.IO.DirectoryInfo]$Path,
        [Parameter()][Switch]$Recurse,
        [Parameter()][Switch]$UseCache,
        [Parameter()][Switch]$UseErrorCache,
        [Parameter()][Switch]$SkipOnError,
        [Parameter()][string[]]$PathTokensToIgnore,
        [Parameter()][Switch]$ShowProgress,
        [Parameter()][Switch]$TruncateFilterBandpass
    )
    begin
    {        
        if($UseCache.IsPresent){
            $PathToCache= Join-Path ($Path.FullName) "Cache.clixml"
            if(Test-Path $PathToCache){
                $cache = Import-Clixml -Path $PathToCache
            }
            else {
                $cache = new-object hashtable
            }
            $entriesBefore=$cache.Count
        }
        if($UseErrorCache.IsPresent){
            $PathToErrorCache= Join-Path ($Path.FullName) "ErrorCache.clixml"
            if(Test-Path $PathToErrorCache){
                $errorCache = Import-Clixml -Path $PathToErrorCache
            }
            else {
                $errorCache = new-object hashtable
            }
            $errorEntriesBefore=$errorCache.Count
        }
    }
    process
    {
        Get-ChildItem -Path $Path -File -Filter *.xisf |
        Skip-PathsContainingTokens -PathTokensToIgnore $PathTokensToIgnore |
        foreach-object {
            $file=$_
            try{
                $file|Get-XisfFitsStats -Cache:$cache -ErrorCache:$errorCache -TruncateFilterBandpass:$TruncateFilterBandpass |where-object ImageType -eq "LIGHT"
            }
            catch{
                if(-not ($SkipOnError.IsPresent)){
                    throw;
                }
            }
        }
        if($UseCache -and ($entriesBefore -ne $cache.Count)){
            $cache|Export-Clixml -Path $PathToCache -Force -Depth 3
        }
        if($UseErrorCache -and ($errorEntriesBefore -ne $errorCache.Count)){
            $errorCache|Export-Clixml -Path $PathToErrorCache -Force -Depth 3
        }
        if($Recurse.IsPresent)
        {
            $directories = Get-ChildItem -Path $Path -Directory
            if($directories){
                $i=0;
                $maxi = $directories.Length
                $directories| ForEach-Object {
                    if($ShowProgress){
                        Write-Progress -PercentComplete (100.0*$i/$maxi) -Activity "Scanning Directories $($Path.Name)" -Status ($_.Name) -Id $_.GetHashCode() -ParentId $Path.GetHashCode()
                    }
                    Get-XisfLightFrames `
                        -Recurse `
                        -Path $_ `
                        -UseCache:$UseCache `
                        -UseErrorCache:$UseErrorCache `
                        -SkipOnError:$SkipOnError `
                        -PathTokensToIgnore:$PathTokensToIgnore `
                        -TruncateFilterBandpass:$TruncateFilterBandpass `
                        -ShowProgress:$ShowProgress
                    $i+=1
                    if($i -eq $maxi){
                        if($ShowProgress){
                            Write-Progress -PercentComplete 100 -Activity "Scanning Directories" -Status ($_.Name) -Id $_.GetHashCode() -ParentId $Path.GetHashCode()
                        }
                    }
                }
            }
        }
    }
}
Function Get-XisfDarkFrames{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)][System.IO.DirectoryInfo]$Path,
        [Parameter()][Switch]$Recurse,
        [Parameter()][Switch]$UseCache,
        [Parameter()][Switch]$SkipOnError
    )
    begin
    {        
        if($UseCache.IsPresent){
            $PathToCache= Join-Path ($Path.FullName) "Cache.clixml"
            if(Test-Path $PathToCache){
                $cache = Import-Clixml -Path $PathToCache
            }
            else {
                $cache = new-object hashtable
            }
            $entriesBefore=$cache.Count    
        }
    }
    process
    {
        Get-ChildItem -Path $Path -File -Filter *.xisf |
        foreach-object {
            $file=$_
            try{
                $file|Get-XisfFitsStats -Cache:$cache |where-object ImageType -eq "DARK"
            }
            catch{
                if(-not ($SkipOnError.IsPresent)){
                    throw;
                }
            }
        }
        if($UseCache -and ($cache.Count) -ne $entriesBefore){
            $cache|Export-Clixml -Path $PathToCache -Force
        }
        if($Recurse.IsPresent)
        {
            Get-ChildItem -Path $Path -Directory | ForEach-Object {
                Get-XisfDarkFrames -Recurse -Path $_ -UseCache:$UseCache -SkipOnError:$SkipOnError
            }
        }
    }
}
Function Invoke-FlatFrameSorting
{
    [CmdletBinding()]
    param (
        [System.IO.DirectoryInfo]$DropoffLocation,
        [System.IO.DirectoryInfo]$ArchiveDirectory,
        [System.IO.DirectoryInfo]$CalibratedFlatsOutput,
        [int]$PixInsightSlot,
        [Switch]$KeepOpen,
        [Switch]$WhenNoMatchingDarkFlatPresentUseMostRecentDarkFlat,
        [Switch]$UseBias,
        [Array]$BiasLibraryFiles
    )
    $toProcess = 
        Get-ChildItem $DropoffLocation "*.xisf" |
        foreach-object { Get-XisfFitsStats -Path $_ } |
        where-object ImageType -eq "FLAT"
    
    
    $toProcess|
        group-object Filter,FocalLength,Gain,Offset,ObsDateMinus12hr |
        foreach-object {
            $flats=$_.Group
            $gain=$flats[0].Gain
            $offset=$flats[0].Offset
            $filter=$flats[0].Filter.Trim()
            $focalLength=($flats[0].FocalLength.ToString().TrimEnd('mm'))+'mm'
            $flatDate = ([DateTime]$flats[0].ObsDateMinus12hr).Date.ToString('yyyyMMdd')

            Write-Host "Processing $($flats.Count) Flats for filter '$filter' at $focalLength"

            $targetDirectory = "$ArchiveDirectory\$focalLength\Flats"
            $masterCalibratedFlat = "$targetDirectory\$flatDate.MasterFlatCal.$filter.xisf"

            if(-not (Test-Path $masterCalibratedFlat)) {
                if($UseBias.IsPresent){
                    $masterBias = $BiasLibraryFiles |
                        where-object Gain -eq $gain |
                        where-object Offset -eq $offset |
                        sort-object ObsDate -Descending |
                        foreach-object {get-item $_.Path} |
                        select-object -First 1
                    if(-not $masterBias){
                        write-error "No matching bias exists. A bias with gain $gain and offset $offset was not found."
                    }
                }
                else{
                    $masterDark = "$targetDirectory\$flatDate.MasterDarkFlat.$filter.xisf"
                    if((-not (Test-Path $masterDark)) -and $WhenNoMatchingDarkFlatPresentUseMostRecentDarkFlat){
                        $masterDark = get-childitem $targetDirectory -Filter "*.MasterDarkFlat.$filter.xisf" -file |
                            sort-object LastWriteTime -Descending |
                            select-object -First 1
                        Write-Warning "No matching dark flat exists: calibrating flat frame with the most recent master dark flat $($masterDark.Name)"
                    }
                    if(-not (Test-Path $masterDark)){
                        write-error "Unable to locate master dark to calibrate filter $filter. "
                    }
                }


                $flats | foreach-object{
                    $x = $_.Path
                    $output = Join-Path ($CalibratedFlatsOutput.FullName) ($x.Name.Replace($x.Extension,"")+"_c.xisf")
                    Add-Member -InputObject $_ -Name "CalibratedFlat" -MemberType NoteProperty -Value $output -Force
                }
                $toCalibrate = $flats | where-object {-not (Test-Path $_.CalibratedFlat)} | foreach-object {$_.Path}
                if($toCalibrate -and $masterDark -and (Test-Path $masterDark)){
                    Invoke-PiFlatCalibration `
                        -Images $toCalibrate `
                        -MasterDark $masterDark `
                        -OutputPath $CalibratedFlatsOutput `
                        -PixInsightSlot $PixInsightSlot `
                        -KeepOpen:$KeepOpen
                }
                elseif($toCalibrate -and $masterBias -and (Test-Path $masterBias)){
                    Invoke-PiFlatCalibration `
                        -Images $toCalibrate `
                        -MasterBias $masterBias `
                        -OutputPath $CalibratedFlatsOutput `
                        -PixInsightSlot $PixInsightSlot `
                        -KeepOpen:$KeepOpen
                }
                $calibratedFlats = $flats|foreach-object {Get-Item $_.CalibratedFlat}
                if($flats) {
                    Invoke-PiFlatIntegration `
                        -Images $calibratedFlats `
                        -PixInsightSlot $PixInsightSlot `
                        -OutputFile $masterCalibratedFlat `
                        -KeepOpen:$KeepOpen
                }
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
            where-object { test-path $_ } |
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
        [Parameter()][System.IO.DirectoryInfo]$ArchiveDirectory,

        [Parameter(Mandatory,ParameterSetName="ByDirectory")][String]$Filter,
        [Parameter(Mandatory,ParameterSetName="ByDirectory")][int]$FocalLength,
        [Parameter(Mandatory,ParameterSetName="ByDirectory")][int]$Exposure,
        [Parameter(Mandatory,ParameterSetName="ByDirectory")][int]$Gain,
        [Parameter(Mandatory,ParameterSetName="ByDirectory")][int]$Offset,
        [Parameter(Mandatory,ParameterSetName="ByDirectory")][int]$SetTemp,
        [Parameter()][System.IO.FileInfo]$MasterBias,
        [Parameter()][System.IO.FileInfo]$MasterDark,
        [Parameter()][Switch]$CalibrateDark,
        [Parameter()][Switch]$OptimizeDark,
        [Parameter()][System.IO.FileInfo]$MasterFlat,
        [Parameter(Mandatory)][System.IO.DirectoryInfo]$OutputPath,

        [Parameter(Mandatory)][int]$PixInsightSlot,
        [System.IO.DirectoryInfo[]]$AdditionalSearchPaths,
        [Parameter(Mandatory=$false)][int]$OutputPedestal = 0,
        [Parameter(Mandatory=$false)][Switch]$KeepOpen,
        [Parameter(Mandatory=$false)][ScriptBlock]$AfterImageCalibrated,
        [Parameter(Mandatory=$false)][ScriptBlock]$AfterImagesCalibrated,

        [Parameter(Mandatory=$false)][Switch]$EnableCFA,
        [Parameter(Mandatory=$false)][Switch]$DoNotArchive,
        [Parameter(Mandatory=$false)][Switch]$TruncateFilterBandpass
    )
    if($null -eq $XisfStats)
    {
        $XisfStats = Get-ChildItem $DropoffLocation "*.xisf" -File |
            Get-XisfFitsStats -TruncateFilterBandpass:$TruncateFilterBandpass |
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
        $OutputDirPerObject = Join-Path ($OutputPath.FullName) $object
        [System.IO.Directory]::CreateDirectory($OutputDirPerObject) >> $null
        $toCalibrate = $_.Group |
            foreach-object {
                $x = $_
                $calibrated = Get-CalibrationFile -Path $_.Path -CalibratedPath $OutputDirPerObject -AdditionalSearchPaths $AdditionalSearchPaths
                Add-Member -InputObject $x -Name "Calibrated" -MemberType NoteProperty -Value $calibrated -Force
                $x
            } |
            where-object {-not (Test-Path $_.Calibrated)} | foreach-object {$_.Path}
        if($toCalibrate){
            Write-Host "Calibrating $($toCalibrate.Count)x$($Exposure)s $Filter Frames for target $object"
            Invoke-PiLightCalibration `
                -Images $toCalibrate `
                -MasterBias $MasterBias `
                -MasterDark $MasterDark `
                -MasterFlat $MasterFlat `
                -CalibrateDark:$CalibrateDark `
                -OptimizeDark:$OptimizeDark `
                -OutputPath $OutputDirPerObject `
                -OutputPedestal $OutputPedestal `
                -PixInsightSlot $PixInsightSlot `
                -EnableCFA:$EnableCFA.IsPresent `
                -KeepOpen:$KeepOpen
        }

        if(-not $DoNotArchive.IsPresent){
            $archive = Join-Path $ArchiveDirectory "$($FocalLength)mm" 
            $archiveObjectWithoutSpaces = Join-Path $archive ($object.Replace(' ',''))
            if(test-path $archiveObjectWithoutSpaces){
                $archive = $archiveObjectWithoutSpaces
            }else{
                $archive = join-path $archive $object.Trim()
            }
            $archive = join-path $archive "$($exposure).00s"
            [System.IO.Directory]::CreateDirectory($archive) >> $null
        }

        $_.Group | Foreach-Object {
            if($AfterImageCalibrated){
                $AfterImageCalibrated.Invoke($_)
            }
        }
        if($AfterImagesCalibrated){
            $AfterImagesCalibrated.Invoke($_.Group)
        }
        
        if(-not $DoNotArchive.IsPresent){
            $_.Group | Foreach-Object {
                Move-Item -Path $_.Path -Destination $archive
            }
        }
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
        [Parameter(Mandatory=$false)][int]$DetectionScales=5,
        [Parameter(Mandatory=$false)][decimal]$ClampingThreshold=0.3,
        
        [ValidateSet("Auto","Lanczos3","Lanczos4","Bilinear","CubicBSplineFilter","MitchellNetravaliFilter","CatmullRomSplineFilter")]
        [Parameter(Mandatory=$false)][String]$Interpolation="Auto"
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
P.minStructureSize = 0;
P.sensitivity = 0.50;
P.peakResponse = 0.50;
P.brightThreshold = 3.00;
P.maxStarDistortion = 0.60;
P.allowClusteredSources = false;
P.localMaximaDetectionLimit = 0.75;
P.upperLimit = 1.000;
P.invert = false;
P.distortionModel = `"`";
P.undistortedReference = false;
P.rigidTransformations = false;
P.distortionCorrection = false;
P.distortionMaxIterations = 20;
P.distortionMatcherExpansion = 1.00;
P.splineOrder = 2;
P.splineSmoothness = 0.005;
P.splineOutlierDetectionRadius = 160;
P.splineOutlierDetectionMinThreshold = 4.0;
P.splineOutlierDetectionSigma = 5.0;
P.matcherTolerance = 0.0500;
P.ransacTolerance = 1.9000;
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
P.inheritAstrometricSolution = false;
P.frameAdaptation = false;
P.randomizeMosaic = false;
P.noGUIMessages = true;
P.pixelInterpolation = StarAlignment.prototype.$($Interpolation);
P.clampingThreshold = $ClampingThreshold;
P.outputExtension = `".xisf`";
P.outputPrefix = `"`";
P.outputPostfix = `"_r`";
P.maskPostfix = `"_m`";
P.distortionMapPostfix = `"_dm`";
P.outputSampleFormat = StarAlignment.prototype.SameAsTarget;
P.overwriteExistingFiles = false;
P.onError = StarAlignment.prototype.Continue;
P.useFileThreads = true;
P.fileThreadOverload = 1.00;
P.maxFileReadThreads = 0;
P.maxFileWriteThreads = 0;
P.outputDirectory = `"$outputDirectory`";
P.targets= [`r`n     $ImageDefinition`r`n   ];
P.memoryLoadControl = true;
P.memoryLoadLimit = 0.85;
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
        [Parameter(Mandatory=$false)][string]$Rejection = "LinearFit", #Rejection_ESD
        [Parameter(Mandatory=$false)][string]$Normalization = "AdditiveWithScaling", #AdaptiveNormalization
        [Parameter(Mandatory=$false)][string]$RejectionNormalization = "Scale", #AdaptiveRejectionNormalization
        [Parameter(Mandatory=$false)][Switch]$GenerateDrizzleData,
        [Parameter(Mandatory=$false)][string]$WeightMode = "KeywordWeight",
        [Parameter(Mandatory=$false)][string]$WeightKeyword = "SSWEIGHT",
        [Parameter(Mandatory=$false)][string]$Combination = "Average",
        [Parameter(Mandatory=$false)][bool]$GenerateRejectionMaps = $true,
        [Parameter(Mandatory=$false)][bool]$EvaluateNoise = $true,
        [Switch]$OutputDefinitionOnly,
        [Switch]$GenerateThumbnail
    )
    $ImageDefinition = [string]::Join("`r`n   , ",
    ($Images | ForEach-Object {
        $x=$_|Format-PiPath
        $y=if($GenerateDrizzleData.IsPresent){$x.Replace(".xisf",".xdrz")}else{""}
        "[true, ""$x"", ""$y"", """"]"
    }))
    if([string]::IsNullOrWhiteSpace($WeightKeyword)){
        $WeightMode="PSFSignalWeight"
    }
    if([string]::IsNullOrWhiteSpace($Combination)){
        $Combination="Average"
    }
    $IntegrationDefinition = 
    "var P = new ImageIntegration;
    P.inputHints = `"`";
    P.combination = ImageIntegration.prototype.$Combination;
    P.weightMode = ImageIntegration.prototype.$WeightMode;
    P.weightKeyword = ""$WeightKeyword"";
    P.weightScale = ImageIntegration.prototype.WeightScale_IKSS;
    P.ignoreNoiseKeywords = false;
    P.normalization = ImageIntegration.prototype.$Normalization;
    P.rejection = ImageIntegration.prototype.$Rejection;
    P.rejectionNormalization = ImageIntegration.prototype.$RejectionNormalization;
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
    P.rangeClipLow = true;
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
    P.generateRejectionMaps = $($GenerateRejectionMaps.ToString().ToLower());
    P.generateIntegratedImage = true;
    P.generateDrizzleData = $($GenerateDrizzleData.IsPresent.ToString().ToLower());
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
    P.evaluateNoise = $($EvaluateNoise.ToString().ToLower());
    P.mrsMinDataFraction = 0.010;
    P.subtractPedestals = false;
    P.truncateOnOutOfRange = false;
    P.noGUIMessages = true;
    P.useFileThreads = true;
    P.fileThreadOverload = 1.00;
    P.images= [`r`n     $ImageDefinition`r`n   ];
    P.launch();
    P.executeGlobal();"

    if($OutputDefinitionOnly){
        Write-Output -NoEnumerate $IntegrationDefinition
    }
    else{
        $executionScript = New-TemporaryFile
        $executionScript = Rename-Item ($executionScript.FullName) ($executionScript.FullName+".js") -PassThru
        try {
            $IntegrationDefinition|Out-File -FilePath $executionScript -Force 
            Invoke-PIIntegrationScript `
                -path $executionScript `
                -PixInsightSlot $PixInsightSlot `
                -OutputFile $OutputFile `
                -KeepOpen:$KeepOpen `
                -GenerateThumbnail:$GenerateThumbnail
        }
        finally {
            Remove-Item $executionScript -Force
        }
    }
} 

class XisfPreprocessingState {
    [XisfFileStats]$Stats
    [System.IO.FileInfo]$Path

    [string]$MasterBias
    [string]$MasterDark
    [string]$MasterFlat

    [System.IO.FileInfo]$Calibrated
    [System.IO.FileInfo]$Corrected
    [System.IO.FileInfo]$Debayered
    [System.IO.FileInfo]$Weighted
    [System.IO.FileInfo]$Aligned

    [bool]IsCalibrated() {
        return $this.Calibrated `
            -and $this.Calibrated.Exists
    }
    [bool]IsCorrected() {
        return $this.Corrected `
            -and $this.Corrected.Exists
    }
    [bool]IsDebayered() {
        return `
                 $this.Debayered `
            -and $this.Debayered.Exists
    }
    [bool]IsWeighted() {
        return `
                 $this.Weighted `
            -and $this.Weighted.Exists
    }
    [bool]IsAligned() {
        return `
                 $this.Aligned `
            -and $this.Aligned.Exists
    }
    [System.IO.FileInfo]GetDrizzleFile(){
        if($this.Aligned){
            $drizzle = $this.Aligned.FullName.Replace(
                $this.Aligned.Extension,
                ".xdrz")
            return [System.IO.FileInfo]::new($drizzle)
        }
        return $null
    }
    [void]RemoveCorrectedFiles() {
        if($this.Corrected -and $this.Corrected.Exists){
            if($this.Corrected -ne $this.Calibrated) {
                Write-Verbose "Removing file $($this.Corrected.FullName)"
                $this.Corrected.Delete()
                $this.Corrected.Refresh()
            }
        }
    }
    [void]RemoveDebayeredFiles() {
        if($this.Debayered -and $this.Debayered.Exists){
            if($this.Debayered -ne $this.Corrected) {
                if($this.Debayered -ne $this.Calibrated) {
                    Write-Verbose "Removing file $($this.Debayered.FullName)"
                    $this.Debayered.Delete()
                    $this.Debayered.Refresh()
                }
            }
        }
    }
    [void]RemoveWeightedFiles() {
        if($this.Weighted -and $this.Weighted.Exists){
            if($this.Weighted -ne $this.Debayered)
            {
                if($this.Weighted -ne $this.Calibrated) {
                    Write-Verbose "Removing file $($this.Weighted.FullName)"
                    $this.Weighted.Delete()
                    $this.Weighted.Refresh()
                }
            }
        }
    }
    [void]RemoveAlignedAndDrizzleFiles() {
        $drizzle = $this.GetDrizzleFile()
        if($drizzle -and $drizzle.Exists)
        {
            Write-Verbose "Removing file $($drizzle.FullName)"
            $drizzle.Delete()
        }
        if($this.Aligned -and $this.Aligned.Exists){
            if($this.Aligned -ne $this.Weighted)
            {
                if($this.Aligned -ne $this.Calibrated) {
                    Write-Verbose "Removing file $($this.Aligned.FullName)"
                    $this.Aligned.Delete()
                    $this.Aligned.Refresh()
                }
            }
        }
    }
}

function New-XisfPreprocessingState {
    [OutputType([XisfPreprocessingState])]
    param(
        [Parameter(Mandatory=$true)][XisfFileStats]$Stats,
        [Parameter(Mandatory=$false)][string]$MasterBias,
        [Parameter(Mandatory=$false)][string]$MasterDark,
        [Parameter(Mandatory=$false)][string]$MasterFlat,

        [Parameter(Mandatory=$false)][System.IO.FileInfo]$Calibrated,
        [Parameter(Mandatory=$false)][System.IO.FileInfo]$Corrected,
        [Parameter(Mandatory=$false)][System.IO.FileInfo]$Debayered,
        [Parameter(Mandatory=$false)][System.IO.FileInfo]$Weighted,
        [Parameter(Mandatory=$false)][System.IO.FileInfo]$Aligned
    )
    if((-not $MasterBias) -and ($Stats.History)) {
        $MasterBias = $Stats.History |
            where-object {$_.StartsWith("ImageCalibration.masterBias.fileName")} |
            foreach-object {$_.Split(":")[1].Trim()} |
            select-object -First 1
    }
    if((-not $MasterDark) -and ($Stats.History)) {
        $MasterDark = $Stats.History |
            where-object {$_.StartsWith("ImageCalibration.masterDark.fileName")} |
            foreach-object {$_.Split(":")[1].Trim()} |
            select-object -First 1
    }
    if((-not $MasterFlat) -and ($Stats.History)) {
        $MasterFlat = $Stats.History |
            where-object {$_.StartsWith("ImageCalibration.masterFlat.fileName")} |
            foreach-object {$_.Split(":")[1].Trim()} |
            select-object -First 1
    }

    [XisfPreprocessingState]@{
        Stats      = $Stats
        Path       = $Stats.Path
        MasterBias = $MasterBias
        MasterDark = $MasterDark
        MasterFlat = $MasterFlat
        Calibrated = $Calibrated
        Corrected  = $Corrected
        Debayered  = $Debayered
        Weighted   = $Weighted
        Aligned    = $Aligned
    }
}

function  Get-XisfCalibrationState {
    [OutputType([XisfPreprocessingState])]
    [CmdletBinding()]
    param (
        [Parameter(Mandatory,ValueFromPipeline=$true)][XisfFileStats]$XisfFileStats,
        [Parameter(Mandatory)][System.IO.DirectoryInfo]$CalibratedPath,
        [Parameter()][System.IO.DirectoryInfo[]]$AdditionalSearchPaths,
        [Parameter()][Switch]$Recurse,
        [Parameter()][Switch]$ShowProgress,
        [Parameter()][int]$ProgressTotalCount
    )
    begin{
        $i=0.0
    }
    process{
        $i+=1.0
        if($ShowProgress -and ($i % 5 -eq 0)){
            Write-Progress -Activity "Get-CalibrationState" -PercentComplete (100.0*$i/$ProgressTotalCount) -Status "Searching for matching calibration files $i of $ProgressTotalCount"
        }
        $Path=$XisfFileStats.Path
        $Target=$XisfFileStats.Object.Trim()
        $calibratedFileName = $Path.Name.Substring(0,$Path.Name.Length-$Path.Extension.Length)+"_c.xisf"

        $calibrated = Join-Path ($CalibratedPath.FullName) $calibratedFileName
        if(test-path $calibrated){
            return New-XisfPreprocessingState `
                -Stats $XisfFileStats `
                -Calibrated $calibrated
        }

        #try a target specific subfolder
        $calibrated = Join-Path ($CalibratedPath.FullName) $Target $calibratedFileName
        if(test-path $calibrated){
            return New-XisfPreprocessingState `
                -Stats $XisfFileStats `
                -Calibrated $calibrated
        }

        Write-Debug "Searching sub folders for calibrated file $($calibratedFileName)"
        #try to find a file in any subfolder
        $calibrated = 
            Get-ChildItem $CalibratedPath -Recurse:$Recurse -Filter $calibratedFileName |
            Select-Object -First 1
        if($calibrated) {
            return New-XisfPreprocessingState `
                -Stats $XisfFileStats `
                -Calibrated $calibrated
        }

        if(($AdditionalSearchPaths)){
            Write-Debug "Unable to locate calibration frame... checking additional search paths for $calibratedFileName"
            $calibrated = $AdditionalSearchPaths | 
                Where-Object { test-path $_.FullName } |
                ForEach-Object { 
                    Get-ChildItem -Recurse:$Recurse -Path $_ -Filter $calibratedFileName
                } |
                Where-Object { test-path $_.FullName } | 
                Select-Object -first 1
            if($calibrated) {
                return New-XisfPreprocessingState `
                    -Stats $XisfFileStats `
                    -Calibrated $calibrated
            }
            else{
                Write-Warning "Unable to locate calibration frame from any search paths: $calibratedFileName"
            }
        }
        
        return New-XisfPreprocessingState `
            -Stats $XisfFileStats
    }
    end{
        if($ShowProgress){
            Write-Progress -Activity "Get-CalibrationState" -PercentComplete 100.0 -Status "Done" -Completed
        }
    }
}

function Get-XisfCosmeticCorrectionState{
    [OutputType([XisfPreprocessingState])]
    [CmdletBinding()]
    param (
        [Parameter(Mandatory,ValueFromPipeline=$true)][XisfPreprocessingState]$State,
        [Parameter(Mandatory)][System.IO.DirectoryInfo]$CosmeticCorrectionPath
    )
    process{
        if(-not $State.IsCalibrated()){
            return $State
        }

        $XisfFileStats=$State.Stats
        $Target=$XisfFileStats.Object.Trim()
        $Path=$State.Calibrated
        $cosmeticallyCorrectedFileName = $Path.Name.Substring(0,$Path.Name.Length-$Path.Extension.Length)+"_cc.xisf"

        $cosmeticallyCorrected = Join-Path ($CosmeticCorrectionPath.FullName) $cosmeticallyCorrectedFileName
        if(test-path $cosmeticallyCorrected){
            return New-XisfPreprocessingState `
                -Stats $XisfFileStats `
                -Calibrated $State.Calibrated `
                -MasterBias $State.MasterBias `
                -MasterDark $State.MasterDark `
                -MasterFlat $State.MasterFlat `
                -Corrected $cosmeticallyCorrected
        }

        #try a target specific subfolder
        $cosmeticallyCorrected = Join-Path ($CosmeticCorrectionPath.FullName) $Target $cosmeticallyCorrectedFileName
        if(test-path $cosmeticallyCorrected){
            return New-XisfPreprocessingState `
                -Stats $XisfFileStats `
                -Calibrated $State.Calibrated `
                -MasterBias $State.MasterBias `
                -MasterDark $State.MasterDark `
                -MasterFlat $State.MasterFlat `
                -Corrected $cosmeticallyCorrected
        }

        return $State
    }
}


function Get-XisfDebayerState{
    [OutputType([XisfPreprocessingState])]
    [CmdletBinding()]
    param (
        [Parameter(Mandatory,ValueFromPipeline=$true)][XisfPreprocessingState]$State,
        [Parameter(Mandatory)][System.IO.DirectoryInfo]$DebayerPath
    )
    process{
        if($State.IsDebayered() -or (-not $State.IsCalibrated())){
            return $State
        }

        $XisfFileStats=$State.Stats
        $Target=$XisfFileStats.Object.Trim()
        if($State.IsWeighted()) {
            $Path = $State.Weighted
        }
        elseif($State.IsCorrected()) {
            $Path = $State.Corrected
        }

        $debayeredFileName = $Path.Name.Substring(0,$Path.Name.Length-$Path.Extension.Length)+"_d.xisf"
        $debayered = Join-Path ($DebayerPath.FullName) $debayeredFileName
        if(test-path $debayered){
            return New-XisfPreprocessingState `
                -Stats $XisfFileStats `
                -Calibrated $State.Calibrated `
                -MasterBias $State.MasterBias `
                -MasterDark $State.MasterDark `
                -MasterFlat $State.MasterFlat `
                -Corrected $State.Corrected `
                -Debayered $debayered
        }

        #try a target specific subfolder
        $debayered = Join-Path ($DebayerPath.FullName) $Target $debayeredFileName
        if(test-path $debayered){
            return New-XisfPreprocessingState `
                -Stats $XisfFileStats `
                -Calibrated $State.Calibrated `
                -MasterBias $State.MasterBias `
                -MasterDark $State.MasterDark `
                -MasterFlat $State.MasterFlat `
                -Corrected $State.Corrected `
                -Debayered $debayered
        }

        return $State
    }
}

function Get-XisfSubframeSelectorState{
    [OutputType([XisfPreprocessingState])]
    [CmdletBinding()]
    param (
        [Parameter(Mandatory,ValueFromPipeline=$true)][XisfPreprocessingState]$State,
        [Parameter(Mandatory)][System.IO.DirectoryInfo]$SubframeSelectorPath
    )
    process{
        if($State.IsWeighted() -or (-not $State.IsCalibrated())){
            return $State
        }

        $XisfFileStats=$State.Stats
        $Target=$XisfFileStats.Object.Trim()
        if($State.IsDebayered()) {
            $Path = $State.Debayered
        }
        elseif($State.IsCorrected()) {
            $Path = $State.Corrected
        }
        else{
            $Path=$State.Calibrated
        }

        $SubframeFileName = $Path.Name.Substring(0,$Path.Name.Length-$Path.Extension.Length)+"_a.xisf"

        $subframe = Join-Path ($SubframeSelectorPath.FullName) $SubframeFileName
        if(test-path $subframe){
            return New-XisfPreprocessingState `
                -Stats $XisfFileStats `
                -Calibrated $State.Calibrated `
                -MasterBias $State.MasterBias `
                -MasterDark $State.MasterDark `
                -MasterFlat $State.MasterFlat `
                -Corrected $State.Corrected `
                -Debayered $State.Debayered `
                -Weighted $subframe
        }

        #try a target specific subfolder
        $subframe = Join-Path ($SubframeSelectorPath.FullName) $Target $SubframeFileName
        if(test-path $subframe){
            return New-XisfPreprocessingState `
                -Stats $XisfFileStats `
                -Calibrated $State.Calibrated `
                -MasterBias $State.MasterBias `
                -MasterDark $State.MasterDark `
                -MasterFlat $State.MasterFlat `
                -Corrected $State.Corrected `
                -Debayered $State.Debayered `
                -Weighted $subframe
        }

        return $State
    }
}

function Get-XisfAlignedState{
    [OutputType([XisfPreprocessingState])]
    [CmdletBinding()]
    param (
        [Parameter(Mandatory,ValueFromPipeline=$true)][XisfPreprocessingState]$State,
        [Parameter(Mandatory)][System.IO.DirectoryInfo]$AlignedPath
    )
    process{
        if($State.IsAligned() -or (-not $State.IsCalibrated())){
            return $State
        }

        $XisfFileStats=$State.Stats
        $Target=$XisfFileStats.Object.Trim()
        if($State.IsWeighted()) {
            $Path = $State.Weighted
        }
        elseif($State.IsDebayered()) {
            $Path = $State.Debayered
        }
        elseif($State.IsCorrected()) {
            $Path = $State.Corrected
        }
        else{
            $Path=$State.Calibrated
        }

        $fileName = $Path.Name.Substring(0,$Path.Name.Length-$Path.Extension.Length)+"_r.xisf"

        $aligned = Join-Path ($AlignedPath.FullName) $fileName
        if(test-path $aligned){
            return New-XisfPreprocessingState `
                -Stats $XisfFileStats `
                -Calibrated $State.Calibrated `
                -MasterBias $State.MasterBias `
                -MasterDark $State.MasterDark `
                -MasterFlat $State.MasterFlat `
                -Corrected $State.Corrected `
                -Debayered $State.Debayered `
                -Weighted $State.Weighted `
                -Aligned $aligned
        }

        #try a target specific subfolder
        $aligned = Join-Path ($AlignedPath.FullName) $Target $fileName
        if(test-path $aligned){
            return New-XisfPreprocessingState `
                -Stats $XisfFileStats `
                -Calibrated $State.Calibrated `
                -MasterBias $State.MasterBias `
                -MasterDark $State.MasterDark `
                -MasterFlat $State.MasterFlat `
                -Corrected $State.Corrected `
                -Debayered $State.Debayered `
                -Weighted $State.Weighted `
                -Aligned $aligned
        }

        return $State
    }
}


Function Invoke-XisfPostCalibrationColorImageWorkflow
{
    param(
        [Parameter(Mandatory)][XisfFileStats[]]$RawSubs,
        [Parameter(Mandatory)][System.IO.DirectoryInfo]$CalibrationPath,
        [Parameter(Mandatory)][System.IO.DirectoryInfo[]]$BackupCalibrationPaths,
        [Parameter(Mandatory)][System.IO.DirectoryInfo]$CorrectedOutputPath,
        [Parameter(Mandatory)][System.IO.DirectoryInfo]$WeightedOutputPath,
        [Parameter(Mandatory)][System.IO.DirectoryInfo]$DebayeredOutputPath,
        [Parameter(Mandatory)][System.IO.DirectoryInfo]$AlignedOutputPath,
        [Parameter(Mandatory)][System.IO.DirectoryInfo]$IntegratedImageOutputDirectory,
        [Parameter()][System.IO.FileInfo]$AlignmentReference,
        [Parameter(Mandatory)][System.IO.DirectoryInfo]$DarkLibraryPath,
        [Switch]$RerunCosmeticCorrection,
        [Switch]$SkipCosmeticCorrection,
        [Switch]$RerunDebayer,
        [Switch]$SkipDebayer,
        [Switch]$RerunWeighting,
        [Switch]$SkipWeighting,
        [Switch]$RerunAlignment,
        [string]$CfaPattern="RGGB",
        [int]$PixInsightSlot,
        [string]$ApprovalExpression,
        [string]$WeightingExpression,
        [switch]$GenerateDrizzleData,
        [switch]$GenerateThumbnail,
        
        [Parameter(Mandatory=$false)][decimal]$ClampingThreshold=0.3,        
        [ValidateSet("Auto","Lanczos3","Lanczos4","Bilinear","CubicBSplineFilter","MitchellNetravaliFilter","CatmullRomSplineFilter")]
        [Parameter(Mandatory=$false)][String]$Interpolation="Auto",

        [Parameter(Mandatory=$false)][decimal]$LinearFitLow=8,
        [Parameter(Mandatory=$false)][decimal]$LinearFitHigh=7,
        [Parameter(Mandatory=$false)][string]$Rejection = "LinearFit", #Rejection_ESD
        [Parameter(Mandatory=$false)][string]$Normalization = "AdditiveWithScaling", #AdaptiveNormalization
        [Parameter(Mandatory=$false)][string]$RejectionNormalization = "Scale" #AdaptiveRejectionNormalization
    )

    
    $OutputFileSuffixWithExtension =""
    if($SkipCosmeticCorrection){
        $OutputFileSuffixWithExtension+=".nocc"
    }
    if($SkipDebayer){
        $OutputFileSuffixWithExtension+=".nodebayer"
    }
    if($SkipWeighting){
        $OutputFileSuffixWithExtension+=".noweights"
    }
    if($Rejection -eq "LinearFit"){
        $OutputFileSuffixWithExtension+=".LF.xisf"
    }
    elseif($Rejection -eq "Rejection_ESD"){
        $OutputFileSuffixWithExtension+=".ESD.xisf"
    }
    else{
        $OutputFileSuffixWithExtension+=".$Rejection.xisf"
    }
    $ToProcess = $RawSubs |
        group-object Filter |
        foreach-object {
            $byFilter = $_.Group
            $reference = $byFilter[0]
            $filter=$reference.Filter
                
            $outputFileName = $reference.Object.Trim()
            $byFilter | group-object Exposure | foreach-object {
                $exposure=$_.Group[0].Exposure;
                $outputFileName+=".$filter.$($_.Group.Count)x$($exposure)s"
            }

            $outputFileName += $OutputFileSuffixWithExtension;
            new-object psobject -Property @{
                Filter = $filter
                IntegratedFile = Join-Path ($IntegratedImageOutputDirectory.FullName) $outputFileName
                Subs=$byFilter
            }
        } |
        foreach-object {
            if(Test-Path $_.IntegratedFile){
                Write-Warning "Filter $($_.Filter) has already been processed: $($_.IntegratedFile)"
            }
            else{
                Write-Verbose "Filter $($_.Filter) estimated to produce file: $($_.IntegratedFile)"
                $_.Subs
            }
        }

    $ToProcess |
        Get-XisfCalibrationState `
            -CalibratedPath $CalibrationPath `
            -AdditionalSearchPaths $BackupCalibrationPaths `
            -Verbose -Recurse -ShowProgress -ProgressTotalCount ($ToProcess.Count) |
        foreach-object {
            $x = $_
            if(-not $x.IsCalibrated()){
                Write-Warning "Skipping file: Uncalibrated: $($x.Path)"
            }
            else {
                $x
            }
        } |
        Get-XisfCosmeticCorrectionState -CosmeticCorrectionPath $CorrectedOutputPath |
        foreach-object {
            $x = $_
            if($SkipCosmeticCorrection){
                $x.Corrected=$x.Calibrated
            }
            else{
                if($RerunCosmeticCorrection) {
                    $x.RemoveCorrectedFiles()
                }
            }
            $x
        } |
        Get-XisfCosmeticCorrectionState -CosmeticCorrectionPath $CorrectedOutputPath |
        Group-Object {$_.IsCorrected()} |
        ForEach-Object {
            $group=$_.Group
            if( $group[0].IsCorrected() ){
                $group
            }
            else {
                $group|group-object{
                    ($_.Calibrated | Get-XisfFitsStats | Get-XisfCalibrationState -CalibratedPath $CalibrationPath).MasterDark
                } | foreach-object {
                    
                    $masterDarkFileName = $_.Name
                    $masterDark = get-childitem $DarkLibraryPath *.xisf -Recurse | where-object {$_.Name -eq $masterDarkFileName} | Select-Object -First 1
                    $images = $_.Group
                    Write-Host "Correcting $($images.Count) Images"
                    if(-not $masterDark) {
                        write-warning "Skipping $($images.Count) files... unable to locate master dark: $masterDarkFileName"
                    }
                    else{
                        Invoke-PiCosmeticCorrection `
                            -Images ($images.Calibrated) `
                            -UseAutoHot:$true `
                            -HotAutoSigma 40 `
                            -HotDarkLevel 0.4 `
                            -MasterDark $masterDark `
                            -ColdAutoSigma 2.4 `
                            -UseAutoCold:$true `
                            -OutputPath $CorrectedOutputPath `
                            -PixInsightSlot $PixInsightSlot `
                            -CFAImages
                    }
                }
                $group |
                    Get-XisfCosmeticCorrectionState `
                        -CosmeticCorrectionPath $CorrectedOutputPath
            }
        } |
        Get-XisfDebayerState -DebayerPath $DebayeredOutputPath |
        foreach-object {
            $x = $_
            if($SkipDebayer){
                $x.Debayered=$x.Corrected
            }
            else{
                if($RerunDebayer -or $RerunCosmeticCorrection) {
                    $x.RemoveDebayeredFiles()
                }
            }
            $x
        } |
        Get-XisfDebayerState -DebayerPath $DebayeredOutputPath |
        Group-Object {$_.IsDebayered()} |
        ForEach-Object {
            $group=$_.Group
            if( $group[0].IsDebayered() ){
                $group
            }
            else {
                Write-Host "Debayering $($group.Count) Images"
                Invoke-PiDebayer `
                    -PixInsightSlot $PixInsightSlot `
                    -Images ($group.Corrected) `
                    -OutputPath $DebayeredOutputPath `
                    -CfaPattern $CfaPattern

                $group |
                    Get-XisfDebayerState `
                        -DebayerPath $DebayeredOutputPath
            }
        } |
        Get-XisfSubframeSelectorState -SubframeSelectorPath $WeightedOutputPath |
        foreach-object {
            $x = $_
            if($SkipWeighting){
                $x.Weighted=$x.Debayered
            }
            else{
                if($RerunWeighting -or $RerunDebayer -or $RerunCosmeticCorrection) {
                    $x.RemoveWeightedFiles()
                }
            }
            $x
        } |
        Group-Object { $_.IsWeighted() } |
        ForEach-Object {
            $group=$_.Group
            if($group | Where-Object {$_.IsWeighted()}) {
                $group
            }
            else{
                $group | group-object {$_.Stats.Filter } | foreach-object {
                    $byFilter = $_.Group
                    $filter=$byFilter[0].Stats.Filter
                    Write-Host "Weighting $($byFilter.Count) Images for filter $filter"
                    Start-PiSubframeSelectorWeighting `
                        -PixInsightSlot $PixInsightSlot `
                        -OutputPath $WeightedOutputPath `
                        -Images ($byFilter.Debayered) `
                        -ApprovalExpression $ApprovalExpression `
                        -WeightingExpression $WeightingExpression
            
                }
                $group |
                    Get-XisfSubframeSelectorState `
                        -SubframeSelectorPath $WeightedOutputPath
            }
        } |
        Get-XisfAlignedState -AlignedPath $AlignedOutputPath |
        foreach-object {
            $x = $_
            if($RerunAlignment -or $RerunWeighting -or $RerunDebayer -or $RerunCosmeticCorrection) {
                $x.RemoveAlignedAndDrizzleFiles()
            }
            $x
        } |
        Group-Object {$_.IsAligned()} |
        ForEach-Object {
            $group=$_.Group
            if( $group[0].IsAligned() ){
                $group
            }
            else {
                $approved = $group |
                    Where-object {$_.IsWeighted()} |
                    foreach-object{$_.Weighted} |
                    Get-XisfFitsStats
                if(-not $approved){
                    $group
                }
                else {
                    if($AlignmentReference -and (-not $AlignmentReference.Exists) )
                    {
                        Write-Warning "The specified alignment file could not be found $($AlignmentReference). A new reference will be selected."
                    }
                    $reference = $AlignmentReference
                    if(-not ($alignmentReference -and (test-path $reference))){
                        $reference =  ($approved |
                            Sort-Object SSWeight -Descending |
                            Select-Object -First 1).Path
                    }

                    Write-Host "Aligning $($approved.Count) Images to $($reference)"
                    Invoke-PiStarAlignment `
                        -PixInsightSlot $PixInsightSlot `
                        -Images ($approved.Path) `
                        -ReferencePath ($reference) `
                        -OutputPath $AlignedOutputPath `
                        -Interpolation:$Interpolation `
                        -ClampingThreshold:$ClampingThreshold
                    $group |
                        Get-XisfAlignedState `
                            -AlignedPath $AlignedOutputPath
                    
                }
            }
        } |
        Group-Object {$_.IsAligned()} |
        ForEach-Object {
            $group=$_.Group
            if( $group[0].IsAligned() ){
                $approved = $group |
                    foreach-object{$_.Aligned} |
                    Get-XisfFitsStats
                $reference =  $approved |
                    Sort-Object SSWeight -Descending |
                    Select-Object -First 1
                
                $group|group-object {$_.Stats.Filter}|foreach-object {
                    $byFilter=$_.Group |
                        where-object {$_.Aligned -and (Test-Path $_.Aligned)} |
                        foreach-object{$_.Aligned} |
                        Get-XisfFitsStats
                    $filter=$byFilter[0].Filter
                        
                    $outputFileName = $reference.Object.Trim()
                    $byFilter | group-object Exposure | foreach-object {
                        $exposure=$_.Group[0].Exposure;
                        $outputFileName+=".$filter.$($_.Group.Count)x$($exposure)s"
                    }
                    $outputFileName += $OutputFileSuffixWithExtension;
                    $outputFile = Join-Path ($IntegratedImageOutputDirectory.FullName) $outputFileName
                    if(-not (test-path $outputFile)) {
                        write-host ("Integrating "+ $outputFileName)
                        $toStack = $byFilter | sort-object SSWeight -Descending
                        $toStack | 
                            Group-Object Filter | 
                            foreach-object {$dur=$_.Group|Measure-ExposureTime -TotalSeconds; new-object psobject -Property @{Filter=$_.Name; ExposureTime=$dur}} |
                            foreach-object {
                                write-host "$($_.Filter): $($_.Exposure)"
                            }
                        $weightKeyword="SSWEIGHT"
                        if($SkipWeighting){
                            $weightKeyword=$null
                        }
                        try {
                        Invoke-PiLightIntegration `
                            -Images ($toStack|foreach-object {$_.Path}) `
                            -OutputFile $outputFile `
                            -KeepOpen `
                            -Rejection:$Rejection `
                            -Normalization:$Normalization `
                            -RejectionNormalization:$RejectionNormalization `
                            -LinearFitLow:$LinearFitLow `
                            -LinearFitHigh:$LinearFitHigh `
                            -PixInsightSlot $PixInsightSlot `
                            -GenerateDrizzleData:$GenerateDrizzleData `
                            -WeightKeyword:$weightKeyword `
                            -GenerateThumbnail:$GenerateThumbnail
                        }
                        catch {
                            write-warning $_.ToString()
                            throw
                        }
                    }
                }

                $group
            }
            else{
                $group
            }
        }
}

Function Invoke-XisfPostCalibrationMonochromeImageWorkflow
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][XisfFileStats[]]$RawSubs,
        [Parameter(Mandatory)][System.IO.DirectoryInfo]$CalibrationPath,
        [Parameter(Mandatory)][System.IO.DirectoryInfo[]]$BackupCalibrationPaths,
        [Parameter(Mandatory)][System.IO.DirectoryInfo]$CorrectedOutputPath,
        [Parameter(Mandatory)][System.IO.DirectoryInfo]$WeightedOutputPath,
        [Parameter(Mandatory)][System.IO.DirectoryInfo]$AlignedOutputPath,
        [Parameter(Mandatory)][System.IO.DirectoryInfo]$IntegratedImageOutputDirectory,
        [Parameter()][System.IO.FileInfo]$AlignmentReference,
        [Parameter(Mandatory)][System.IO.DirectoryInfo]$DarkLibraryPath,
        [Switch]$RerunCosmeticCorrection,
        [Switch]$SkipCosmeticCorrection,
        [Switch]$RerunWeighting,
        [Switch]$SkipWeighting,
        [Switch]$RerunAlignment,
        [int]$PixInsightSlot,
        [string]$ApprovalExpression,
        [string]$WeightingExpression,
        [Switch]$PSFSignalWeightWeighting,
        [switch]$GenerateDrizzleData,
        [switch]$GenerateThumbnail,

        [Parameter(Mandatory=$false)][decimal]$ClampingThreshold=0.3,        
        [ValidateSet("Auto","Lanczos3","Lanczos4","Bilinear","CubicBSplineFilter","MitchellNetravaliFilter","CatmullRomSplineFilter")]
        [Parameter(Mandatory=$false)][String]$Interpolation="Auto",

        [Parameter(Mandatory=$false)][decimal]$LinearFitLow=8,
        [Parameter(Mandatory=$false)][decimal]$LinearFitHigh=7,
        [Parameter(Mandatory=$false)][string]$Rejection = "LinearFit", #Rejection_ESD
        [Parameter(Mandatory=$false)][string]$Normalization = "AdditiveWithScaling", #AdaptiveNormalization
        [Parameter(Mandatory=$false)][string]$RejectionNormalization = "Scale", #AdaptiveRejectionNormalization
        [Parameter(Mandatory=$false)][decimal]$HotAutoSigma=40,
        [Switch]$KeepOpen
    )
    $OutputFileSuffixWithExtension =""
    if($SkipCosmeticCorrection){
        $OutputFileSuffixWithExtension+=".nocc"
    }
    if($SkipDebayer){
        $OutputFileSuffixWithExtension+=".nodebayer"
    }
    if($SkipWeighting){
        if($PSFSignalWeightWeighting){
            $OutputFileSuffixWithExtension+=".PSFSW"
        }
        else{
            $OutputFileSuffixWithExtension+=".noweights"
        }        
    }
    if($Rejection -eq "LinearFit"){
        $OutputFileSuffixWithExtension+=".LF.xisf"
    }
    elseif($Rejection -eq "Rejection_ESD"){
        $OutputFileSuffixWithExtension+=".ESD.xisf"
    }
    else{
        $OutputFileSuffixWithExtension+=".$Rejection.xisf"
    }

    $ToProcess = $RawSubs |
        group-object Filter |
        foreach-object {
            $byFilter = $_.Group
            $reference = $byFilter[0]
            $filter=$reference.Filter
                
            $outputFileName = $reference.Object.Trim()
            $byFilter | group-object Exposure | foreach-object {
                $exposure=$_.Group[0].Exposure;
                $outputFileName+=".$filter.$($_.Group.Count)x$($exposure)s"
            }

            $outputFileName += $OutputFileSuffixWithExtension;
            new-object psobject -Property @{
                Filter = $filter
                IntegratedFile = Join-Path ($IntegratedImageOutputDirectory.FullName) $outputFileName
                Subs=$byFilter
            }
        } |
        foreach-object {
            if(Test-Path $_.IntegratedFile){
                Write-Warning "Filter $($_.Filter) has already been processed: $($_.IntegratedFile)"
            }
            else{
                Write-Verbose "Filter $($_.Filter) estimated to produce file: $($_.IntegratedFile)"
                $_.Subs
            }
        }

    

    $ToProcess |
        Get-XisfCalibrationState `
            -CalibratedPath $CalibrationPath `
            -AdditionalSearchPaths $BackupCalibrationPaths `
            -Verbose -Recurse -ShowProgress -ProgressTotalCount ($ToProcess.Count) |
        foreach-object {
            $x = $_
            if(-not $x.IsCalibrated()){
                Write-Warning "Skipping file: Uncalibrated: $($x.Path)"
            }
            else {
                $x
            }
        } |
        Get-XisfCosmeticCorrectionState -CosmeticCorrectionPath $CorrectedOutputPath |
        foreach-object {
            $x = $_
            if($SkipCosmeticCorrection){
                $x.Corrected=$x.Calibrated
            }
            else{
                if($RerunCosmeticCorrection) {
                    $x.RemoveCorrectedFiles()
                }
            }
            $x
        } |
        Get-XisfCosmeticCorrectionState -CosmeticCorrectionPath $CorrectedOutputPath |
        Group-Object {$_.IsCorrected()} |
        ForEach-Object {
            $group=$_.Group
            if( $group[0].IsCorrected() ){
                $group
            }
            else {
                $group|group-object{
                    ($_.Calibrated | Get-XisfFitsStats | Get-XisfCalibrationState -CalibratedPath $CalibrationPath).MasterDark
                } | foreach-object {
                    $masterDarkFileName = $_.Name
                    $masterDark = get-childitem $DarkLibraryPath *.xisf -Recurse | where-object {$_.Name -eq $masterDarkFileName} | Select-Object -First 1
                    $images = $_.Group
                    Write-Host "Correcting $($images.Count) Images"
                    Invoke-PiCosmeticCorrection `
                        -Images ($images.Calibrated) `
                        -UseAutoHot:$true `
                        -HotAutoSigma 4.0 `
                        -HotDarkLevel 0.4 `
                        -MasterDark $masterDark `
                        -ColdAutoSigma 2.4 `
                        -UseAutoCold:$true `
                        -OutputPath $CorrectedOutputPath `
                        -PixInsightSlot $PixInsightSlot `
                        -KeepOpen:$KeepOpen.IsPresent
                }
                $group |
                    Get-XisfCosmeticCorrectionState `
                        -CosmeticCorrectionPath $CorrectedOutputPath
            }
        } |
        Get-XisfSubframeSelectorState -SubframeSelectorPath $WeightedOutputPath |
        foreach-object {
            $x = $_
            if($SkipWeighting){
                $x.Weighted=$x.Corrected
            }
            else{
                if($RerunWeighting -or $RerunCosmeticCorrection) {
                    $x.RemoveWeightedFiles()
                }
            }
            $x
        } |
        Group-Object { $_.IsWeighted() } |
        ForEach-Object {
            $group=$_.Group
            if($group | Where-Object {$_.IsWeighted()}) {
                $group
            }
            else{
                $group | group-object {$_.Stats.Filter } | foreach-object {
                    $byFilter = $_.Group
                    $filter=$byFilter[0].Stats.Filter
                    Write-Host "Weighting $($byFilter.Count) Images for filter $filter"
                    Start-PiSubframeSelectorWeighting `
                        -PixInsightSlot $PixInsightSlot `
                        -OutputPath $WeightedOutputPath `
                        -Images ($byFilter.Corrected) `
                        -ApprovalExpression $ApprovalExpression `
                        -WeightingExpression $WeightingExpression
            
                }
                $group |
                    Get-XisfSubframeSelectorState `
                        -SubframeSelectorPath $WeightedOutputPath
            }
        } |
        Get-XisfAlignedState -AlignedPath $AlignedOutputPath |
        foreach-object {
            $x = $_
            if($RerunAlignment -or $RerunWeighting -or $RerunCosmeticCorrection) {
                $x.RemoveAlignedAndDrizzleFiles()
            }
            $x
        } |
        Group-Object {$_.IsAligned()} |
        ForEach-Object {
            $group=$_.Group
            if( $group[0].IsAligned() ){
                $group
            }
            else {
                $approved = $group |
                    Where-object {$_.IsWeighted()} |
                    foreach-object{$_.Weighted} |
                    Get-XisfFitsStats
                if(-not $approved){
                    $group
                }
                else {
                    if($AlignmentReference -and (-not $AlignmentReference.Exists) )
                    {
                        Write-Warning "The specified alignment file could not be found $($AlignmentReference). A new reference will be selected."
                    }
                    $reference = $AlignmentReference
                    if(-not ($alignmentReference -and (test-path $reference))){
                        $reference =  ($approved |
                            Sort-Object SSWeight -Descending |
                            Select-Object -First 1).Path
                    }

                    Write-Host "Aligning $($approved.Count) Images to $($reference)"
                    Invoke-PiStarAlignment `
                        -PixInsightSlot $PixInsightSlot `
                        -Images ($approved.Path) `
                        -ReferencePath ($reference) `
                        -OutputPath $AlignedOutputPath `
                        -Interpolation:$Interpolation `
                        -ClampingThreshold:$ClampingThreshold `
                        -KeepOpen:$KeepOpen.IsPresent
                    $group |
                        Get-XisfAlignedState `
                            -AlignedPath $AlignedOutputPath
                    
                }
            }
        } |
        Group-Object {$_.IsAligned()} |
        ForEach-Object {
            $group=$_.Group
            if( $group[0].IsAligned() ){
                $approved = $group |
                    foreach-object{$_.Aligned} |
                    Get-XisfFitsStats
                $reference =  $approved |
                    Sort-Object SSWeight -Descending |
                    Select-Object -First 1
                
                $group|group-object {$_.Stats.Filter}|foreach-object {
                    $byFilter=$_.Group |
                        where-object {$_.Aligned -and (Test-Path $_.Aligned)} |
                        foreach-object{$_.Aligned} |
                        Get-XisfFitsStats
                    $filter=$byFilter[0].Filter
                        
                    $outputFileName = $reference.Object.Trim()
                    $byFilter | group-object Exposure | foreach-object {
                        $exposure=$_.Group[0].Exposure;
                        $outputFileName+=".$filter.$($_.Group.Count)x$($exposure)s"
                    }

                    $outputFileName += $OutputFileSuffixWithExtension;
                    $outputFile = Join-Path ($IntegratedImageOutputDirectory.FullName) $outputFileName
                    if(-not (test-path $outputFile)) {
                        write-host ("Integrating  "+ $outputFileName)
                        $toStack = $byFilter | sort-object SSWeight -Descending
                        $toStack | 
                        Group-Object Filter | 
                        foreach-object {$dur=$_.Group|Measure-ExposureTime -TotalSeconds; new-object psobject -Property @{Filter=$_.Name; ExposureTime=$dur}} |
                        foreach-object {
                            write-host "$($_.Filter): $($_.Exposure)"
                        }
                        $weightKeyword="SSWEIGHT"
                        if($SkipWeighting){
                            $weightKeyword=$null
                        }
                        try {
                            if($PSFSignalWeightWeighting){
                                Invoke-PiLightIntegration `
                                    -Images ($toStack|foreach-object {$_.Path}) `
                                    -OutputFile $outputFile `
                                    -KeepOpen `
                                    -Rejection:$Rejection `
                                    -Normalization:$Normalization `
                                    -RejectionNormalization:$RejectionNormalization `
                                    -LinearFitLow:$LinearFitLow `
                                    -LinearFitHigh:$LinearFitHigh `
                                    -PixInsightSlot $PixInsightSlot `
                                    -GenerateDrizzleData:$GenerateDrizzleData `
                                    -WeightKeyword:$weightKeyword `
                                    -WeightMode:"PSFSignalWeight" `
                                    -GenerateThumbnail:$GenerateThumbnail
                            }
                            else{
                                Invoke-PiLightIntegration `
                                    -Images ($toStack|foreach-object {$_.Path}) `
                                    -OutputFile $outputFile `
                                    -KeepOpen `
                                    -Rejection:$Rejection `
                                    -Normalization:$Normalization `
                                    -RejectionNormalization:$RejectionNormalization `
                                    -LinearFitLow:$LinearFitLow `
                                    -LinearFitHigh:$LinearFitHigh `
                                    -PixInsightSlot $PixInsightSlot `
                                    -GenerateDrizzleData:$GenerateDrizzleData `
                                    -WeightKeyword:$weightKeyword `
                                    -GenerateThumbnail:$GenerateThumbnail
                            }

                        }
                        catch {
                            write-warning $_.ToString()
                            throw
                        }
                    }
                }

                $group
            }
            else{
                $group
            }
        }
}
Function Get-MasterBiasLibrary {
    param (
        [Parameter(Mandatory)][System.IO.DirectoryInfo]$Path,
        [Parameter(Mandatory)][string]$Pattern
    )
    begin{
        $regex=[System.Text.RegularExpressions.Regex]::new($Pattern)
    }
    process{
        Get-ChildItem ($Path.FullName) -File |
        ForEach-Object {
            $y = $regex.Matches($_.Name)
            if($y.Success){
                $gain = $y.Groups | where-object Name -eq "gain" | foreach-object {$_.Value}
                $offset = $y.Groups | where-object Name -eq "offset" | foreach-object {$_.Value}
                $numberOfExposures = $y.Groups | where-object Name -eq "numberOfExposures" | foreach-object {$_.Value}
                $exposure = $y.Groups | where-object Name -eq "exposure" | foreach-object {$_.Value}

                if(-not $gain){
                    write-warning "Unable to parse value for gain file... $($_.Name)"
                }
                elseif(-not $offset){
                    write-warning "Unable to parse value for offset file... $($_.Name)"
                }
                elseif(-not $exposure){
                    write-warning "Unable to parse value for exposure file... $($_.Name)"
                }
                else{
                    $x = $_ | Get-XisfFitsStats
                    $x.Gain=[Decimal]$gain
                    $x.Offset=[Decimal]$offset
                    $x.Exposure=[decimal]$exposure
                    $x | Add-Member -NotePropertyName "NumberOfExposures" -NotePropertyValue ([int]$numberOfExposures)
                    $x
                }
            }
            else{
                Write-Warning "Skipping file: does not match specified pattern... $($_.Name)"
            }
        }
    }
}
Function Get-MasterDarkLibrary {
    param (
        [Parameter(Mandatory)][System.IO.DirectoryInfo]$Path,
        [Parameter(Mandatory)][string]$Pattern
    )
    begin{
        $regex=[System.Text.RegularExpressions.Regex]::new($Pattern)
    }
    process{
        Get-ChildItem ($Path.FullName) -File |
        ForEach-Object {
            $y = $regex.Matches($_.Name)
            if($y.Success){
                $gain = $y.Groups | where-object Name -eq "gain" | foreach-object {$_.Value}
                $offset = $y.Groups | where-object Name -eq "offset" | foreach-object {$_.Value}
                $temp = $y.Groups | where-object Name -eq "temp" | foreach-object {$_.Value}
                $numberOfExposures = $y.Groups | where-object Name -eq "numberOfExposures" | foreach-object {$_.Value}
                $exposure = $y.Groups | where-object Name -eq "exposure" | foreach-object {$_.Value}

                if(-not $gain){
                    write-warning "Unable to parse value for gain file... $($_.Name)"
                }
                elseif(-not $offset){
                    write-warning "Unable to parse value for offset file... $($_.Name)"
                }
                elseif(-not $temp){
                    write-warning "Unable to parse value for temp file... $($_.Name)"
                }
                elseif(-not $exposure){
                    write-warning "Unable to parse value for exposure file... $($_.Name)"
                }
                else{
                    $x = $_ | Get-XisfFitsStats
                    $x.SetTemp = [Decimal]$temp
                    $x.Gain=[Decimal]$gain
                    $x.Offset=[Decimal]$offset
                    $x.Exposure=[decimal]$exposure
                    $x | Add-Member -NotePropertyName "NumberOfExposures" -NotePropertyValue ([int]$numberOfExposures)
                    $x
                }
            }
            else{
                Write-Warning "Skipping file: does not match specified pattern... $($_.Name)"
            }
        }
    }
}