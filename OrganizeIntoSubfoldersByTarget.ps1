Import-Module $PSScriptRoot/PsXisfReader.psm1 -Force
Import-Module $PSScriptRoot/PixInsightPreProcessing.psm1

$path = "E:\PixInsightLT\Calibrated"

$path|Get-ChildItem  -File |
    Get-XisfFitsStats |
    Where-Object Object -ne $null |
    Select-Object -First 13000 |
    foreach-object {
        $x=$_
        $object = $x.Object
        $targetDirectory = Join-Path $path $object
        [System.IO.Directory]::CreateDirectory($targetDirectory) > $null

        $targetFile = Join-Path $targetDirectory ($x.Path.Name)
        Move-Item ($x.Path) -Destination $targetFile
    }