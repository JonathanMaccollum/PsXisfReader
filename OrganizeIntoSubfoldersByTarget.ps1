Import-Module PsXisfReader
$path = "F:\PixInsightLT\Calibrated"
Get-ChildItem $path -File |
    Get-XisfFitsStats |
    Where-Object Object -ne $null |
    Select-Object -First 3000 |
    foreach-object {
        $x=$_
        $object = $x.Object
        $targetDirectory = Join-Path $path $object
        [System.IO.Directory]::CreateDirectory($targetDirectory) > $null

        $targetFile = Join-Path $targetDirectory ($x.Path.Name)
        if(-not (test-path $targetFile)){
            Move-Item ($x.Path) -Destination $targetFile
        }        
    }