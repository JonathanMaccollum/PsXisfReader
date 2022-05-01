Get-ChildItem D:\Backups\Camera\Dropoff\NINA\Sequences *.json |
foreach-object {
    $file=$_.FullName
    $content = Get-Content $_.FullName 
    if($content.Contains("1999")){
        Write-Host $file
        $json = $content | ConvertFrom-Json
        $x=$file
    }    
}