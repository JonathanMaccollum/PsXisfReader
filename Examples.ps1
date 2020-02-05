Import-Module $PSScriptRoot/PsXisfReader.psm1



$PathToCache = "D:\Backups\Camera\Astrophotography\Cache.20200203.clixml"
if($null -eq $cache){
    if(-not (Test-Path $PathToCache)){
        $cache = Import-Clixml -Path PathToCache
    }
    else {
        $cache = new-object hashtable
    }
}

$data = @()
$data += Get-ChildItem "D:\Backups\Camera\Astrophotography" *.xisf -Recurse
$data += Get-ChildItem "D:\Backups\Camera\2019" *.xisf -Recurse

$data |
    Where-object {-not $_.FullName.Contains('Reject')} |
    Where-object {-not $_.FullName.Contains('reject')} |
    Where-object {-not $_.FullName.Contains('Testing')} |
    Where-object {-not $_.FullName.Contains('testing')} |
    Where-object {-not $_.FullName.Contains('cloud')} |
    Where-object {-not $_.FullName.Contains('Cloud')} |
    foreach-object { Get-XisfFitsStats -Path $_ -Cache $cache } |
    where-object ImageType -eq 'LIGHT' |
    where-object ImageType -NE $null |
    sort-object LocalDate |
    group-object {([DateTime]$_.LocalDate).Month} |
    foreach-object {
        $sum=($_.Group|measure-object Exposure -Sum)
    
        "$($_.Name) - $([TimeSpan]::FromSeconds($sum.Sum).TotalHours.ToString("0.00"))hrs"
    }

$cache|Export-Clixml -Path $PathToCache -Force