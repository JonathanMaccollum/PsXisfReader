Import-Module $PSScriptRoot/PsXisfReader.psm1
Import-Module $PSScriptRoot/PixInsightPreProcessing.psm1


#find top weighted items by object in a folder
Get-XisfFile -Path "S:\PixInsight\Weighted" |
    group-object Object |
    foreach-object {
        $group=$_
        $target = $group.Name
        $topWeight = $group.Group | sort-object SSWeight -Descending | select-object -first 1
        write-host "$target - $($topWeight.SSWeight): $($topWeight.Path)"
    }


$PathToCache = "E:\Astrophotography\Cache.20200416.clixml"
if($null -eq $cache){
    if(Test-Path $PathToCache){
        $cache = Import-Clixml -Path $PathToCache
    }
    else {
        $cache = new-object hashtable
    }
}

$data = @()
$data += Get-ChildItem "E:\Astrophotography" *.xisf -Recurse
$withStats=$data |
    Where-object {-not $_.FullName.Contains('Reject')} |
    Where-object {-not $_.FullName.Contains('reject')} |
    Where-object {-not $_.FullName.Contains('Testing')} |
    Where-object {-not $_.FullName.Contains('testing')} |
    Where-object {-not $_.FullName.Contains('cloud')} |
    Where-object {-not $_.FullName.Contains('Cloud')} |
    foreach-object { 
        Get-XisfFitsStats -Path $_ -Cache $cache         
    }
$withStats|
    where-object ImageType -eq 'LIGHT' |
    where-object ImageType -NE $null |
    sort-object ObsDateMinus12hr |
    group-object {
        $obsDate=$_.ObsDateMinus12hr
        $obsDate.AddDays(-$obsDate.Day+1).ToString("yyyy-MM")
    } |
    foreach-object {
        $sum=($_.Group|measure-object Exposure -Sum)
    
        "$($_.Name) - $([TimeSpan]::FromSeconds($sum.Sum).TotalHours.ToString("0.00"))hrs"
    }

$cache|Export-Clixml -Path $PathToCache -Force