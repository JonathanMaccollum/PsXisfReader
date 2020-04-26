Import-Module "C:\Users\jmaccollum\Downloads\csharpfits.1.1.0\lib\Net20\CSharpFITS_v1.1.dll"
Function Get-FitsHeaders
{
    param
    (
        [Parameter(ValueFromPipeline=$true,Mandatory=$true)][System.IO.FileInfo]$Path,
        [Parameter(ValueFromPipeline=$false)][Hashtable]$Cache
    )
    $fits = new-object nom.tam.fits.Fits $Path.FullName
    $hdu = $fits.GetHDU(0)
    $header = $hdu.Header
    $x = new-object psobject
    for($i=0;$i -lt $header.NumberOfCards; $i++){
        ($card = $header.GetCard($i)) >> $null
        $card
        $parts = $card.Split('=')
        $key = $parts[0].Trim()
        $value=$null
        if($parts[1])
        {
            $value=$parts[1].Trim()
        }
        Add-Member -InputObject $x -Name $key -Value $value -MemberType NoteProperty
    }
    $x
}
<#
$fitsFiles = Get-ChildItem "D:\Backups\Camera\2019" *.fits -Recurse
$fitsStats = $fitsFiles | Foreach-Object {
    $x = Get-FitsHeaders -Path $_ 

}
$lights = $fitsStats | 
    where-object Object -ne $null |
    where-object ImageType -eq 'LIGHT' |
#>

if($cache -eq $null)
{
    $cache = new-object hashtable
    #$cache|Export-Clixml -Path "D:\Backups\Camera\Astrophotography\Cache.20200203.clixml" -Force
    $cache = Import-Clixml -Path "D:\Backups\Camera\Astrophotography\Cache.20200203.clixml"
}

$data = @()
$data += Get-ChildItem "D:\Backups\Camera\Astrophotography" *.xisf -Recurse
$data += Get-ChildItem "D:\Backups\Camera\2019" *.xisf -Recurse

$stats = 
    $data |
    Where-object {-not $_.FullName.Contains('Reject')} |
    Where-object {-not $_.FullName.Contains('reject')} |
    Where-object {-not $_.FullName.Contains('Testing')} |
    Where-object {-not $_.FullName.Contains('testing')} |
    Where-object {-not $_.FullName.Contains('cloud')} |
    Where-object {-not $_.FullName.Contains('Cloud')} |
    foreach-object { Get-XisfFitsStats -Path $_ -Cache $cache } |
    where-object ImageType -eq 'LIGHT'
cls
$stats|
    where-object ImageType -NE $null |
    where-object ImageType -eq 'LIGHT'|
    sort-object LocalDate|
    group-object {([DateTime]$_.LocalDate).Month} | foreach-object {
    $sum=($_.Group|measure-object Exposure -Sum)
    
    "$($_.Name) - $([TimeSpan]::FromSeconds($sum.Sum).TotalHours.ToString("0.00"))hrs"
}
$FinishedTargets = @(
'NGC 6871 Panel 0'
'North America Nebula Panel 0'
'Ghost Nebula'
'Hidden Galaxy'
'Elephant and Flying Bat'
'IC 1396', 'Elephant Widefield'
'Iris Nebula'
'NGC2403'
'Pacman'
'Sh 2-132 Lion Nebula'
'Tulip'
'Tulip Widefield'
'Veil'
'WR 134'
)
$stats|
    where-object ImageType -eq 'LIGHT'|
    where-object Filter -ne $null |
    where-object Object -ne $null |
    Where-Object Object -CNotIn $FinishedTargets |
    sort-object Object,Filter |
    group-object Object |
    foreach-object {
        Write-Host " "
        Write-Host "==========================================="
        Write-Host " "
        $target = $_.Name
        write-host $target
        $_.Group | Group-Object Filter  | foreach-object {
            $filter =$_.Name
            $duration = $_.Group | measure-object Exposure -Sum
            Write-Host "     - $($filter): $($duration.Count) files = $([TimeSpan]::FromSeconds($duration.Sum))"
        }
        $totals = $_.Group | measure-object Exposure -Sum
        Write-Host "     ======================================"
        Write-Host "     - Total: $($totals.Count) files = $([TimeSpan]::FromSeconds($totals.Sum))"

    }
[TimeSpan]::FromSeconds(
($stats|
    where-object ImageType -eq 'LIGHT'|
    where-object Object -ne $null |
    measure-object Exposure -Sum).Sum)