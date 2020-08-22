import-module $PSScriptRoot/PsXisfReader.psd1 -Force
$ErrorActionPreference="STOP"

Clear-Host

$PathToCache = "E:\Astrophotography\50mm\Patchwork Cygnus\Cache.clixml"
if(Test-Path $PathToCache){
    $cache = Import-Clixml -Path $PathToCache
}
else {
    $cache = new-object hashtable
}
$data =
    (Get-ChildItem "E:\Astrophotography\50mm" -Directory -Filter "PatchworkCygnus*")+
    (Get-ChildItem "E:\Astrophotography\135mm" -Directory -Filter "Sadr*") +
    (Get-ChildItem "E:\Astrophotography\135mm" -Directory -Filter "DWB*") |
    ForEach-Object {Get-XisfLightFrames -Path $_ -Recurse -Cache $cache } |
    Where-Object {-not $_.HasTokensInPath(@("reject","process","testing","clouds","draft","cloudy"))} |
    Where-Object {-not $_.IsIntegratedFile()} |
    Where-Object {$_.Filter}
$cache|Export-Clixml -Path $PathToCache -Force

$filters = $data|Group-Object Filter|ForEach-Object{$_.Name}
$data|
    Group-Object FocalLength,Object,Filter,Exposure |
    ForEach-Object {
        $x=$_.Group
        $count=$_.Count
        $r=new-object psobject -Property @{
            FocalLength=$x[0].FocalLength
            Object=$x[0].Object
            Filter=$x[0].Filter
            Exposure=$x[0].Exposure
            Count=$count
        }
        Add-Member -InputObject $r -Value "$($count)x$($r.Exposure)s" -MemberType NoteProperty -Name "Summary"
        Add-Member -InputObject $r -Value ([TimeSpan]::FromSeconds($($count)*$($r.Exposure))) -MemberType NoteProperty -Name "TotalExposure"
        Add-Member -InputObject $r -Value ($($count)*$($r.Exposure)) -MemberType NoteProperty -Name "TotalExposureS"
        $r | Write-Output
    } |
    Group-Object FocalLength,Object |
    Foreach-Object {
        $x=$_.Group
        $r = new-object psobject -Property @{
            Object = $x[0].Object
            FocalLength=$x[0].FocalLength
        }
        $filters|foreach-object {
            $filter = $_
            $subs = $x | where-object Filter -eq $_
            #Add-Member -InputObject $r -Value $subs -MemberType NoteProperty -Name $filter
            $total = [TimeSpan]::FromSeconds(($subs | Measure-Object -Property TotalExposureS -Sum).Sum)
            if($total -eq [TimeSpan]::Zero) {
                $total = $null
            }
            Add-Member -InputObject $r -Value $total -MemberType NoteProperty -Name "Total $filter"
        }
        $r | Write-Output
    } |
    Format-Table