import-module $PSScriptRoot/PsXisfReader.psd1 -Force
$ErrorActionPreference="STOP"

Clear-Host

$data =
    (Get-ChildItem "E:\Astrophotography\50mm" -Directory -Filter "PatchworkCygnus*") |
    # (Get-ChildItem "E:\Astrophotography\135mm" -Directory -Filter "Sadr*") +
    # (Get-ChildItem "E:\Astrophotography\135mm" -Directory -Filter "Squid*") +
    # (Get-ChildItem "E:\Astrophotography\135mm" -Directory -Filter "PatchworkCygnus*") +
    # (Get-ChildItem "E:\Astrophotography\1000mm" -Directory -Filter "Sadr*") +
    # (Get-ChildItem "E:\Astrophotography\1000mm" -Directory -Filter "vdb131*") +
    # (Get-ChildItem "E:\Astrophotography\1000mm" -Directory -Filter "Sh 2-108*") +
    # (Get-ChildItem "E:\Astrophotography\1000mm" -Directory -Filter "Eye of*") +
    # (Get-ChildItem "E:\Astrophotography\1000mm" -Directory -Filter "Ring Nebula") +
    # (Get-ChildItem "E:\Astrophotography\135mm" -Directory -Filter "DWB*") |
    ForEach-Object {Get-XisfLightFrames -Path $_ -Recurse -UseCache } |
    Where-Object {-not $_.HasTokensInPath(@("reject","process","testing","clouds","draft","cloudy"))} |
    Where-Object {-not $_.IsIntegratedFile()} |
    Where-Object {$_.Filter}

$filters = $data|Group-Object Filter|ForEach-Object{$_.Name}
$stats=$data|
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
            $total = [TimeSpan]::FromSeconds(($subs | Measure-Object -Property TotalExposureS -Sum).Sum)
            if($total -eq [TimeSpan]::Zero) {
                $total = $null
            }
            Add-Member -InputObject $r -Value $total -MemberType NoteProperty -Name "Total $filter"
        }
        $r | Write-Output
    } |
    Foreach-Object {
        $x=$_
        $combinedHa = [TimeSpan]::FromSeconds(
            ($x."Total BHS_Ha").TotalSeconds + ($x."Total Ha").TotalSeconds)
        $combinedOiii = [TimeSpan]::FromSeconds(
            ($x."Total BHS_Oiii").TotalSeconds + ($x."Total Oiii").TotalSeconds)
        $combinedSii = [TimeSpan]::FromSeconds(
            ($x."Total BHS_Sii").TotalSeconds + ($x."Total Sii").TotalSeconds)
        if($combinedHa -gt [TimeSpan]::Zero){
            Add-Member -InputObject $x -Value $combinedHa -MemberType NoteProperty -Name "Combined Ha"
        }
        if($combinedOiii -gt [TimeSpan]::Zero){
            Add-Member -InputObject $x -Value $combinedOiii -MemberType NoteProperty -Name "Combined Oiii"
        }
        if($combinedSii -gt [TimeSpan]::Zero){
            Add-Member -InputObject $x -Value $combinedSii -MemberType NoteProperty -Name "Combined Sii"
        }
        $totalSeconds=0
        $filters | foreach-object{
            $totalSeconds+=$x."Total $_".TotalSeconds
        }
        Add-Member -InputObject $x -Value ([TimeSpan]::FromSeconds($totalSeconds)) -MemberType NoteProperty -Name "Total Combined"
        Add-Member -InputObject $x -Value ([TimeSpan]::FromSeconds($totalSeconds).TotalMinutes) -MemberType NoteProperty -Name "Total Combined Minutes"
        $x
    }

$stats|    Format-Table Object,FocalLength,"Combined Ha","Combined Oiii","Combined Sii","Total L3","Total L","Total R","Total G","Total B","Total Combined"

$minDate=($data|Measure-Object ObsDateMinus12hr -Minimum).Minimum
$maxDate=($data|Measure-Object ObsDateMinus12hr -Maximum).Maximum
$byDate = $data | Group-Object {[int]$_.ObsDateMinus12hr.ToString("yyyyMMdd")}

Function Get-DateRange([DateTime]$StartDate,[DateTime]$EndDate){
    $x=$StartDate
    while($x -le $EndDate){
        Write-Output $x
        $x=$x.AddDays(1.0)
    }
}

$total = [TimeSpan]::FromMinutes(
    ($stats|Measure-Object "Total Combined Minutes" -Sum).Sum)
$total.TotalHours.ToString("00")+":"+$total.Minutes.ToString("00")