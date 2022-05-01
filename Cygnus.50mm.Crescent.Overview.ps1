#import-module $PSScriptRoot/PsXisfReader.psd1 -Force
$ErrorActionPreference="STOP"

Clear-Host

$data =
    (Get-ChildItem "E:\Astrophotography\50mm" -Directory -Filter "PatchworkCygnus*") |
    #(Get-ChildItem "E:\Astrophotography\50mm" -Directory -Filter "Ceph*") +
    #(Get-ChildItem "E:\Astrophotography\135mm" -Directory -Filter "Sadr*") +
    # (Get-ChildItem "E:\Astrophotography\135mm" -Directory -Filter "Squid*") +
    # (Get-ChildItem "E:\Astrophotography\135mm" -Directory -Filter "PatchworkCygnus*") +
    # (Get-ChildItem "E:\Astrophotography\1000mm" -Directory -Filter "Sadr*") +
    # (Get-ChildItem "E:\Astrophotography\1000mm" -Directory -Filter "vdb131*") +
    # (Get-ChildItem "E:\Astrophotography\1000mm" -Directory -Filter "Sh 2-108*") +
    # (Get-ChildItem "E:\Astrophotography\1000mm" -Directory -Filter "Eye of*") +
    # (Get-ChildItem "E:\Astrophotography\1000mm" -Directory -Filter "Ring Nebula") +
     #(Get-ChildItem "E:\Astrophotography\135mm" -Directory -Filter "DWB*") |
    ForEach-Object {Get-XisfLightFrames -Path $_ -Recurse -UseCache } |
    Where-Object {-not $_.HasTokensInPath(@("reject","process","testing","clouds","draft","cloudy"))} |
    Where-Object {-not $_.IsIntegratedFile()} |
    Where-Object {$_.Filter} |
    Where-Object {$_.Filter -ne "L3"}

$filters = $data|Group-Object Filter|ForEach-Object{$_.Name}
$stats=$data|
    Group-Object FocalLength,Object,Filter,Exposure |
    ForEach-Object {
        $x = $_.Group
        $count=$_.Count
        $filter=$x[0].Filter
        $exposure=$x[0].Exposure
        $totalExposure=$x|Measure-ExposureTime -TotalSeconds
        new-object psobject -Property @{
            FocalLength=$x[0].FocalLength
            Object=$x[0].Object
            Filter=$filter
            Exposure=$x[0].Exposure
            Count=$count
            Exposures=$x
            Summary="$($count)x$($exposure)s"
            TotalExposure=$totalExposure.Total
            TotalExposureS=$totalExposure.TotalSeconds
        }
    } |
    Group-Object FocalLength,Object |
    Foreach-Object {
        $x = $_.Group
        $r = new-object psobject -Property @{
            Object = $x[0].Object
            FocalLength=$x[0].FocalLength
        }
        $filters|foreach-object {
            $filter = $_
            $subs = $x | where-object Filter -eq $filter | foreach-object {$_.Exposures} | sort-object ObsDate
            Add-Member -MemberType NoteProperty -InputObject $r -Value $subs  -Name "All $filter"

            $total = $subs|Measure-ExposureTime -First -Last
            $first = $total.First.ObsDate
            $last = $total.Last.ObsDate
            if($total.Total -eq [TimeSpan]::Zero) {
                $total = $null
            }
            Add-Member -MemberType NoteProperty -InputObject $r -Value $total.Total -Name "Total $filter"
            Add-Member -MemberType NoteProperty -InputObject $r -Value $first -Name "First $filter"
            Add-Member -MemberType NoteProperty -InputObject $r -Value $last  -Name "Last $filter"
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

$stats|    Format-Table Object,"Combined Ha","Combined Oiii","Combined Sii","Total Combined"
$stats|    Format-Table Object,"Last BHS_Ha","Last Ha","Last BHS_Oiii","Last BHS_Sii","Total Combined"

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