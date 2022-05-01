#import-module $PSScriptRoot/PsXisfReader.psd1 -Force
$ErrorActionPreference="STOP"

#Clear-Host

$40mmdata =
    (Get-ChildItem "E:\Astrophotography\40mm\Cygnus Panel 1" -Directory -Filter "*") +
    (Get-ChildItem "E:\Astrophotography\40mm\Cygnus Panel 2" -Directory -Filter "*") +
    (Get-ChildItem "E:\Astrophotography\40mm\Cygnus Panel 3" -Directory -Filter "*") +
    (Get-ChildItem "E:\Astrophotography\40mm\Cepheus on Lobster Claw" -Directory -Filter "*") +
    (Get-ChildItem "E:\Astrophotography\40mm\Cepheus on Lion" -Directory -Filter "*") +
    (Get-ChildItem "E:\Astrophotography\40mm\Cassiopeia near 12 Cas" -Directory -Filter "*") +
    (Get-ChildItem "E:\Astrophotography\40mm\Cassiopeia on HR561" -Directory -Filter "*") +
    (Get-ChildItem "E:\Astrophotography\40mm\Orion on Mu Ori" -Directory -Filter "*") +
    (Get-ChildItem "E:\Astrophotography\40mm\Orion on NGC1788" -Directory -Filter "*") +
    (Get-ChildItem "E:\Astrophotography\40mm\California on 42Per" -Directory -Filter "*") +
    (Get-ChildItem "E:\Astrophotography\40mm\Perseus on Theta Persei" -Directory -Filter "*") +
    (Get-ChildItem "E:\Astrophotography\40mm\Sh2-126 in Lacerta" -Directory -Filter "*") +
    (Get-ChildItem "E:\Astrophotography\40mm\Sh2-240" -Directory -Filter "*") +
    @()|
    ForEach-Object {Get-XisfLightFrames -Path $_ -Recurse -UseCache -SkipOnError } |
    Where-Object {-not $_.HasTokensInPath(@("reject","process","testing","clouds","draft","cloudy","_ez_LS_","quickedit","integration","drizzle"))} |
    Where-Object {-not $_.IsIntegratedFile()} |
    Where-Object {$_.Filter}

$filters = $40mmdata|Group-Object Filter|ForEach-Object{$_.Name}
$stats=$40mmdata|
    Group-Object FocalLength,Object,Filter,Exposure,Instrument |
    ForEach-Object {
        $x = $_.Group
        $count=$_.Count
        $filter=$x[0].Filter
        $exposure=$x[0].Exposure
        $instrument=$x[0].Instrument
        $totalExposure=$x|Measure-ExposureTime -TotalSeconds
        new-object psobject -Property @{
            FocalLength=$x[0].FocalLength
            Object=$x[0].Object.Replace(" in Cygnus","").Replace(" Panel "," ")
            Filter=$filter
            Exposure=$exposure
            Instrument=$instrument
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
<#
$stats|
Where-Object Object -eq "Sh 2-108" |
Format-List `
    Object,"Last Ha", "Last Sii6nm", "Last Oiii6nm", "Last L3"
    
    #>
$stats|
Sort-Object "Total Sii3" |
Format-Table `
    "Object","Total Ha","Total Oiii","Total Sii3","Total Combined"

#[TimeSpan]::FromMinutes(  ($stats | Measure-Object {$_.'Total Combined Minutes'} -Sum).Sum)

<#
$stats|
    Where-Object Instrument -eq "ZWO ASI071MC Pro" |
    Format-Table `
        Object,"Total Ha","Total Oiii6nm","Total Sii6nm","Total L3","Total R","Total G","Total B","Total Combined"
        #>    