#import-module $PSScriptRoot/PsXisfReader.psd1 -Force
$ErrorActionPreference="STOP"

Clear-Host

$1000mmdata =
    @()+
    (Get-ChildItem "E:\Astrophotography\1000mm\Eastern Smaug in Cygnus Panel -1" -Directory -Filter "*") +
    (Get-ChildItem "E:\Astrophotography\1000mm\Eastern Smaug in Cygnus Panel 0" -Directory -Filter "*") +
    (Get-ChildItem "E:\Astrophotography\1000mm\Eastern Smaug in Cygnus Panel 1" -Directory -Filter "*") +
    (Get-ChildItem "E:\Astrophotography\1000mm\Eastern Smaug in Cygnus Panel 2" -Directory -Filter "*") +
    (Get-ChildItem "E:\Astrophotography\1000mm\Eastern Smaug in Cygnus Panel 3" -Directory -Filter "*") +
    (Get-ChildItem "E:\Astrophotography\1000mm\Eastern Smaug in Cygnus Panel 4" -Directory -Filter "*") +
    (Get-ChildItem "E:\Astrophotography\1000mm\Eastern Smaug in Cygnus Panel 5" -Directory -Filter "*") +
    
    (Get-ChildItem "E:\Astrophotography\1000mm\Western Smaug in Cygnus Panel -1" -Directory -Filter "*") +
    (Get-ChildItem "E:\Astrophotography\1000mm\Western Smaug in Cygnus Panel 0" -Directory -Filter "*") +
    (Get-ChildItem "E:\Astrophotography\1000mm\Western Smaug in Cygnus Panel 1" -Directory -Filter "*") +
    (Get-ChildItem "E:\Astrophotography\1000mm\Western Smaug in Cygnus Panel 2" -Directory -Filter "*") +
    (Get-ChildItem "E:\Astrophotography\1000mm\Western Smaug in Cygnus Panel 3" -Directory -Filter "*") +
    (Get-ChildItem "E:\Astrophotography\1000mm\Western Smaug in Cygnus Panel 4" -Directory -Filter "*") +
    (Get-ChildItem "E:\Astrophotography\1000mm\Western Smaug in Cygnus Panel 5" -Directory -Filter "*") +
    

     (Get-ChildItem "E:\Astrophotography\1000mm\Eye of Smaug in Cygnus" -Directory -Filter "*") +
     (Get-ChildItem "E:\Astrophotography\1000mm\Eye of Smaug Panel 2 OSC" -Directory -Filter "*") +
     (Get-ChildItem "E:\Astrophotography\1000mm\Eye of Smaug in Cygnus Panel 0" -Directory -Filter "*") +
     (Get-ChildItem "E:\Astrophotography\1000mm\Eye of Smaug in Cygnus Panel -1" -Directory -Filter "*") +
    #(Get-ChildItem "E:\Astrophotography\1000mm\Eye of Smaug Panel 2" -Directory -Filter "*") +
     (Get-ChildItem "E:\Astrophotography\1000mm\Eye of Smaug in Cygnus Panel 3" -Directory -Filter "*") +
     (Get-ChildItem "E:\Astrophotography\1000mm\Eye of Smaug in Cygnus Panel 4" -Directory -Filter "*")  +
     (Get-ChildItem "E:\Astrophotography\1000mm\Eye of Smaug in Cygnus Panel 5" -Directory -Filter "*")  +
     (Get-ChildItem "E:\Astrophotography\1000mm\vdB131 vdB132" -Directory -Filter "*")  +
     (Get-ChildItem "E:\Astrophotography\1000mm\vdB131 vdB132 Panel2" -Directory -Filter "*")  +
    #(Get-ChildItem "E:\Astrophotography\1000mm\NGC5905-5908 Panel 1" -Directory -Filter "*")  +
    #(Get-ChildItem "E:\Astrophotography\1000mm\NGC5905-5908 Panel 2" -Directory -Filter "*")  +
    #(Get-ChildItem "E:\Astrophotography\1000mm\Coddingtons" -Directory -Filter "*")  +
    #(Get-ChildItem "E:\Astrophotography\1000mm\Coddington's Nebula" -Directory -Filter "*")  +
    #(Get-ChildItem "E:\Astrophotography\1000mm\Coddingtons OSC" -Directory -Filter "*")  +
    #(Get-ChildItem "E:\Astrophotography\1000mm\Owl Nebula" -Directory -Filter "*")  +
    #(Get-ChildItem "E:\Astrophotography\1000mm\Jellyfish Exhaust" -Directory -Filter "*")  +
    #(Get-ChildItem "E:\Astrophotography\1000mm\Abell 39" -Directory -Filter "*")  +
    #(Get-ChildItem "E:\Astrophotography\1000mm\Abell 31" -Directory -Filter "*")  +
    # (Get-ChildItem "E:\Astrophotography\1000mm\Sh 2-108" -Directory -Filter "*")  +
    #(Get-ChildItem "E:\Astrophotography\1000mm\M38" -Directory -Filter "*")  +
    @() |
    ForEach-Object {Get-XisfLightFrames -Path $_ -Recurse -UseCache -SkipOnError } |
    Where-Object {-not $_.HasTokensInPath(@("reject","process","testing","clouds","draft","cloudy","quickedit","integration"))} |
    Where-Object {-not $_.IsIntegratedFile()} |
    Where-Object {$_.Filter}

$filters = $1000mmdata|Group-Object Filter|ForEach-Object{$_.Name}
$stats=$1000mmdata|
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
            Object=$x[0].Object
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
    Group-Object FocalLength,Object,Instrument |
    Foreach-Object {
        $x = $_.Group
        $r = new-object psobject -Property @{
            Object = $x[0].Object
            FocalLength=$x[0].FocalLength
            Instrument=$x[0].Instrument
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

$stats|    Format-Table Object,Instrument,FocalLength,"Combined Ha","Combined Oiii","Combined Sii","Total Ha","Total Oiii","Total Oiii6nm","Total Sii","Total Sii6nm","Total L3","Total L", "Total D1","Total R","Total G","Total B","Total Combined"
#$stats|    Format-Table Object,FocalLength,"Last BHS_Ha","Last Ha","Last BHS_Oiii","Last Oiii","Last Sii","Last L3","Last L","Last R","Last G","Last B","Total Combined"
<#
$minDate=($1000mmdata|Measure-Object ObsDateMinus12hr -Minimum).Minimum
$maxDate=($1000mmdata|Measure-Object ObsDateMinus12hr -Maximum).Maximum
$byDate = $1000mmdata | Group-Object {[int]$_.ObsDateMinus12hr.ToString("yyyyMMdd")}

Function Get-DateRange([DateTime]$StartDate,[DateTime]$EndDate){
    $x=$StartDate
    while($x -le $EndDate){
        Write-Output $x
        $x=$x.AddDays(1.0)
    }
}
#>
$total = [TimeSpan]::FromMinutes(
    ($stats|Measure-Object "Total Combined Minutes" -Sum).Sum)
$total.TotalHours.ToString("00")+":"+$total.Minutes.ToString("00")
<#
$1000mmdata |   
    group-object {$_.ObsDateMinus12hr.ToString("yyyyMM")+":"+$_.Filter+":"+$_.Instrument+":"+$_.Gain+":"+$_.Offset} | 
    foreach-object { 
        $x=$_.Group
        $et=$x|Measure-ExposureTime -TotalHours
        new-object psobject -Property @{Name=$_.Name; Stats=$et }
    }
    #>