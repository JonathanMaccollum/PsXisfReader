#import-module $PSScriptRoot/PsXisfReader.psd1 -Force
$ErrorActionPreference="STOP"

#Clear-Host

$1000mmdata =
    (Get-ChildItem "E:\Astrophotography\1000mm\Eye of Smaug in Cygnus" -Directory -Filter "*") +
    (Get-ChildItem "E:\Astrophotography\1000mm\Eye of Smaug Panel 2 OSC" -Directory -Filter "*") +
    (Get-ChildItem "E:\Astrophotography\1000mm\Eye of Smaug in Cygnus Panel 0" -Directory -Filter "*") +
    (Get-ChildItem "E:\Astrophotography\1000mm\Eye of Smaug in Cygnus Panel 2" -Directory -Filter "*") +
    (Get-ChildItem "E:\Astrophotography\1000mm\Eye of Smaug in Cygnus Panel -1" -Directory -Filter "*") +
    (Get-ChildItem "E:\Astrophotography\1000mm\Eye of Smaug in Cygnus Panel 3" -Directory -Filter "*") +
    (Get-ChildItem "E:\Astrophotography\1000mm\Eye of Smaug in Cygnus Panel 4" -Directory -Filter "*")  +
    (Get-ChildItem "E:\Astrophotography\1000mm\Eye of Smaug in Cygnus Panel 5" -Directory -Filter "*") +
    (Get-ChildItem "E:\Astrophotography\1000mm\Eastern Smaug in Cygnus Panel -1" -Directory -Filter "*")  +
    (Get-ChildItem "E:\Astrophotography\1000mm\Eastern Smaug in Cygnus Panel 0" -Directory -Filter "*")  +
    (Get-ChildItem "E:\Astrophotography\1000mm\Eastern Smaug in Cygnus Panel 1" -Directory -Filter "*")  +
    (Get-ChildItem "E:\Astrophotography\1000mm\Eastern Smaug in Cygnus Panel 2" -Directory -Filter "*")  +
    (Get-ChildItem "E:\Astrophotography\1000mm\Eastern Smaug in Cygnus Panel 3" -Directory -Filter "*")  +
    (Get-ChildItem "E:\Astrophotography\1000mm\Eastern Smaug in Cygnus Panel 4" -Directory -Filter "*")  +
    (Get-ChildItem "E:\Astrophotography\1000mm\Eastern Smaug in Cygnus Panel 5" -Directory -Filter "*")  +
    (Get-ChildItem "E:\Astrophotography\1000mm\Western Smaug in Cygnus Panel -1" -Directory -Filter "*")  +
    (Get-ChildItem "E:\Astrophotography\1000mm\Western Smaug in Cygnus Panel 0" -Directory -Filter "*")  +
    (Get-ChildItem "E:\Astrophotography\1000mm\Western Smaug in Cygnus Panel 1" -Directory -Filter "*")  +
    (Get-ChildItem "E:\Astrophotography\1000mm\Western Smaug in Cygnus Panel 2" -Directory -Filter "*")  +
    (Get-ChildItem "E:\Astrophotography\1000mm\Western Smaug in Cygnus Panel 3" -Directory -Filter "*")  +
    (Get-ChildItem "E:\Astrophotography\1000mm\Western Smaug in Cygnus Panel 4" -Directory -Filter "*")  +
    (Get-ChildItem "E:\Astrophotography\1000mm\Western Smaug in Cygnus Panel 5" -Directory -Filter "*")  +
    #(Get-ChildItem "E:\Astrophotography\1000mm\IC1396 Elephants Trunk Nebula" -Directory -Filter "*") +
    #(Get-ChildItem "E:\Astrophotography\1000mm\IC1871" -Directory -Filter "*") +
    #(Get-ChildItem "E:\Astrophotography\1000mm\Lobster Claw in Cepheus" -Directory -Filter "*") +
    #(Get-ChildItem "E:\Astrophotography\1000mm\Sh2-132 - Lion Nebula" -Directory -Filter "*") +
    #(Get-ChildItem "E:\Astrophotography\1000mm\vdB131 vdB132" -Directory -Filter "*") +
    #(Get-ChildItem "E:\Astrophotography\1000mm\vdB131 vdB132 Panel2" -Directory -Filter "*") +
    #(Get-ChildItem "E:\Astrophotography\1000mm\Cave Nebula OSC" -Directory -Filter "*") +
    #(Get-ChildItem "E:\Astrophotography\1000mm\IC1871" -Directory -Filter "*") +
    #(Get-ChildItem "E:\Astrophotography\1000mm\Tulip Panel 1" -Directory -Filter "*") +
    #(Get-ChildItem "E:\Astrophotography\1000mm\Tulip Panel 2" -Directory -Filter "*") +
    #(Get-ChildItem "E:\Astrophotography\1000mm\Tulip Panel 3" -Directory -Filter "*") +
    #(Get-ChildItem "E:\Astrophotography\1000mm\Tulip Panel 4" -Directory -Filter "*") +
    (Get-ChildItem "E:\Astrophotography\1000mm\wr134" -Directory -Filter "*") +
    #(Get-ChildItem "E:\Astrophotography\1000mm\Sh 2-108" -Directory -Filter "*") +
    #(Get-ChildItem "E:\Astrophotography\1000mm\sh2-86" -Directory -Filter "*") +
    #(Get-ChildItem "E:\Astrophotography\1000mm\sh2-115" -Directory -Filter "*") +
    #(Get-ChildItem "E:\Astrophotography\1000mm\NGC7000 Framing 2" -Directory -Filter "*") +
    @()|
    ForEach-Object {Get-XisfLightFrames -Path $_ -Recurse -UseCache -SkipOnError } |
    Where-Object {-not $_.HasTokensInPath(@("reject","process","testing","clouds","draft","cloudy","_ez_LS_","quickedit","integration","drizzle"))} |
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
    Where-Object Instrument -ne "ZWO ASI071MC Pro" |
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
#Sort-Object "Total Oiii" |
Format-Table `
    Object,"Total Ha","Total Oiii","Total Oiii6nm","Total Sii3","Total Sii6nm","Total L3","Total R","Total G","Total B","Total Combined"

[TimeSpan]::FromMinutes(  ($stats | Measure-Object {$_.'Total Combined Minutes'} -Sum).Sum)

<#
$stats|
    Where-Object Instrument -eq "ZWO ASI071MC Pro" |
    Format-Table `
        Object,"Total Ha","Total Oiii6nm","Total Sii6nm","Total L3","Total R","Total G","Total B","Total Combined"
        #>    