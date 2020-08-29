import-module $PSScriptRoot/PsXisfReader.psd1 -Force
$ErrorActionPreference="STOP"

Clear-Host

$data =
    @()+
    (Get-ChildItem "E:\Astrophotography\50mm" -Directory ) +
    (Get-ChildItem "E:\Astrophotography\135mm" -Directory ) +
    (Get-ChildItem "D:\Backups\Camera\2019" -Directory ) +
    (Get-ChildItem "E:\Astrophotography\1000mm" -Directory ) |
    ForEach-Object {
        Get-XisfLightFrames -Path $_ -Recurse -UseCache -SkipOnError 
    } |
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
        $totalSeconds=0
        $filters | foreach-object{
            $totalSeconds+=$x."Total $_".TotalSeconds
        }
        Add-Member -InputObject $x -Value ([TimeSpan]::FromSeconds($totalSeconds)) -MemberType NoteProperty -Name "Total Combined"
        Add-Member -InputObject $x -Value ([TimeSpan]::FromSeconds($totalSeconds).TotalMinutes) -MemberType NoteProperty -Name "Total Combined Minutes"
        $x
    }

$stats|Sort-Object "Total Combined Minutes" -Descending|    Format-Table Object,FocalLength,"Total Ha","Total Oiii","Total Sii","Total L","Total R","Total G","Total B","Total Combined"

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

Get-DateRange -StartDate $minDate -EndDate $maxDate |
    ForEach-Object {
        $date=$_
        $x = ($byDate |? Name -eq $date.ToString("yyyyMMdd"))

        new-object psobject -Property @{
            Date=$date.ToString("yyyy-MM-dd")
            ExposureTime=($x.Group|Measure-ExposureTime).TotalMinutes
        }    
    }
