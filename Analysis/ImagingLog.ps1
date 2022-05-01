#import-module PsXisfReader
$WarningPreference="Continue"
$Ignorables = @("reject","process","testing","clouds","draft","cloudy","_ez_ls_","quickedit","draft","\flats\","pipp","solved")
$rawStats = 
    @(
        Get-XisfLightFrames -Path "E:\Astrophotography\40mm" -UseCache -UseErrorCache -SkipOnError -Recurse -PathTokensToIgnore $Ignorables -ShowProgress
        Get-XisfLightFrames -Path "E:\Astrophotography\50mm" -UseCache -UseErrorCache -SkipOnError -Recurse -PathTokensToIgnore $Ignorables -ShowProgress
        Get-XisfLightFrames -Path "E:\Astrophotography\135mm" -UseCache -UseErrorCache -SkipOnError -Recurse -PathTokensToIgnore $Ignorables -ShowProgress
        Get-XisfLightFrames -Path "E:\Astrophotography\1000mm" -UseCache -UseErrorCache -SkipOnError -Recurse -PathTokensToIgnore $Ignorables -ShowProgress
    )|
    Where-Object {-not $_.IsIntegratedFile()} |
    where-object {$_.ObsDateMinus12hr}
    #Where-Object {$_.ObsDateMinus12hr.Year -eq 2021}
    #Where-Object {$_.ObsDateMinus12hr -gt [datetime]::Today.AddDays(-365)}
exit
$byDate = $rawStats |
    Group-Object ObsDateMinus12hr
$byDate | 
    where-object {([DateTime]$_.Name) -ge ([DateTime]"2021-01-01")} |
    foreach-object {
        new-object psobject -Property @{
        Date = ([DateTime]$_.Name)
        Total = ($_.Group | Measure-ExposureTime).Total
        }
    } |
    group-object {$_.Date.Month} |
    foreach-object {
        $totalHours = [TimeSpan]::FromHours(
            ($_.Group | Measure-Object {$_.Total.TotalHours} -Sum).Sum).TotalHours
        "$($_.Name): $($totalHours.ToString('000.00'))"
    }
$byImageScale = $rawStats|Group-Object {$_|Get-XisfImageScale}

<#
$byGear = $rawStats | Group-Object Instrument,FocalLength,Filter|foreach-object{
    $instrument=$_.Group[0].Instrument
    $focalLength=$_.Group[0].FocalLength
    $filter=$_.Group[0].Filter
    $gearStats = $_.Group|Measure-ExposureTime -TotalHours
    new-object psobject -Property @{
        Instrument=$instrument
        FocalLength=$focalLength
        Filter=$filter
        GearStats=$gearStats
        TotalHours=$gearStats.TotalHours
        Data = $_.Group
    }
}
$byGear|Sort-Object TotalHours|Format-Table Instrument,FocalLength,Filter,TotalHours
#>
Write-Host "Session Stats by Day"
$byDate |
    foreach-object{
        $date = $_.Group[0].ObsDateMinus12hr
        $sessionStats = $_.Group|Measure-ExposureTime -First -Last -TotalHours
        $_.Group|Group-Object Object,Instrument,FocalLength,Filter|foreach-object{
            $instrument=$_.Group[0].Instrument
            $focalLength=$_.Group[0].FocalLength
            $filter=$_.Group[0].Filter
            $object=$_.Group[0].Object
            $gearStats = $_.Group|Measure-ExposureTime -TotalHours
            new-object psobject -Property @{
                Date=$date
                Instrument=$instrument
                FocalLength=$focalLength
                Filter=$filter
                Object=$object
                TotalHours=$gearStats.TotalHours
                SessionStartTime=[DateTime]::SpecifyKind(
                    $sessionStats.First.ObsDate,
                    [DateTimeKind]::Utc).ToLocalTime()
                SessionEndTime=[DateTime]::SpecifyKind(
                    $sessionStats.Last.ObsDate,
                    [DateTimeKind]::Utc).ToLocalTime()
                SessionTotalHours=$sessionStats.TotalHours.ToString("00.0")
                SessionMoonPhase = 100*(Get-MoonPercentIlluminated -Date $date)
            }
        }
    } |
    #Where-Object {$_.Object.ToLower().Contains("1333")} |
    Format-Table Date,Object,FocalLength,Instrument,Filter,TotalHours#,SessionStartTime,SessionEndTime,SessionTotalHours,SessionMoonPhase

Write-Host "Session Stats by Target"
$rawStats|
    #Where-Object FocalLength -eq 1000 |
    #Where-Object Filter -eq "L3" |
    Where-Object {$_.Object.ToLower().Contains("codd")} |
    Group-Object Object|
    foreach-object{
        $object = $_.Group[0].Object
        $objectStats = $_.Group|Measure-ExposureTime -First -Last -TotalHours
        $_.Group | 
        group-object ObsDateMinus12hr | 
        ForEach-Object {
            $date = $_.Group[0].ObsDateMinus12hr
            $sessionStats = $_.Group|Measure-ExposureTime -First -Last -TotalHours
            $_.Group|Group-Object Object,Instrument,FocalLength,Filter|foreach-object{
                $instrument=$_.Group[0].Instrument
                $focalLength=$_.Group[0].FocalLength
                $filter=$_.Group[0].Filter
                $object=$object
                $gearStats = $_.Group|Measure-ExposureTime -TotalHours
                new-object psobject -Property @{
                    Date=$date
                    Instrument=$instrument
                    FocalLength=$focalLength
                    Filter=$filter
                    Object=$object
                    SessionHours = $sessionStats.TotalHours
                    TotalHours=$gearStats.TotalHours
                    ObjectTotalHours=$objectStats.TotalHours
                    SessionStartTime=[DateTime]::SpecifyKind(
                        $sessionStats.First.ObsDate,
                        [DateTimeKind]::Utc).ToLocalTime()
                    SessionEndTime=[DateTime]::SpecifyKind(
                        $sessionStats.Last.ObsDate,
                        [DateTimeKind]::Utc).ToLocalTime()
                    SessionTotalHours=$sessionStats.TotalHours.ToString("00.0")
                    SessionMoonPhase = 100*(Get-MoonPercentIlluminated -Date $date)
                }
            }
        } 
    } |
    Sort-Object Object,Date,FocalLength,Filter |
    Format-Table Object,Date,FocalLength,Filter,SessionTotalHours,TotalHours
    

