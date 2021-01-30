import-module $PSScriptRoot/PsXisfReader.psd1 -Force
$WarningPreference="SilentlyContinue"
$byDate = 
    (Get-XisfLightFrames -Path "E:\Astrophotography\135mm" -UseCache -SkipOnError -Recurse) +
    (Get-XisfLightFrames -Path "E:\Astrophotography\50mm" -UseCache -SkipOnError -Recurse) +
    (Get-XisfLightFrames -Path "E:\Astrophotography\1000mm" -UseCache -SkipOnError -Recurse) |
    Where-Object {-not $_.HasTokensInPath(@("reject","process","testing","clouds","draft","cloudy","_ez_ls_","quickedit","draft"))} |
    Where-Object {-not $_.IsIntegratedFile()} |
    Where-Object {$_.ObsDateMinus12hr -gt [datetime]::Today.AddDays(-180)} |
    Group-Object ObsDateMinus12hr

$stats=$byDate|
    foreach-object{
        $date = $_.Group[0].ObsDateMinus12hr
        $sessionStats = $_.Group|Measure-ExposureTime -First -Last -TotalHours
        $byGearStats = $_.Group|Group-Object Instrument,FocalLength,Object,Filter|foreach-object{
            $instrument=$_.Group[0].Instrument
            $focalLength=$_.Group[0].FocalLength
            $filter=$_.Group[0].Filter
            $object=$_.Group[0].Object
            $gearStats = $_.Group|Measure-ExposureTime -TotalHours
            new-object psobject -Property @{
                Instrument=$instrument
                FocalLength=$focalLength
                Filter=$filter
                Object=$object
                GearStats=$gearStats
            }
        }
        new-object psobject -Property @{
            Date=$date
            SessionStats=$sessionStats
            ByGear=$byGearStats
        }
    } |
    foreach-object {
        new-object psobject -Property @{
            Date=$_.Date
            StartTime=[DateTime]::SpecifyKind(
                $_.SessionStats.First.ObsDate,
                [DateTimeKind]::Utc).ToLocalTime()
            EndTime=[DateTime]::SpecifyKind(
                $_.SessionStats.Last.ObsDate,
                [DateTimeKind]::Utc).ToLocalTime()
            TotalHours=$_.SessionStats.TotalHours.ToString("00.0")
            MoonPhase = 100*(Get-MoonPercentIlluminated -Date $_.Date)
        }
    }
Write-Host "Session Stats by Day"
$stats|Format-Table Date,StartTime,EndTime,TotalHours,MoonPhase