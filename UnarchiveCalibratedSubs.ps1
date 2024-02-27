import-module psxisfreader -Force

$LastNDays=380
$AnchorDate=[DateTime]::Today.AddDays(-$LastNDays)
$ArchivePath="M:\PixinsightLT\Calibrated"
$ArchivedTargets = @(
    ''
)
$activeTargets = @(40,90,950) |
    foreach-object {
        get-childitem "E:\Astrophotography\$($_)mm" -Directory
    } |
    where-object Name -ne "Flats" |
    where-object Name -ne "Planning" |
    where-object Name -eq "Cygnus on HD192143" |
    where-object { -not ([decimal]::TryParse($_.Name.TrimEnd("s"),[ref]$null))} |
    foreach-object {
        $target = $_
        $subs = (Get-XisfLightFrames -Path $target.FullName `
                -SkipOnError `
                -Recurse `
                -UseCache `
                -PathTokensToIgnore @("reject","process","testing","clouds","draft","cloudy","_ez_LS_","drizzle","skyflats","quickedit","calibrated")) |
            Where-Object {$_.LocalDate } |
            Where-Object {-not $_.IsIntegratedFile()}
        $subs|
            group-object Instrument,FocalLength | 
            foreach-object {
                $instrument = $_.Group[0].Instrument
                $focalLength = $_.Group[0].FocalLength
                $stats=$_.Group|Measure-ExposureTime -TotalHours -TotalMinutes -TotalSeconds -First -Last
                if(($_.Group | 
                    where-object {$_.LocalDate.Date -gt $AnchorDate}))
                {
                    new-object psobject -Property @{
                        Instrument=$instrument
                        FocalLength=$focalLength
                        TargetPath=$target
                        TargetName=$target.Name
                        LightFrames=$_.Group
                        Count=$stats.Count
                        TotalHours=$stats.TotalHours
                        Total=$stats.Total
                        First=$stats.First
                        Last=$stats.Last
                    }
                }
        }
    }
$activeTargets | 
    where-object TotalHours -gt 8 |
    format-table TargetName,Instrument,FocalLength,Count,TotalHours

$activeTargets |
foreach-object {
    $targetName=$_.TargetName
    $calibrationPath="E:\Calibrated\$($_.FocalLength)mm"
    $targetPath = join-path $calibrationPath $targetName
    [IO.Directory]::CreateDirectory($calibrationPath)>>$null
    [IO.Directory]::CreateDirectory($targetPath)>>$null
    $toArchiveDestinationDir = join-path $ArchivePath $targetName
    [IO.Directory]::CreateDirectory($toArchiveDestinationDir)>>$null
    $state=$_.LightFrames |
        Get-XisfCalibrationState `
            -CalibratedPath $calibrationPath `
            -Verbose
    $archiveState=$_.LightFrames |
        Get-XisfCalibrationState `
            -CalibratedPath (join-path $ArchivePath $targetName) `
            -Verbose
        
    for($i=0;$i -lt $_.LightFrames.Count; $i += 1){
            
        if($state[$i].Calibrated -and $archiveState[$i].Calibrated){
            continue;
        }
        if($archiveState[$i].Calibrated){
            Copy-Item $archiveState[$i].Calibrated.FullName -Destination $targetPath
        }
        elseif($state[$i].Calibrated){
            $toArchiveDestination = join-path $toArchiveDestinationDir ($state[$i].Calibrated.Name)
            Copy-Item $state[$i].Calibrated.FullName -Destination $toArchiveDestination
        }
        else{
            write-warning "No calibrated file was found for $($state[$i].Path)"
        }
    }
}

<#
Get-XisfCalibrationState `
    -CalibratedPath (join-path $ArchivePath $targetName) `
    -XisfFileStats ($state[$i].Stats) 
    #>