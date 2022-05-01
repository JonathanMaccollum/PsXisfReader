Clear-Host
$leadingFilter = "Ha3nm"
$data = Get-ChildItem "D:\Backups\Camera\Dropoff\NINA\AutoFocus" |
Where-Object CreationTime -gt "2022-02-18 23:00:00" |
Where-Object CreationTime -lt "2022-02-19 16:00:00" |
sort-object CreationTime |
#select-object -First 120 |
foreach-object {
    $c=get-content $_ | ConvertFrom-Json
    $c|add-member Position $c.CalculatedFocusPoint.Position -PassThru
}

$data | group-object Filter |
foreach-object {
    $filter = $_.Name
    $set = $_.Group
    $temp = $set | Measure-Object  Temperature -Minimum -Maximum
    $position = $set | Measure-Object Position -Minimum -Maximum
    $tempDiff = $temp.Maximum-$temp.Minimum
    $positionDiff = $position.Maximum - $position.Minimum
    $slope = $positionDiff/$tempDiff
    $runCount = $set.Count
    new-object psobject -property @{
        Filter = $filter
        RunCount = $runCount
        Slope = $slope
        ChangeInTemp = $tempDiff
        ChangeInPosition = $positionDiff
    }
    #"$($filter.PadRight(8)): $positionDiff steps in $tempDiff degrees = $($slope) steps/deg. Runs: $runs"
} |
#sort-object Filter |
format-table Filter,RunCount,ChangeInTemp,ChangeInPosition,Slope

$mostRecentLeaderRun = $null
$offsetsMeasured = $data |
    sort-object Timestamp |
    foreach-object {
        if($_.Filter -eq $leadingFilter){
            if($null -ne $mostRecentLeaderRun){
                $offset = $_.Position - $mostRecentLeaderRun.Position
                new-object psobject -property @{
                    Timestamp = $_.Timestamp
                    Filter = $_.Filter
                    LeaderPosition = $mostRecentLeaderRun.Position
                    Position = $_.Position
                    Offset = $offset
                    TimeSinceLeaderRun = $_.TimeStamp-$mostRecentLeaderRun.TimeStamp
                }
            }
            $mostRecentLeaderRun = $_
        }
        elseif($null -eq $mostRecentLeaderRun){
            write-warning "Skipping run of filter $($_.Filter). No leader information present."
        }
        else{
            $offset = $_.Position - $mostRecentLeaderRun.Position
            new-object psobject -property @{
                Timestamp = $_.Timestamp
                Filter = $_.Filter
                LeaderPosition = $mostRecentLeaderRun.Position
                Position = $_.Position
                Offset = $offset
                TimeSinceLeaderRun = $_.TimeStamp-$mostRecentLeaderRun.TimeStamp
            }
        }
    } 
$offsetsMeasured|
    sort-object Timestamp |
    format-table Timestamp,Filter,Position,Offset,TimeSinceLeaderRun
$offsetsMeasured | 
    group-object Filter |
    foreach-object {
        $filter = $_.Name
        $values = $_.Group
        $stats = $values | Measure-Object Offset -Maximum -Minimum -Average -StandardDeviation
        add-member -inputObject $stats -NotePropertyName "Filter" -NotePropertyValue $filter -PassThru
    } |
    foreach-object {
        add-member -inputObject $_ -NotePropertyName "DriftAdjusted" -NotePropertyValue ($_.Average-106) -PassThru
    } |
    format-table Filter,DriftAdjusted,Count,Minimum,Maximum,Average,StandardDeviation