cls
Join-Path ([Environment]::GetEnvironmentVariable("LocalAppData")) "NINA\Logs" |
get-childitem -filter *.log |
Sort-Object LastWriteTime -Descending |
Select-Object -First 20 |
get-content | 
where-object {$_.Contains("Sensor humidity:")} |
where-object {-not $_.Contains("NaN%")} | 
foreach-object {
    $tokens=$_.Split("|")
    new-object psobject -Property @{
        TimeStamp = ([DateTime]$tokens[0])
        Humidity = $tokens[$tokens.Count-1].Split(":")[2].Trim()
    }
} | export-csv -Path (Join-Path ([Environment]::GetEnvironmentVariable("LocalAppData")) "NINA\Logs\Humidity.csv") -Delimiter ","



#join-path ([Environment]::GetEnvironmentVariable("LocalAppData")) "NINA\Logs" |
"D:\Backups\Camera\Dropoff\NINA\Logs" |
get-childitem -filter *.log |
Sort-Object LastWriteTime -Descending |
Select-Object -First 50 |
get-content | 
where-object {$_.Contains("Sensor humidity:")} |
where-object {-not $_.Contains("NaN%")} | 
foreach-object {
    $tokens=$_.Split("|")
    new-object psobject -Property @{
        TimeStamp = ([DateTime]$tokens[0])
        Humidity = $tokens[$tokens.Count-1].Split(":")[2].Trim()
    }
} | 
sort-object TimeStamp |
export-csv -Path "SomeLocation\Humidity.csv" -Delimiter "," -Force