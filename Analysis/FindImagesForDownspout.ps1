$searchResults = Get-ChildItem D:\Backups\Camera -File -Recurse |
Where-Object {
    ($_.LastWriteTime -gt "2017-09-01") -and
    ($_.LastWriteTime -lt "2018-03-01")
}
$searchResults.Count

$searchResults |
    where-object {$ext = $_.Extension.ToUpper(); ($ext -eq ".NEF") -or ($ext -eq ".JPG")} |
    foreach-object {$x = $_; $x.Directory.FullName} |
    #where-object {-not $_.Contains("Astrophotography")} |
    #where-object {-not $_.Contains("ASI071mc-cool")} |
    #where-object {-not $_.Contains("PHD2")} |
    #where-object {-not $_.Contains("\SC\")} |
    #where-object {-not $_.Contains("\SGP\")} |
    select-object -Unique