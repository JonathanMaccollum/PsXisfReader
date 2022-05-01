$results = 
    Get-ChildItem F:\PixInsightLT\Calibrated -Directory |
    foreach-object {
        $targetDir = $_
        $byLastWriteTime = Get-ChildItem $targetDir.FullName -File -Recurse -Filter *.xisf | sort-object LastWriteTime
        $sizeOfChildren = $byLastWriteTime | Measure-Object -Sum Length
        $firstChild = $byLastWriteTime | Select-Object -First 1
        $lastChild = $byLastWriteTime | Select-Object -Last 1
        new-object psobject -Property @{
            Target=$targetDir.Name
            Size = $sizeOfChildren.Sum
            Count = $sizeOfChildren.Count
            EarliestDate = $firstChild.LastWriteTime
            LatestDate = $lastChild.LastWriteTime
        }
    }
$results | 
    sort-object Size -Descending |
    select-object -first 100 |
    sort-object LatestDate |
    format-table Target, Size, Count, EarliestDate, LatestDate