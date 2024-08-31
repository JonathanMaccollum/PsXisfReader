import-module PSConsumerPipeline
import-module PSSQLite
$VerbosePreference="SilentlyContinue"
$ErrorActionPreference="STOP"
Function Update-HashLibrary
{
    [CmdletBinding()]
    param(
        [Parameter()][IO.FileInfo]$DBFile,
        [Parameter()][IO.DirectoryInfo]$Path,
        [string]$Filter,
        [Switch]$RecreateDB
    )
    if($RecreateDB -and $DBFile.Exists){
        $DBFile.Delete()
    }
    write-verbose "Connecting to SQLite Database $($DBFile.Name)"
    $conn = New-SQLiteConnection -DataSource $DBFile
    try{
        Invoke-SqliteQuery `
            -Query "create table if not exists FileHashes(Path varchar(255) not null,FileHash char(64) not null)" `
            -SQLiteConnection $conn
        Invoke-SqliteQuery `
            -Query "create unique index if not exists UC_FileHashes on FileHashes(Path)" `
            -SQLiteConnection $conn
        write-verbose "Fetching existing file hashes"
        $existingHashes = Invoke-SqliteQuery `
            -Query "select * from FileHashes" `
            -SQLiteConnection $conn |
            Group-Object Path -AsHashTable
        write-verbose "Calculating files to process"
        $toProcess=
            get-childitem $Path -Recurse -filter:$Filter -file |
                where-object {
                    $x=$_
                    if(-not $existingHashes){
                        return $true
                    }
                    if(-not (test-path $x.FullName)){
                        return $false
                    }
                    $relativePath = Resolve-Path -Path $x.FullName -RelativeBasePath $Path -Relative
                    return -not ($existingHashes.ContainsKey($relativePath))
                }
        $i=0
        $toProcess |
            foreach-object {
                $x = $_
                $relativePath = Resolve-Path -Path $x.FullName -RelativeBasePath $Path -Relative
                $results = $null
                $results = Get-FileHash -Path $x -Algorithm SHA256 -ErrorAction:Continue
                if(-not $results){
                    write-warning "Failed to calculate hash for $($x.FullName)"
                    return
                }
                Invoke-SqliteQuery `
                    -Query "insert into FileHashes(Path,FileHash)values(@Path,@FileHash)" `
                    -SqlParameters @{
                        Path=$relativePath
                        FileHash=$results.Hash
                    } `
                    -SQLiteConnection $conn
                $i+=1
                if($i % 100 -eq 0){
                    write-progress -Activity "Calculating Hashes" -PercentComplete (100.0*$i/$toProcess.Count) -Status "Calculated $i of $($toProcess.Count) files."
                }
            }
        Write-Progress -Activity "Calculating Hashes" -Completed
    }
    finally{
        $conn.Dispose()
    }
}
Function Update-XisfHeaderLibrary
{
    [CmdletBinding()]
    param(
        [Parameter()][IO.FileInfo]$DBFile,
        [Parameter()][IO.DirectoryInfo]$Path,
        [string]$Filter,
        [Switch]$RecreateDB
    )
    if($RecreateDB -and $DBFile.Exists){
        $DBFile.Delete()
    }
    write-verbose "Connecting to SQLite Database $($DBFile.Name)"
    $conn = New-SQLiteConnection -DataSource $DBFile
    try{
        Invoke-SqliteQuery `
            -Query "create table if not exists XisfHeaders(Path varchar(255) not null,HeaderXml text not null)" `
            -SQLiteConnection $conn
        Invoke-SqliteQuery `
            -Query "create unique index if not exists UC_XisfHeaders on XisfHeaders(Path)" `
            -SQLiteConnection $conn
        write-verbose "Fetching existing headers"
        $existingHeaders = Invoke-SqliteQuery `
            -Query "select Path from XisfHeaders" `
            -SQLiteConnection $conn |
            Group-Object Path -AsHashTable
        write-verbose "Calculating files to process"
        $toProcess=
            get-childitem $Path -Recurse -filter:$Filter -file |
                where-object {
                    $x=$_
                    if(-not $existingHeaders){
                        return $true
                    }
                    if(-not (test-path $x.FullName)){
                        return $false
                    }
                    $relativePath = Resolve-Path -Path $x.FullName -RelativeBasePath $Path -Relative
                    return -not ($existingHeaders.ContainsKey($relativePath))
                }
        $i=0
        $toProcess |
            foreach-object {
                $x = $_
                $relativePath = Resolve-Path -Path $x.FullName -RelativeBasePath $Path -Relative
                $results = $null
                $results = Get-XisfHeader -Path $x -Raw
                if(-not $results){
                    write-warning "Failed to retrieve headers from $($x.FullName)"
                    return
                }
                try{
                    Invoke-SqliteQuery `
                    -Query "insert into XisfHeaders(Path,HeaderXml)values(@Path,@HeaderXml)" `
                    -SqlParameters @{
                        Path=$relativePath
                        HeaderXml=$results.OuterXml
                    } `
                    -SQLiteConnection $conn
                }
                catch{
                    write-warning "Skipping file $($x.FullName). $($_.Exception.ToSTring())"
                }
                $i+=1
                if($i % 100 -eq 0){
                    write-progress -Activity "Scanning Xisf Headers" -PercentComplete (100.0*$i/$toProcess.Count) -Status "Calculated $i of $($toProcess.Count) files."
                }
            }
        Write-Progress -Activity "Scanning Xisf Headers" -Completed
    }
    finally{
        $conn.Dispose()
    }
}
Update-XisfHeaderLibrary -DBFile "W:\Astrophotography\XisfHeaders.db" -Path "W:\Astrophotography" -Filter "*.xisf" -Verbose
Update-XisfHeaderLibrary -DBFile "E:\Calibrated\XisfHeaders.db" -Path "E:\Calibrated" -Filter "*.xisf"
Update-XisfHeaderLibrary -DBFile "E:\Astrophotography\XisfHeaders.db" -Path "E:\Astrophotography" -Filter "*.xisf"
Update-XisfHeaderLibrary -DBFile "F:\Astrophotography\XisfHeaders.db" -Path "F:\Astrophotography" -Filter "*.xisf"
Update-XisfHeaderLibrary -DBFile "G:\Astrophotography\XisfHeaders.db" -Path "G:\Astrophotography" -Filter "*.xisf"
Update-XisfHeaderLibrary -DBFile "M:\Calibrated\XisfHeaders.db" -Path "M:\Calibrated" -Filter "*.xisf"
Update-XisfHeaderLibrary -DBFile "N:\Calibrated\XisfHeaders.db" -Path "N:\Calibrated" -Filter "*.xisf"
Update-XisfHeaderLibrary -DBFile "P:\Astrophotography\XisfHeaders.db" -Path "P:\Astrophotography" -Filter "*.xisf"

#Update-HashLibrary -DBFile "W:\Astrophotography\Hashes.db" -Path "W:\Astrophotography" -Filter "*.xisf" -Verbose
#Update-HashLibrary -DBFile "E:\Calibrated\Hashes.db" -Path "E:\Calibrated" -Filter "*.xisf" -RecreateDB
#Update-HashLibrary -DBFile "E:\Astrophotography\Hashes.db" -Path "E:\Astrophotography" -Filter "*.xisf"
#Update-HashLibrary -DBFile "F:\Astrophotography\Hashes.db" -Path "F:\Astrophotography" -Filter "*.xisf"
#Update-HashLibrary -DBFile "G:\Astrophotography\Hashes.db" -Path "G:\Astrophotography" -Filter "*.xisf"
#Update-HashLibrary -DBFile "M:\Calibrated\Hashes.db" -Path "M:\Calibrated" -Filter "*.xisf"
#Update-HashLibrary -DBFile "N:\Calibrated\Hashes.db" -Path "N:\Calibrated" -Filter "*.xisf"
#Update-HashLibrary -DBFile "P:\Astrophotography\Hashes.db" -Path "P:\Astrophotography" -Filter "*.xisf"


$Libraries = @(
    "W:\Astrophotography"
    "E:\Calibrated"
    "M:\Calibrated"
    "N:\Calibrated"
    "E:\Astrophotography"
    "G:\Astrophotography"
    "F:\Astrophotography"
    "P:\Astrophotography"
) | foreach-object {
    new-object psobject -property @{
        DBFile = "$_\Hashes.db"
        Path = $_
    }
}

$hashes = $Libraries | foreach-object {
    $Library = $_.Path
    $DBFile = $_.DBFile
    $conn = New-SQLiteConnection -DataSource $DBFile
    try
    {
        Invoke-SqliteQuery `
            -Query "select * from FileHashes" `
            -SQLiteConnection $conn |
            add-member -NotePropertyName Library -NotePropertyValue $Library -PassThru
    }
    finally{
        $conn.Dispose()
    }
} |
group-object Library 

$hashes| 
    select-object name,Count,{($_.Group|sort-object Path | select-object -last 1).Path}

$hashes.count
$sourceFiles = ($hashes|where-object Name -eq "E:\Astrophotography").Group
$targetFiles = ($hashes|where-object Name -eq "W:\Astrophotography").Group | group-object Path -AsHashTable
$sourceFiles.Count
$targetFiles.Keys.Count
$sourceFiles | 
    sort-object |
    foreach-object {
        $key = $_.Path
        $sourceHash=$_.FileHash
        if(-not $targetFiles.ContainsKey($key)){
            return new-object psobject -property @{
                Path = $key
                SourceHash = $sourceHash
                Result = "Missing from destination."
            }
        }
        $targetHash = $targetFiles[$key].FileHash
        if($sourceHash -ne $targetHash){
            new-object psobject -property @{
                Path = $key
                Result = "Hash Mismatch."
                SourceHash = $sourceHash
                TargetHash = $targetHash
            }
        }
    } | 
    select-object -first 100 Path,Result,SourceHash,TargetHash 
