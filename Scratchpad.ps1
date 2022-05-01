import-module $PSScriptRoot\PsXisfReader -Force

Function Get-XisfImageDataFloat32
{
    [CmdletBinding()]
    param(
        [System.IO.Stream]$stream,
        [int]$position,
        [int]$size
    )
    begin{
        $originalPosition=$stream.Position
        $stream.Position=$position
        $endPosition=$position+$size
        $reader=new-object System.IO.BinaryReader $stream
    }
    process{
        while($stream.Position -lt $endPosition){
            $reader.ReadSingle()
        }
    }
    end
    {
        $stream.Position=$originalPosition
        $reader.Dispose()
    }
}

$testFile = "E:\Astrophotography\DarkLibrary\ZWO ASI183MM Pro\20291220.MasterDark.Gain.53.Offset.10.-15C.74x600s.xisf"
$stream = [System.IO.File]::OpenRead($testFile)
$reader=new-object System.IO.BinaryReader $stream
try
{
    $sig = Read-XisfSignature $reader
    $header=Read-XisfHeader $reader
    $imageLocation=$header.xisf.Image.location
    $geometry = $header.xisf.Image.geometry
    $sampleFormat=$header.xisf.Image.sampleFormat

    if($imageLocation.ToLower().StartsWith("attachment:")){
        $parts=$imageLocation.ToLower().Split(":")
        $position=$parts[1]
        $size=$parts[2]
        #Get-XisfImageDataFloat32 -stream $stream -position $position -size $size
    }

}
catch [System.Xml.XmlException]{
    Write-Warning ("An error occured reading the file "+($Path.FullName))
    Write-Verbose $_.Exception.ToString()
}
catch {
    Write-Warning "An unexpected error occured processing file $($Path.FullName)"
    Write-Warning $_.Exception.ToString()
    throw
}
finally
{
    $reader.Dispose()
    $stream.Dispose()
}


$DarkLibraryFiles|foreach-object {
    $dark=$_

    #$dest = "E:\Astrophotography\DarkLibrary\ZWO ASI183MM Pro\Analysis\G.$($dark.Gain).O.$($dark.Offset).Temp.$($dark.SetTemp).Duration.$($dark.Exposure)"
    $dest = "E:\Astrophotography\DarkLibrary\ZWO ASI183MM Pro\Analysis\G.$($dark.Gain).O.$($dark.Offset).Duration.$($dark.Exposure)"
    [System.IO.Directory]::CreateDirectory($dest)>>$null
    Copy-Item ($dark.Path) -Destination $dest
}

install-module PsXisfReader

import-module PsXisfReader

Get-XisfLightFrames -Path "E:\Astrophotography\1000mm\Eye of Smaug in Cygnus" -Recurse -UseCache -SkipOnError |
    Where-Object {-not $_.HasTokensInPath(@("reject","process","testing","clouds","draft","cloudy","_ez_LS_","drizzle"))} |
    group-object Filter|foreach-object {
        $_.Group |
            Measure-ExposureTime -TotalHours -TotalMinutes |
            add-member -NotePropertyName Filter -NotePropertyValue ($_.Name) -PassThru
    } | 
    format-table Filter,Count,TotalMinutes,TotalHours

    



install-module PsXisfReader
import-module PsXisfReader

Get-XisfLightFrames -Path "S:\PixInsight\Aligned" -SkipOnError |
    sort-Object SSWeight -Descending | 
    Format-Table Path,SSWeight,*

Start-PixInsight -PixInsightSlot 200


Get-ChildItem "E:\Astrophotography\1000mm\Abell 39" -Recurse -Filter *.xisf |
    foreach-object {
        $path = $_
        $header=Get-XisfHeader -Path $path
    }

#if(-not $cred){$cred=Get-Credential -Message "supply username and password to database connect"}
$conn = new-object system.data.sqlclient.sqlconnection "data source=localhost;initial catalog=ImagingLogDev;integrated security=true;"
$conn.Open()
try {
    $cmd = $conn.CreateCommand()
    <#
    $cmd.CommandText = "create schema Inbound;"
    $cmd.ExecuteNonQuery() >> $null
    $cmd.CommandText = "
    create table Inbound.FilesToProcess
    (
        []
        [RelativePath] nvarchar(512) not null,
        [FileName] nvarchar(2048),
        [XisfHeader] xml not null,
        [DateCreated] datetime2(2) not null
    )
    with (data_compression=page)
    "
    $cmd.ExecuteNonQuery() >> $null
    #>

    $cmd.CommandText = "insert into Calibrated.CalibratedFiles(
            [RelativePath],[FileName],[XisfHeader],[DateCreated],[SizeInBytes]
        )
        values
        (
            @RelativePath,@FileName,@XisfHeader,@DateCreated,@SizeInBytes
        )
        "
    $cmd.Parameters.Clear()
    $RelativePath=$cmd.Parameters.Add("@RelativePath",[System.Data.SqlDbType]::NVarChar,2048)
    $SizeInBytes=$cmd.Parameters.Add("@SizeInBytes",[System.Data.SqlDbType]::BigInt)
    $FileName=$cmd.Parameters.Add("@FileName",[System.Data.SqlDbType]::NVarChar,512)
    $XisfHeader=$cmd.Parameters.Add("@XisfHeader",[System.Data.SqlDbType]::Xml,-1)
    $DateCreated=$cmd.Parameters.Add("@DateCreated",[System.Data.SqlDbType]::DateTime2,12)
    $cmd.Prepare()
    Get-ChildItem "F:\PixInsightLT\Calibrated\" -Recurse -Filter *.xisf |
        where-object LastWriteTime -gt "2021-08-14 11:02:28.15" |
        foreach-object {
            $path = $_
            $RelativePath.Value = $path.FullName.Replace("F:\PixInsightLT\Calibrated\","")
            $FileName.Value = $path.Name
            $header = Get-XisfHeader -Path $path
            $reader = new-object System.Xml.XmlNodeReader $header
            try {
                $XisfHeader.Value = new-object System.Data.SqlTypes.SqlXml -ArgumentList @($reader)
                $DateCreated.Value = $path.CreationTime
                $SizeInBytes.Value = $path.Length
                write-host "Inserting row... $($path.Name)"
                $cmd.ExecuteNonQuery() >> $null
            }
            finally {
                $reader.Dispose()    
            }            
        }
}
finally {
    $conn.Dispose()
}





$rejects = Get-ChildItem E:\Astrophotography -Recurse -Filter *.xisf -File |
    where-object {$_.Directory.FullName.ToLower().Contains(@("rejection"))}
$rejectionStats = $rejects| Measure-Object Length -Sum -Average
write-host "$($rejectionStats.Count.ToString('#,0')) files with a total size of $(($rejectionStats.Sum/1024/1024/1024).ToString('#,0.0')) GB"


$all = Get-XisfLightFrames -Path "E:\Astrophotography" -Recurse -UseCache -UseErrorCache -SkipOnError -ShowProgress
$byEquipment = $all | group-object Instrument

$byEquipment | select Name
