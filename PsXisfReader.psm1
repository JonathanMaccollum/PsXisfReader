Add-Type -AssemblyName 'System.Xml.Linq'
Function Read-XisfSignature([System.IO.BinaryReader]$reader)
{
    $buffer=[char[]]::new(8)
    $reader.Read($buffer,0,8)>$null
    [System.String]::new($buffer)
}
Function Read-XisfHeader([System.IO.BinaryReader]$reader)
{
    $length = $reader.ReadUInt32()
    $buffer=[char[]]::new(4)
    $reader.Read($buffer,0,4)>$null
    $header=[byte[]]::new($length)
    $reader.Read($header,0,$length)>$null

    $memoryStream=[System.IO.MemoryStream]::new($header)
    try
    {
        $xmlReader=[System.Xml.XmlReader]::Create($memoryStream)
        try
        {
            [xml][System.Xml.Linq.XDocument]::Load($xmlReader)
        }
        finally
        {
            $xmlReader.Dispose()
        }
    }
    finally
    {
        $memoryStream.Dispose()
    }
}
Function Get-XisfFitsStats
{
    [CmdletBinding()]
    param
    (
        [Parameter(ValueFromPipeline=$true,Mandatory=$true)][System.IO.FileInfo]$Path,
        [Parameter(ValueFromPipeline=$false)][Hashtable]$Cache
    )
    if($Cache -and $Cache.ContainsKey($Path.FullName))
    {
        Write-Output ($Cache[$Path.FullName])
    }
    else
    {
        $stream = [System.IO.File]::OpenRead($Path.FullName)
        $reader=new-object System.IO.BinaryReader $stream
        try
        {

            Read-XisfSignature $reader > $null
            $header=Read-XisfHeader $reader
            $fits = $header.xisf.Image.FITSKeyword
            $filter = $fits.Where({$_.Name -eq 'FILTER'}).value
            if($filter)
            {
                $filter = $filter.Replace('5nm','').Replace('3nm','')
            }
            $result = new-object psobject -Property @{
                Exposure=$fits.Where{$_.Name -eq 'EXPOSURE'}.value
                Filter=$filter
                Object=$fits.Where{$_.Name -eq 'OBJECT'}.value
                Gain=$fits.Where{$_.Name -eq 'GAIN'}.value
                Offset=$fits.Where{$_.Name -eq 'OFFSET'}.value
                ImageType=$fits.Where{$_.Name -eq 'IMAGETYP'}.value
                CCDTemp=$fits.Where{$_.Name -eq 'CCD-TEMP'}.value
                FocalLength=$fits.Where{$_.Name -eq 'FOCALLEN'}.value
                FocalRatio=$fits.Where{$_.Name -eq 'FOCRATIO'}.value
                ObsDate=$fits.Where{$_.Name -eq 'DATE-OBS'}.value
                LocalDate=$fits.Where{$_.Name -eq 'DATE-LOC'}.value
                SSWeight=$fits.Where{$_.Name -eq 'SSWEIGHT'}.value
                Path=$Path
            }
            if($Cache)
            {
                $Cache.Add($Path.FullName,$result)
            }
            Write-Output $result
        }
        catch [System.Xml.XmlException]{
            Write-Warning ("An error occured reading the file "+($Path.FullName))
            Write-Verbose $_.Exception.ToString()
        }
        finally
        {
            $reader.Dispose()
            $stream.Dispose()
        }
    }
}
