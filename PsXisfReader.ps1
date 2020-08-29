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

class XisfFileStats {
    [decimal]$Exposure
    [string]$Filter
    [string]$Instrument
    [string]$Object
    [decimal]$Gain
    [decimal]$Offset
    [string]$ImageType
    [decimal]$CCDTemp
    [decimal]$SetTemp
    [decimal]$FocalLength
    [decimal]$FocalRatio
    [nullable[DateTime]]$ObsDate
    [nullable[DateTime]]$ObsDateMinus12hr
    [nullable[DateTime]]$LocalDate
    [decimal]$SSWeight
    [decimal]$Pedestal
    [string[]]$History
    [System.IO.FileInfo]$Path

    [bool]HasTokensInPath([string[]]$tokens){
            $hasToken=$false
            foreach( $x in $tokens) {
                if($this.Path.FullName.ToLower().Contains($x.ToLower())){
                    $hasToken=$true
                    break;
                }
            }
            return $hasToken;
        }
    [bool]IsIntegratedFile() {
        return [bool]($this.History)
    }
}
Function Get-XisfFitsStats
{
    [CmdletBinding()]
    [OutputType([XisfFileStats])]
    param
    (
        [Parameter(ValueFromPipeline=$true,Mandatory=$true)][System.IO.FileInfo]$Path,
        [Parameter(ValueFromPipeline=$false)][Hashtable]$Cache
    )
    process{            
        if($Cache -and $Cache.ContainsKey($Path.FullName))
        {
            $result = $Cache[$Path.FullName]
            
            Write-Output ([XisfFileStats]@{
                Exposure=$result.Exposure
                Filter=$result.Filter
                Instrument=$result.Instrument
                Object=$result.Object
                Gain=$result.Gain
                Offset=$result.Offset
                ImageType=$result.ImageType
                CCDTemp=$result.CCDTemp
                SetTemp=$result.SetTemp
                FocalLength=$result.FocalLength
                FocalRatio=$result.FocalRatio
                ObsDate=$result.ObsDate
                ObsDateMinus12hr=$result.ObsDateMinus12hr
                LocalDate=$result.LocalDate
                SSWeight=$result.SSWeight
                Pedestal=$result.Pedestal
                History=$result.History
                Path=$Path
            })
        }
        else
        {
            $result=$null
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
                    $filter = $filter.Replace('5nm','').Replace('3nm','').Replace("'","")
                }
                $obsDate = $fits.Where{$_.Name -eq 'DATE-OBS'}.value | Select-Object -First 1
                $obsDateMinus12hr= $null
                if($obsDate){
                    $obsDateMinus12hr=$obsDate|%{([DateTime]($_.Trim("'"))).AddHours(-12).Date}
                }
                $results=@{
                    Exposure=$fits.Where{$_.Name -eq 'EXPOSURE'}.value
                    Filter=$filter
                    Instrument=$fits.Where{$_.Name -eq 'INSTRUME'}.value|%{if($_){$_.Trim("'")}}
                    Object=$fits.Where{$_.Name -eq 'OBJECT'}.value|%{if($_){$_.Trim("'")}}
                    Gain=$fits.Where{$_.Name -eq 'GAIN'}.value
                    Offset=$fits.Where{$_.Name -eq 'OFFSET'}.value
                    ImageType=$fits.Where{$_.Name -eq 'IMAGETYP'}.value|%{if($_){$_.Trim("'")}}
                    CCDTemp=$fits.Where{$_.Name -eq 'CCD-TEMP'}.value
                    SetTemp=$fits.Where{$_.Name -eq 'SET-TEMP'}.value
                    FocalLength=$fits.Where{$_.Name -eq 'FOCALLEN'}.value|%{if($_){$_.Trim("'").TrimEnd("mm")}}
                    FocalRatio=$fits.Where{$_.Name -eq 'FOCRATIO'}.value
                    ObsDate=$obsDate|%{if($_){$_.Trim("'")}}
                    ObsDateMinus12hr=$obsDateMinus12hr
                    LocalDate=$fits.Where{$_.Name -eq 'DATE-LOC'}.value|%{if($_){$_.Trim("'")}}
                    SSWeight=$fits.Where{$_.Name -eq 'SSWEIGHT'}.value|%{if($_){$_.Trim("'")}}
                    Pedestal=$fits.Where{$_.Name -eq 'PEDESTAL'}.value|%{if($_){$_.Trim("'")}}
                    History=$fits.Where{$_.Name -eq 'HISTORY'}.comment
                    Path=$Path
                }
                $result = [XisfFileStats] $results
                if($Cache)
                {
                    $Cache.Add($Path.FullName,$result)
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

            if($result){
                Write-Output $result
            }
        }
    }
}



Function Measure-ExposureTime
{
    param(
        [Parameter(ValueFromPipeline=$true)][XisfFileStats[]]$Input
    )
    begin {
        $totalSeconds=0.0;
    }
    process{
        $totalSeconds+=$Input.Exposure
    }
    end {
        [TimeSpan]::FromSeconds($totalSeconds)
    }
}
