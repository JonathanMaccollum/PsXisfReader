Function ConvertTo-FileWithTimestamp{
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline)]
        [IO.FileInfo]
        $File
    )
    process{
        $name = $File.Name
        $date = $name.Split("_")[3]
        $time = $name.Split("_")[4].Replace("-",":")
        $timestamp = [DateTime]::Parse($date + " "+$time)
        new-object psobject -Property @{
            FileInfo = $File
            Timestamp = $timestamp
        }
    }
}
Function Add-Index
{
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline)]
        [Object]
        $Item,
        [switch]
        $Force
    )
    begin{
        $i = 0
    }
    process{
        $result = $Item.PSObject.Copy()
        Add-Member -NotePropertyName "Index" -NotePropertyValue $i -InputObject $result -PassThru -Force:$Force.IsPresent
        $i++
    }
}


$filesWithTimestamp = 
    get-childitem S:\PixInsight\Timelapse -File -Filter All-Sky*.tiff |
    ConvertTo-FileWithTimestamp |
    Sort-Object Timestamp |
    Where-object Timestamp -lt "2021-12-01 04:30:00 AM" |
    Add-Index

$r = $filesWithTimestamp |
    where-object {$_.Index % 3 -eq 0} |
    Add-Index -Force |
    where-object {$_.Index % 60 -eq 0} |
    where-object {$_.Timestamp.Hour }
$g = $filesWithTimestamp |
    where-object {$_.Index % 3 -eq 1} |
    Add-Index -Force |
    where-object {$_.Index % 60 -eq 2}
$b = $filesWithTimestamp |
    where-object {$_.Index % 3 -eq 2} |
    Add-Index -Force |
    where-object {$_.Index % 60 -eq 4}


Invoke-PiLightIntegration `
    -OutputDefinitionOnly `
    -PixInsightSlot 200 `
    -Images ($b.FileInfo) `
    -KeepOpen -OutputFile "E:\Astrophotography\2.5mm\All-Sky\20211130\RGBForEveryThree_MaxOfEachTen.B.xisf"