if (-not (get-module psxisfreader)){import-module psxisfreader}
$ErrorActionPreference="STOP"

$targets = @(
    #"E:\Astrophotography\90mm\M81 M82 Region"
    #"E:\Astrophotography\950mm\vdb131 vdb132 Take 2"
    "E:\Astrophotography\950mm\vdb15"
    # "E:\Astrophotography\950mm\Eye of Smaug Take 2"
    # "E:\Astrophotography\950mm\Eye of Smaug Take 2P2"
    # "E:\Astrophotography\950mm\Properller and Smaug Panel 1"
    # "E:\Astrophotography\950mm\Properller and Smaug Panel 2"
    #"P:\Astrophotography\40mm\Cygnus near Sh2-115"
)

Function Format-ExposureTime
{
    param(
        [Parameter(Mandatory,ValueFromPipeline)][TimeSpan]$TimeSpan)

    process{
        if($TimeSpan -gt [timespan]::Zero){
            Write-Output "$([Math]::Floor($TimeSpan.TotalHours))h $($TimeSpan.ToString('mm'))'"
        }
        else{
            Write-Output ""
        }        
    }
}

$filterMap = @{
    "Sii3"="Antlia 3.5nm Narrowband Sulfur II 2`""
    "L"="Antlia Luminance 2`""
    "L3"="Astronomik Luminance 2`""
    "B"="Chroma Blue 2`""
    "G"="Chroma Green 2`""
    "R"="Chroma Red 2`""
    "Ha"="Chroma H-alpha 5nm Bandpass 2`""
    "Ha6nmMaxFr"="Astronomik Ha 6nm MaxFR 2`""
    "Oiii"="Chroma OIII 3nm Bandpass 2`""
}


$filterMap = @{
    "Sii3"="Antlia 3.5nm Narrowband Sulfur II 2`""
    "L"="Antlia Luminance 2`""
    "L3"="Astronomik Luminance 2`""
    "B"="Chroma Blue 2`""
    "G"="Chroma Green 2`""
    "R"="Chroma Red 2`""
    "Ha"="Chroma H-alpha 5nm Bandpass 2`""
    "Ha6nmMaxFr"="Astronomik Ha 6nm MaxFR 2`""
    "Oiii"="Chroma OIII 3nm Bandpass 2`""
}
$filterMap = @{
    #https://app.astrobin.com/equipment/explorer/filter/5282/chroma-lum-2
    "L"="5282"
    #https://app.astrobin.com/equipment/explorer/filter/3565/chroma-red-2
    "R"="3565"
    #https://app.astrobin.com/equipment/explorer/filter/3995/chroma-green-2
    "G"="3995"
    #https://app.astrobin.com/equipment/explorer/filter/3994/chroma-blue-2
    "B"="3994"
    #https://app.astrobin.com/equipment/explorer/filter/5240/chroma-h-alpha-5nm-bandpass-2
    "Ha5nm"="5240"
    #https://app.astrobin.com/equipment/explorer/filter/4160/chroma-oiii-3nm-bandpass-2
    "Oiii3nm"="4160"
    #https://app.astrobin.com/equipment/explorer/filter/4381/antlia-35nm-narrowband-sulfur-ii-2
    "Sii35nm"="4381"
}

$targetFilesPath = "E:\Astrophotography\950mm\vdb 15"
$outputCsvFile = "E:\Astrophotography\950mm\vdb 15\AstrobinExport.csv"
$results = Get-XisfLightFrames -Path $targetFilesPath `
    -SkipOnError `
    -Recurse `
    -UseCache:$false `
    -PathTokensToIgnore @("reject","process","testing","clouds","draft","cloudy","_ez_LS_","drizzle","skyflats","quickedit","calibrated") |
    Where-Object {$_.LocalDate } |
    Where-Object {$_.Filter -ne "Ha3nm"} |
    Where-Object {$_.Filter -ne "Ha5nm"} |
    Where-Object {-not $_.IsIntegratedFile()} |
    where-object {-not [string]::IsNullOrWhiteSpace($_.Object)} |
    Group-Object ObsDateMinus12hr,Filter,Exposure,ISO,Gain,CCDTemp,FocalRatio |
    foreach-object {
        $group=$_.Group
        $first=$group[0]
        $filter=$filterMap."$($first.Filter)"
        if(-not $filter){
            $filter = $first.Filter
        }
        new-object psobject -Property ([Ordered]@{
            date = $first.ObsDateMinus12hr.ToString("yyyy-MM-dd")
            filter=$filter
            number=$group.Count
            duration=$first.Exposure
            #iso = $first.iso
            binning = 1
            gain=$first.Gain
            sensorCooling=([Math]::Round($first.CCDTemp/7)*7)
            fNumber=$first.FocalRatio
            darks=30
            flats=16
            flatDarks=16
            bias=200
            bortle=4
        })
    } |
    sort-object Date,Filter,Number,Duration

$results | format-table date,filter,number,duration,iso,binning,gain,sensorcooling,fnumber,darks,flats,flatdarks,bias,bortle
$results | Export-Csv $outputCsvFile 

$astrobinSessions = $targets | 
    foreach-object {
        Get-XisfLightFrames -Path $_ `
            -SkipOnError `
            -Recurse `
            -UseCache `
            -PathTokensToIgnore @("reject","process","testing","clouds","draft","cloudy","_ez_LS_","drizzle","skyflats","quickedit","calibrated") 
    } |
    Where-Object {$_.LocalDate } |
    Where-Object {-not $_.IsIntegratedFile()} |
    where-object {-not [string]::IsNullOrWhiteSpace($_.Object)} |
    Group-Object ObsDateMinus12hr,Filter,Exposure,ISO,Gain,CCDTemp,FocalRatio
$data =$astrobinSessions|
    foreach-object {
        $group=$_.Group
        $first=$group[0]
        $filter=$filterMap."$($first.Filter)"
        new-object psobject -Property @{
            Date = $first.ObsDateMinus12hr.ToString("yyyy-MM-dd")
            Filter=$filter
            Duration=$first.Exposure
            Number=$group.Count
            ISO = $first.iso
            Gain=$first.Gain
            SensorCooling=$first.CCDTemp
            FocalRatio=$first.FocalRatio
        }        
    } 
$data |
    sort-object Date,Filter,Number,Duration |
    format-table Date,Filter,Number,Duration,ISO,Gain,SensorCooling,FocalRatio

"Astrobin Acquisition Sessions: $($astrobinSessions.Length)"

"Dates ($(($data | 
group-object Date).Length) sessions) : $(
    $data | 
    group-object Date | 
    foreach-object {
        " $($_.Name)"
    }
)"


"Frames: $(
    $data | 
    group-object Filter,Exposure | 
    foreach-object {
        $first=$_.Group[0]; 
        $count=([int]($_.Group | measure-object -Sum Count).Sum)
        new-object psobject -Property @{
            Filter=$first.Filter;
            Exposure=$first.Exposure; 
            Count=$count
            Display=Format-ExposureTime -TimeSpan ([TimeSpan]::FromSeconds($first.Exposure*$count))
        }
    } | 
    foreach-object {
        "
    $($_.Filter): $($_.Count)x$($_.Exposure)""" ($($_.Display))"
    }
)"

"Integration: $(
Format-ExposureTime -TimeSpan `
    ([TimeSpan]::FromSeconds(
    ($data | 
    group-object Exposure |
    foreach-object {
        $first=$_.Group[0]
        $exposure=$first.Exposure
        $count=($_.Group|measure-object -Property Count -sum).Sum
        [int]$exposure*$count
    } |
    measure-object -sum).Sum)))"
exit
