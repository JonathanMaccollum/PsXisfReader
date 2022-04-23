if (-not (get-module psxisfreader)){import-module psxisfreader}
$ErrorActionPreference="STOP"
Import-module "C:\Program Files\N.I.N.A. - Nighttime Imaging 'N' Astronomy\NINA.Astrometry.dll"

Function Format-ExposureTime
{
    param(
        [Parameter(Mandatory,ValueFromPipeline)][TimeSpan]$TimeSpan)

    process{
        if($TimeSpan -gt [timespan]::Zero){
            Write-Output "$([Math]::Floor($TimeSpan.TotalHours)):$($TimeSpan.ToString('mm'))"
        }
        else{
            Write-Output ""
        }        
    }
}

Function Format-MDExposureTableDatesOverFilter
{
    param(
        [Array]$Data,
        [Switch]$IncludeTotals
    )

    $result="| Date | "
        $byFilter = $Data | group-object Filter
        $byFilter | foreach-object {
            $result += $_.Name + " | "
        }
        $result = $result.Trim()
        $result += "
| ---- | "
    $byFilter | foreach-object {
        $result += "----" + " | "
    }
    $result = $result.Trim()

    $result += "
"
    $Data | 
        group-object ObsDateMinus12hr |
        foreach-object {
            $byDate = $_.Group
            $obsDate = $byDate[0].ObsDateMinus12hr
            $result += "|$($obsDate.ToString('yyyy-MM-dd'))|"
            
            $byFilter |
                foreach-object {
                    $filter = $_.Group[0].Filter

                    $filterDataOnDate = $byDate | where-object Filter -eq $filter
                    
                    $integrationTimeByFilter = $filterDataOnDate | Measure-ExposureTime
                    $result += "$($integrationTimeByFilter.Total | Format-ExposureTime)|"
                }
            $result += "
"
        }

    if($IncludeTotals.IsPresent){
            
        $result += "| **Totals** | "
        $byFilter |
        foreach-object {
            $filter = $_
            $integrationTimeByFilter = $filter.Group | Measure-ExposureTime
            $result = $result + "**$($integrationTimeByFilter.Total | Format-ExposureTime)**|"
        }
        $result = $result.Trim()
    }

    Write-Output $result
}

Function Format-MDExposureTableObjectOverFilter
{
    param(
        [Array]$Data,
        [Switch]$IncludeFilters,
        [ScriptBlock]$FormatObjectName
    )

    $result="| Object | "
        $byFilter = $Data | group-object Filter
        if($IncludeFilters){
            $byFilter | foreach-object {
                $result += $_.Name + " | "
            }
        }
        $result += "**Total** |"
        $result = $result.Trim()
        $result += "
| ---- | "
    if($IncludeFilters){
        $byFilter | foreach-object {
            $result += "----" + " | "
        }
    }
    $result += " ---- | "
    $result += "
"
    $Data | 
        group-object Object |
        sort-object { ($_.Group | Measure-Object LocalDate -Maximum).Maximum } -Descending |
        foreach-object {
            $byObject = $_.Group
            if($FormatObjectName){
                $object = $FormatObjectName.Invoke($byObject[0].Object)
            }
            else{
                $object = $byObject[0].Object
            }     
            $result += "|$($object)|"
            if($IncludeFilters){
                $byFilter |
                    foreach-object {
                        $filter = $_.Group[0].Filter

                        $filterDataByObject = $byObject | where-object Filter -eq $filter
                        
                        $integrationTimeByFilter = $filterDataByObject | Measure-ExposureTime
                        $result += "$($integrationTimeByFilter.Total | Format-ExposureTime)|"
                    }
            }
            $result += "**$(($byObject | Measure-ExposureTime).Total | Format-ExposureTime)**|"
            $result += "
"
        }

    Write-Output $result
}


Function Format-MDExposureTableDatesAndObjectOverFilter
{
    param(
        [Array]$Data,
        [Switch]$IncludeTotals,
        [ScriptBlock]$FormatObjectName
    )

    $result="| Date | Object |"
        $byFilter = $Data | group-object Filter
        $byFilter | foreach-object {
            $result += $_.Name + " | "
        }
        $result = $result.Trim()
        $result += "
| ---- | ------ | "
    $byFilter | foreach-object {
        $result += "----" + " | "
    }
    $result = $result.Trim()

    $result += "
"
    $Data | 
        group-object ObsDateMinus12hr,Object |
        sort-object {$_.Group[0].ObsDateMinus12hr} -Descending |
        foreach-object {
            $byDate = $_.Group
            $obsDate = $byDate[0].ObsDateMinus12hr
            if($FormatObjectName){
                $object = $FormatObjectName.Invoke($byDate[0].Object)
            }
            else{
                $object = $byDate[0].Object
            }            
            $result += "|$($obsDate.ToString('yyyy-MM-dd'))|$object|"
            
            $byFilter |
                foreach-object {
                    $filter = $_.Group[0].Filter

                    $filterDataOnDate = $byDate | where-object Filter -eq $filter
                    
                    $integrationTimeByFilter = $filterDataOnDate | Measure-ExposureTime
                    $result += "$($integrationTimeByFilter.Total | Format-ExposureTime)|"
                }
            $result += "
"
        }

    Write-Output $result
}
@(
    #40,
    #50,
    90
    135
    1000
    )|ForEach-Object{
        $focalLength=$_

        $data = 
            (Get-XisfLightFrames -Path "E:\Astrophotography\$($focalLength)mm" -SkipOnError -Recurse -UseCache -PathTokensToIgnore @("reject","process","testing","clouds","draft","cloudy","_ez_LS_","drizzle","skyflats","quickedit")) |
            Where-Object {$_.LocalDate } |
            Where-Object {-not $_.IsIntegratedFile()} 

        if(-not $data){

        }

        $targets = Format-MDExposureTableObjectOverFilter -Data $data `
            -FormatObjectName {
                param($objectName)
                "[$objectName](/ImagingLog/$($focalLength)mm/$($objectName.Replace(' ','%20')).html)"
            }
        $targets|Out-File -FilePath "S:\JonsAstro\projects.darkflats.com\source\ImagingLog\Targets.$($focalLength)mm.md" -Force


        $data |
            group-object Object |
            foreach-object {
                new-object psobject -Property @{
                    Group = ($_.Group)
                    Object = ($_.Group[0].Object)
                    FocalLength = $focalLength
                    IntegrationTime = ($_.Group | Measure-ExposureTime)
                    Stats = ($_.Group | Measure-Object LocalDate -Minimum -Maximum)
                }
            } |
            foreach-object {
                $group = $_.Group
                $object = $_.Object
                $integrationTime = $_.IntegrationTime
                $stats=$_.Stats

                $fileName = "$($object).md"
                $outputDir = "S:\JonsAstro\projects.darkflats.com\source\ImagingLog\$($focalLength)mm"
                [System.IO.Directory]::CreateDirectory($outputDir)>>$null
                $file = join-path $outputDir $fileName
                $content = "---
title: $object
date: $($stats.Maximum.ToString('yyyy-MM-dd HH:mm:ss'))
---
**$Object**

* Total time: $($integrationTime.Total | Format-ExposureTime)

$(Format-MDExposureTableDatesOverFilter -Data $group -IncludeTotals)
"
                $content | out-file $file -Force
            }

        $fullTable = Format-MDExposureTableDatesAndObjectOverFilter `
            -IncludeTotals `
            -Data $data `
            -FormatObjectName {
                param($objectName)

                "[$objectName](/ImagingLog/$($focalLength)mm/$($objectName.Replace(' ','%20')).html)"
            }

        $fullTable|Out-File -FilePath "S:\JonsAstro\projects.darkflats.com\source\ImagingLog\ImagingLog.$($focalLength)mm.md" -Force
    }

Push-Location "S:\JonsAstro\projects.darkflats.com"
try{
    git add -A
    git commit -m "Updated Imaging Log $([DateTime]::Now.ToString('yyyyMMdd'))"
    hexo generate
    & S:\JonsAstro\publishProjectsDarkFlats.ps1
}
finally{
    Pop-Location 
}