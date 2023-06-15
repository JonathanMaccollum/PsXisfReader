if (-not (get-module psxisfreader)){import-module psxisfreader}
Import-Module ResizeImageModule

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
        $byFilter = $Data | group-object Filter | sort-object {
            if($_.Name.StartsWith("L")){
                1
            }
            elseif($_.Name.StartsWith("R")){
                2
            }
            elseif($_.Name.StartsWith("G")){
                3
            }
            elseif($_.Name.StartsWith("B")){
                4
            }
            elseif($_.Name.StartsWith("H")){
                5
            }
            elseif($_.Name.StartsWith("O")){
                6
            }
            elseif($_.Name.StartsWith("S")){
                7
            }
            else{
                8
            }
        }
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
        $byFilter = $Data | group-object Filter | sort-object {
            if($_.Name.StartsWith("L")){
                1
            }
            elseif($_.Name.StartsWith("R")){
                2
            }
            elseif($_.Name.StartsWith("G")){
                3
            }
            elseif($_.Name.StartsWith("B")){
                4
            }
            elseif($_.Name.StartsWith("H")){
                5
            }
            elseif($_.Name.StartsWith("O")){
                6
            }
            elseif($_.Name.StartsWith("S")){
                7
            }
            else{
                8
            }
        }
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
        group-object {$_.Object.Trim()} |
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

Function Out-TargetImageDetailPage
{
    param(
        $fileName,
        $outputDir,
        $date,
        $object,
        $integrationTime,
        $telescope,
        $instrument,
        $nightsImaged,
        $started,
        $latest,
        $thumbnails,
        $subsForTarget
    )

    [System.IO.Directory]::CreateDirectory($outputDir)>>$null
    $file = join-path $outputDir $fileName
    $content = "---
title: $object
date: $($date.ToString('yyyy-MM-dd HH:mm:ss'))
---
**$Object**

* Total time: $($integrationTime.Total | Format-ExposureTime)
* Scope: $telescope
* Camera: $instrument
* Nights Imaged: $nightsImaged
* Started: $($started.ToString('yyyy-MM-dd')) 
* Latest: $($latest.ToString('yyyy-MM-dd'))

"
    $thumbnails | foreach-object {
        if($_.Details){
            $content += "* $($_.Details)

"
        }
        $content += "![$($_.Name)]($($_.Url) `"$($_.Name)`")
"
    }

    $content += "
$(Format-MDExposureTableDatesOverFilter -Data $subsForTarget -IncludeTotals)"

    $content | out-file $file -Force
}
@(
    #40
    #50
    #90
    #135
    #350
    950
    #1000
    )|ForEach-Object{
        $focalLength=$_

        $topLevelFolder = "E:\Astrophotography\$($focalLength)mm"
        $data = 
            (Get-XisfLightFrames -Path $topLevelFolder `
                -SkipOnError `
                -Recurse `
                -UseCache `
                -PathTokensToIgnore @("reject","process","testing","clouds","draft","cloudy","_ez_LS_","drizzle","skyflats","quickedit","calibrated")) |
            Where-Object {$_.LocalDate } |
            Where-Object {-not $_.IsIntegratedFile()} |
            where-object {-not [string]::IsNullOrWhiteSpace($_.Object)}


        $data | foreach-object {
            $targetObjectRootPath = 
                get-item(
                    Join-Path $topLevelFolder ([IO.Path]::GetRelativePath($topLevelFolder, $_.Path).Split("\")[0]))
            if($_.Object.Trim() -ne $targetObjectRootPath.Name){
                $_.Object =$targetObjectRootPath.Name
            }
        }

        $targets = 
        [string]::Concat(
            (Format-MDExposureTableObjectOverFilter -Data $data `
                -FormatObjectName {
                    param($objectName)
                    "[$objectName](/ImagingLog/$($focalLength)mm/$($objectName.Trim().Replace(' ','%20')).html)"
                }),
                "
", (Format-MDExposureTableObjectOverFilter -Data $data -IncludeFilters `
                -FormatObjectName {
                    param($objectName)
                    "[$objectName](/ImagingLog/$($focalLength)mm/$($objectName.Trim().Replace(' ','%20')).html)"
                }))

        $targets|Out-File -FilePath "S:\JonsAstro\projects.darkflats.com\source\ImagingLog\Targets.$($focalLength)mm.md" -Force


        $data |
            group-object {$_.Object.Trim()} |
            foreach-object {
                $x=$_
                new-object psobject -Property @{
                    Group = ($x.Group)
                    Object = ($x.Group[0].Object.Trim())
                    FocalLength = $focalLength
                    IntegrationTime = ($x.Group | Measure-ExposureTime)
                    Stats = ($x.Group | Measure-Object ObsDateMinus12hr -Minimum -Maximum )
                }
            } |
            foreach-object {
                $subsForTarget = $_.Group

                $targetObjectRootPath = Join-Path $topLevelFolder ([IO.Path]::GetRelativePath($topLevelFolder, $subsForTarget[0].Path).Split("\")[0])
                $targetThumbnailsPath = Join-Path $targetObjectRootPath "Thumbnails"

                $object = (get-item $targetObjectRootPath).Name
                $integrationTime = $_.IntegrationTime
                $stats=$_.Stats
                $outputDir = "S:\JonsAstro\projects.darkflats.com\source\ImagingLog\$($focalLength)mm"
                [System.IO.Directory]::CreateDirectory($outputDir)>>$null
                $thumbnailOutputDir = Join-Path $outputDir Thumbnails
                [System.IO.Directory]::CreateDirectory($thumbnailOutputDir)>>$null
                
                $headers = ($subsForTarget[0].Path | Get-XisfHeader).xisf.Image.FITSKeyword
                $telescope = $null
                try{$telescope = $headers |
                        where-object Name -eq "TELESCOP" |
                        foreach-object {$_.Value}}
                catch{write-warning "Unable to determine telescope used for $($object)"}
                if(-not $telescope){
                    $telescope = ""
                }

                $nights = $subsForTarget | group-object ObsDateMinus12hr
                

                # Fetch and Generated Thumbnails
                $thumbnails = @()
                $thumbnails += Get-ChildItem $targetObjectRootPath -File -Recurse |
                    where-object { $_.FullName.Contains("Processing") } |
                    where-object { -not $_.FullName.Contains(".Web.") } |
                    where-object { -not $_.FullName.Contains("CatalogStars.") } |
                    where-object {$_.Extension -in @('.jpeg','.jpg','.png')} |
                    where-object {
                        $xisfEquiv=($_.FullName.Substring(0,$_.FullName.Length-$_.Extension.Length)+".xisf")
                        Test-Path $xisfEquiv
                    } |
                    sort-object LastWriteTime -Descending |
                    foreach-object {
                        $x = $_
                        $targetSpecificOutputPath = Join-Path $thumbnailOutputDir $object
                        [IO.Directory]::CreateDirectory($targetSpecificOutputPath)>>$null
                        $thumbnailOutputPath = Join-Path $targetSpecificOutputPath $x.Name
                        
                        try{
                            if(-not (Test-Path $thumbnailOutputPath)){
                                write-host "Creating thumbnail for $($x.Name)"
                                Resize-Image -InputFile $x -OutputFile $thumbnailOutputPath -ProportionalResize $true -Width 1280 -Height 1280
                            }

                            $webPath = [System.Web.HttpUtility]::UrlPathEncode((Join-Path $object $x.Name))
                            new-object psobject -Property @{
                                Name = $x.Name.Replace("."," ")
                                Details = $null
                                Url = "/ImagingLog/$($focalLength)mm/Thumbnails/$webPath"
                                LastWriteTime = $x.LastWriteTime
                            }
                        }
                        catch{
                            write-warning "Failed to produce thumbnail for $($x.Name)"
                        }
                    }

                # Add Filter-Specific Thumbnails to the page
                $thumbnails += $subsForTarget | 
                    Group-Object Filter | 
                    sort-object {
                        if($_.Name.StartsWith("L")){
                            1
                        }
                        elseif($_.Name.StartsWith("R")){
                            2
                        }
                        elseif($_.Name.StartsWith("G")){
                            3
                        }
                        elseif($_.Name.StartsWith("B")){
                            4
                        }
                        elseif($_.Name.StartsWith("H")){
                            5
                        }
                        elseif($_.Name.StartsWith("O")){
                            6
                        }
                        elseif($_.Name.StartsWith("S")){
                            7
                        }
                        else{
                            8
                        }
                    } | ForEach-Object {
                        $filter=$_.name
                        $byFilter = $_.Group
                        
                        if(-not (test-path $targetThumbnailsPath)){
                            return
                        }
                        $mostRecentThumbnail = Get-ChildItem $targetThumbnailsPath "$object.$filter.*.jpeg" |
                            Sort-Object LastWriteTime -Descending |
                            Select-Object -First 1
                        if(-not $mostRecentThumbnail){
                            return
                        }
                        $thumbnailOutputPath = Join-Path $thumbnailOutputDir $mostRecentThumbnail.Name
                        if(-not (Test-Path $thumbnailOutputPath)){
                            write-host "Creating thumbnail for $($mostRecentThumbnail.Name)"
                            Resize-Image -InputFile $mostRecentThumbnail -OutputFile $thumbnailOutputPath -ProportionalResize $true -Width 1280 -Height 1280
                        }
                        $webPath=[System.Web.HttpUtility]::UrlPathEncode($mostRecentThumbnail.Name)
                        $integrationTimeByFilter = $byFilter | Measure-ExposureTime
                        new-object psobject -Property @{
                            Name = "$object - $filter"
                            Url = "/ImagingLog/$($focalLength)mm/Thumbnails/$webPath"
                            Details = "$filter : $($integrationTimeByFilter.Total | Format-ExposureTime)"
                        }
                    }

                    Out-TargetImageDetailPage `
                        -fileName "$($object).md" `
                        -outputDir "S:\JonsAstro\projects.darkflats.com\source\ImagingLog\$($focalLength)mm" `
                        -date $stats.Maximum `
                        -object $object `
                        -integrationTime $integrationTime `
                        -telescope $telescope `
                        -instrument $subsForTarget[0].Instrument `
                        -nightsImaged $nights.Count `
                        -started $stats.Minimum `
                        -latest $stats.Maximum `
                        -thumbnails $thumbnails `
                        -subsForTarget $subsForTarget
            }

        $fullTable = Format-MDExposureTableDatesAndObjectOverFilter `
            -IncludeTotals `
            -Data $data `
            -FormatObjectName {
                param($objectName)

                "[$objectName](/ImagingLog/$($focalLength)mm/$($objectName.Trim().Replace(' ','%20')).html)"
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