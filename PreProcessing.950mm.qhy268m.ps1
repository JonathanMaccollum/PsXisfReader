if (-not (get-module psxisfreader)){import-module psxisfreader}

$ErrorActionPreference="STOP"
$WarningPreference="Continue"
$DropoffLocation = "D:\Backups\Camera\Dropoff\NINA"
$ArchiveDirectory="W:\Astrophotography"
$CalibratedOutput = "E:\Calibrated\950mm"
<#
Invoke-BiasFrameSorting `
    -DropoffLocation $DropoffLocation `
    -ArchiveDirectory $ArchiveDirectory `
    -PixInsightSlot 201 -KeepOpen
Invoke-DarkFrameSorting `
    -DropoffLocation $DropoffLocation `
    -ArchiveDirectory $ArchiveDirectory `
    -PixInsightSlot 201 -KeepOpen
    #>
<#
    #>

$PushToLightBucket=$true
[Array]$BiasLibraryFiles=Get-MasterBiasLibrary `
    -Path "W:\Astrophotography\BiasLibrary\QHY268M" `
    -Pattern "^(?<date>\d+).MasterBias.Gain.(?<gain>\d+).Offset.(?<offset>\d+).2CMS-1.USB33.(?<numberOfExposures>\d+)x(?<exposure>\d+\.?\d*)s.xisf$" |
    where-object Geometry -eq "6252:4176:1" |
    where-object ObsDateMinus12hr -gt "2023-03-22" |
    group-object Gain,Offset |
    foreach-object {
        $_.Group | sort-object NumberOfExposures -Descending | Select-Object -First 1
    }
$BiasLibraryFiles = $BiasLibraryFiles + (Get-MasterBiasLibrary `
    -Path "W:\Astrophotography\BiasLibrary\QHY268M" `
    -Pattern "^(?<date>\d+).MasterBias.Gain.(?<gain>\d+).Offset.(?<offset>\d+).(?<numberOfExposures>\d+)x(?<exposure>\d+\.?\d*)s.xisf$" |
    where-object Geometry -eq "6252:4176:1" |
    where-object ObsDateMinus12hr -gt ([DateTime]"2024-03-01") |
    group-object Gain,Offset |
    foreach-object {
        $_.Group | sort-object NumberOfExposures -Descending | Select-Object -First 1
    })
$BiasLibraryFiles|format-table UsbLimit,numberOfExposures,Instrument,Gain,Offset,ReadoutMode

$DarkLibraryFiles=Get-MasterDarkLibrary `
    -Path "W:\Astrophotography\DarkLibrary\QHY268M" `
    -Pattern "^(?<date>\d+).MasterDark.Gain.(?<gain>\d+).Offset.(?<offset>\d+).(?<temp>-?\d+)C.(?<numberOfExposures>\d+)x(?<exposure>\d+)s.xisf$" |
    where-object Geometry -eq "6252:4176:1" |
    where-object ObsDateMinus12hr -gt "2023-03-22" |
    group-object Instrument,Gain,Offset,SetTemp,Exposure |
    foreach-object {
        $_.Group | sort-object NumberOfExposures -Descending | select-object -First 1
    }
$DarkLibraryFiles | sort-object Gain,Offset,SetTemp,Exposure|format-table Instrument,Exposure,Gain,Offset,SetTemp,NumberOfExposures,Path
$DarkLibrary=($DarkLibraryFiles|group-object Instrument,Gain,Offset,Exposure,SetTemp|foreach-object {
    $instrument=$_.Group[0].Instrument
    $gain=$_.Group[0].Gain
    $offset=$_.Group[0].Offset
    $exposure=$_.Group[0].Exposure
    $setTemp=$_.Group[0].SetTemp
    
    $dark=$_.Group | sort-object {(Get-Item $_.Path).LastWriteTime} -Descending | select-object -First 1
    new-object psobject -Property @{
        Instrument=$instrument
        Gain=$gain
        Offset=$offset
        Exposure=$exposure
        SetTemp=$setTemp
        Path=$dark.Path
    }
})

#restack data from 12-30 and newer

# move all morning flats to "Flats" folder
Get-XisfFile -Path $DropoffLocation | 
    where-object ImageType -eq "Flat" | 
    where-object {$_.LocalDate.Hour -lt 12} |
    foreach-object {
        move-item $_.Path -Destination $DropoffLocation\Flats\
    }
$toCalibrate = 
    Get-XisfLightFrames -Path $DropoffLocation |
    group-object ObsDateMinus12hr,Filter |
    foreach-object {new-object psobject -property @{ObsDateMinus12hr=$_.Group[0].ObsDateMinus12hr;Filter=$_.Group[0].Filter}}
Get-XisfFile -Path $DropoffLocation | 
    where-object ImageType -eq "Flat" | 
    group-object ObsDateMinus12hr,Filter|
    foreach-object {
        $filter=$_.Group[0].Filter
        $obsDateMinus12hr=$_.Group[0].ObsDateMinus12hr
        if(-not ($toCalibrate | where-object Filter -eq $filter | where-object ObsDateMinus12hr -eq $obsDateMinus12hr)){
            write-host "No Lights on $obsDateMinus12hr with filter $filter"
            $_.Group | foreach-object {
                move-item $_.Path -Destination $DropoffLocation\Flats\ -whatif
            }
        }
    }
    <#
Get-XisfFile -Path $DropoffLocation | 
    where-object ImageType -eq "Flat" | 
    group-object ObsDateMinus12hr,Filter |
    select-object Name,Count
#>
Invoke-FlatFrameSorting `
    -DropoffLocation $DropoffLocation `
    -ArchiveDirectory $ArchiveDirectory `
    -CalibratedFlatsOutput "E:\Calibrated\CalibratedFlats" `
    -PixInsightSlot 201 -UseBias -BiasLibraryFiles $BiasLibraryFiles #-KeepOpen 
#exit
#Install-module ResizeImageModule
Import-Module ResizeImageModule
Import-module "C:\Program Files\N.I.N.A. - Nighttime Imaging 'N' Astronomy\NINA.Astrometry.dll"
if($null -eq $apiKey){
    $apiKey = Get-Credential -Message "Supply username and apiKey from url: https://app.lightbucket.co/api_credentials"
}
Function Update-LightBucketWithNewImageCaptured
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][System.IO.FileInfo]$LightFrame,
        [Parameter()][Byte[]]$ThumbnailData
    )
    $authString="$($apiKey.UserName):$($apiKey.GetNetworkCredential().Password)"
    $basicToken = [Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes($authString))
    $stats= $lightFrame|Get-XisfFitsStats
    $headers=$lightFrame|Get-XisfHeader
    $ra = [NINA.Astrometry.AstroUtil]::HMSToDegrees(($headers.xisf.Image.FITSKeyword |? name -eq "OBJCTRA").value)
    $dec = [NINA.Astrometry.AstroUtil]::DMSToDegrees(($headers.xisf.Image.FITSKeyword |? name -eq "OBJCTDEC").value)
    $base64ThumbnailData=$null
    if($ThumbnailData){
        $base64ThumbnailData=[Convert]::ToBase64String($ThumbnailData)
    }
    $request = new-object psobject -Property @{
        image= @{
            filter_name= $stats.Filter
            duration= $stats.Exposure
            gain= $stats.Gain
            offset= $stats.Offset
            binning= "1x1"
            captured_at= $stats.LocalDate.ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss')
            thumbnail=$base64ThumbnailData
        }
        target= @{
            name= $stats.Object
            ra= $ra
            dec= $dec
            rotation= ($headers.xisf.Image.FITSKeyword |? name -eq "OBJCTROT").value
        }
        equipment= @{
            telescope_name= ($headers.xisf.Image.FITSKeyword |? name -eq "TELESCOP").value
            camera_name=($headers.xisf.Image.FITSKeyword |? name -eq "INSTRUME").value
        }
    } |ConvertTo-Json -Depth 6
    #$request
    try {
        Invoke-WebRequest `
        -Uri "https://app.lightbucket.co/api/image_capture_complete" `
        -Method Post `
        -Body $request `
        -ContentType "application/json" `
        -Headers @{
            Authorization="Basic $basicToken"
        } 
    }
    catch {
        $result = $_.Exception.Response.GetResponseStream()
        $reader = New-Object System.IO.StreamReader($result)
        $reader.BaseStream.Position = 0
        $reader.DiscardBufferedData()
        $responseBody = $reader.ReadToEnd();
        Write-Warning "An unexpected error occured posting to lightbucket..."
        Write-Verbose $request
        write-warning $responseBody
        throw;
    }    
}

while($true){
    Get-ChildItem $DropoffLocation *.xisf -ErrorAction Continue |
        foreach-object { try{ $_ | Get-XisfFitsStats -ErrorAction Continue}catch{} }|
        where-object Instrument -eq "QHY268M" |
        where-object ImageType -eq "LIGHT" |
        where-object FocalLength -eq "950" |
        where-object Geometry -eq "6252:4176:1" |
        #where-object ReadoutMode -in @("High Gain 2CMS") |
        #where-object Object -ne "M 33" |
        #where-object ObsDateMinus12hr -le "2024-10-06" |
        #where-object Filter -eq "B" |
        group-object Instrument,SetTemp,Gain,Offset,Exposure,ObsDateMinus12hr |
        foreach-object {
            $lights = $_.Group
            $x=$lights[0]

            $instrument=$x.Instrument
            $gain=[decimal]$x.Gain
            $offset=[decimal]$x.Offset
            $exposure=[decimal]$x.Exposure
            $ccdTemp = [decimal]$x.CCDTemp
            $setTemp=[decimal]$x.SetTemp
            $obsDateMinus12hr=$x.ObsDateMinus12hr.ToString('yyyyMMdd')

            $masterBias = $BiasLibraryFiles |
                where-object Gain -eq $gain |
                where-object Offset -eq $offset |
                where-object Instrument -eq $instrument |
                sort-object ObsDate -Descending |
                select-object -First 1
            if($masterBias){
                write-host "Master bias available for $instrument at Gain=$gain Offset=$offset. $($masterBias.Path.Name)"
            }
            else{
                Write-Warning "No master bias available for $instrument at Gain=$gain Offset=$offset."
            }

            $masterDark = $DarkLibrary | where-object {
                $dark = $_
                ($dark.Instrument-eq $instrument) -and
                ($dark.Gain-eq $gain) -and
                ($dark.Offset-eq $offset) -and
                ($dark.Exposure-eq $exposure) -and
                ([Math]::abs($dark.SetTemp - $ccdTemp) -lt 5)
            } | select-object -first 1

            if(-not $masterDark){
                if($masterBias){
                    $masterDark = $DarkLibrary | where-object {
                        $dark = $_
                        ($dark.Instrument-eq $instrument) -and
                        ($dark.Gain-eq $gain) -and
                        ($dark.Offset-eq $offset) -and
                        ($dark.Exposure-eq $exposure) -and
                        ([Math]::abs($dark.SetTemp - $ccdTemp) -lt 9)
                    } | select-object -first 1
                    if($masterDark){
                        Write-Warning "No master dark available for $instrument at Gain=$gain Offset=$offset Exposure=$exposure (s) and SetTemp $setTemp. Scaling dark."
                    }
                    else{
                        Write-Warning "No master dark available for $instrument at Gain=$gain Offset=$offset Exposure=$exposure (s) and SetTemp $setTemp. Using bias only."
                        #return
                    }
                }
                else{
                    Write-Warning "No master dark available for $instrument at Gain=$gain Offset=$offset Exposure=$exposure (s) and SetTemp $setTemp. Skipping."
                    return 
                }
            }
            $lights |
                group-object Filter,FocalLength,FocalRatio |
                foreach-object {
                    $filter = $_.Group[0].Filter
                    $focalLength=$_.Group[0].FocalLength

                    $masterFlat ="W:\Astrophotography\$($focalLength)mm\Flats\$($obsDateMinus12hr).MasterFlatCal.$($filter).xisf"
                    #$masterFlat ="W:\Astrophotography\$($focalLength)mm\Flats\20241008.MasterFlatCal.$($filter).xisf"
                    
                    if($masterFlat -and (-not (test-path $masterFlat))) {
                        Write-Warning "Skipping $($_.Group.Count) frames at ($focalLength)mm with filter $filter. Reason: No master flat was found."
                    }
                    else{
                        $masterBiasFile=$masterBias.Path
                        $masterDarkFile=$masterDark.Path
                        $optimizeDark = ($masterBiasFile -and $masterDarkFile)
                        $calibrateDark= ($masterBiasFile -and $masterDarkFile)
                        Write-Host "Sorting $($_.Group.Count) frames at ($focalLength)mm with filter $filter"
                        if($masterBias){
                            Write-Host " Bias: $($masterBias.Path)"
                        }
                        if($masterDark) {
                            Write-Host " Dark: $($masterDark.Path)"
                        }
                        Write-Host " Flat: $($masterFlat)"
                        
                        Invoke-LightFrameSorting `
                            -XisfStats ($_.Group) -ArchiveDirectory $ArchiveDirectory `
                            -MasterBias $masterBiasFile -OptimizeDark:$OptimizeDark -CalibrateDark:$calibrateDark `
                            -MasterDark $masterDarkFile `
                            -MasterFlat $masterFlat `
                            -OutputPath $CalibratedOutput `
                            -PixInsightSlot 201 `
                            -OutputPedestal 80 `
                            -Verbose -KeepOpen `
                            -AfterImagesCalibrated {
                                param($LightFrames)

                                if(-not $PushToLightBucket){
                                    return;
                                }

                                $Last = $LightFrames | select-object -Last 1
                                $LightFrames | 
                                    where-object {-not [string]::IsNullOrWhiteSpace($_.Object)} |
                                    foreach-object {
                                        $LightFrame = $_
                                        $ThumbnailData=$null
                                        
                                        if($LightFrame -eq $Last){
                                            try{
                                                $OutputFolder = Join-Path $CalibratedOutput $LightFrame.Object.Trim()
                                                $calibrated=Get-CalibrationFile -Path ($LightFrame.Path) -CalibratedPath $OutputFolder
                                                $fileName = [System.IO.Path]::GetFileNameWithoutExtension($LightFrame.Path)
                                                $ThumbnailFolder = Join-Path $OutputFolder "Thumbnails"
                                                [System.IO.Directory]::CreateDirectory($ThumbnailFolder) >> $null
                                                $ThumbnailFile = Join-Path $ThumbnailFolder ($fileName+".jpeg")
                                                $ThumbnailSmall = Join-Path $ThumbnailFolder ($fileName+".small.jpeg")
                                                ConvertTo-XisfStfThumbnail -Path $calibrated -OutputPath $ThumbnailFile -PixInsightSlot 201
                                                Resize-Image -InputFile $ThumbnailFile -OutputFile $ThumbnailSmall -Width 300 -ProportionalResize $true -Height 300
                                                $ThumbnailData = Get-Content $ThumbnailSmall -AsByteStream
                                            }
                                            catch{
                                                write-warning $_.Exception.ToString()
                                            }
                                        }
                                        try {
                                            Update-LightBucketWithNewImageCaptured -LightFrame ($LightFrame.Path) -ThumbnailData $ThumbnailData
                                        }
                                        catch {
                                            if($_.scriptstacktrace){
                                                write-warning $_.scriptstacktrace
                                            }
                                            if($_.Exception){
                                                write-warning $_.Exception.ToString()
                                            }                                            
                                            throw
                                        }
                                    }
                            }
                            
                    }
            }
        }
    write-host "waiting for next set of data..."
    Wait-Event -Timeout 60
}