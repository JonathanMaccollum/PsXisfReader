if (-not (get-module psxisfreader)){import-module psxisfreader}

$ErrorActionPreference="STOP"

$DropoffLocation = "D:\Backups\Camera\Dropoff\NINACS"
$ArchiveDirectory="E:\Astrophotography"
$CalibratedOutput = "F:\PixInsightLT\Calibrated"

<#
Invoke-DarkFrameSorting `
    -DropoffLocation $DropoffLocation `
    -ArchiveDirectory $ArchiveDirectory `
    -PixInsightSlot 201 `
    -Verbose -KeepOpen
Invoke-DarkFlatFrameSorting `
    -DropoffLocation $DropoffLocation `
    -ArchiveDirectory $ArchiveDirectory `
    -PixInsightSlot 201
Invoke-FlatFrameSorting `
    -DropoffLocation $DropoffLocation `
    -ArchiveDirectory $ArchiveDirectory `
    -CalibratedFlatsOutput "F:\PixInsightLT\CalibratedFlats" `
    -PixInsightSlot 201
#>
$DarkLibraryFiles=Get-MasterDarkLibrary `
    -Path "E:\Astrophotography\DarkLibrary\ZWO ASI071MC Pro" `
    -Pattern "^(?<date>\d+).MasterDark.Gain.(?<gain>\d+).Offset.(?<offset>\d+).(?<temp>-?\d+)C.(?<numberOfExposures>\d+)x(?<exposure>\d+)s.xisf$"
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

    Get-ChildItem $DropoffLocation *.xisf |
        Get-XisfFitsStats | 
        where-object Instrument -eq "ZWO ASI071MC Pro" |
        where-object ImageType -eq "LIGHT" |
        where-object FocalLength -eq "90" |
        where-object Object -ne "StarTrails.20220130" |
        #where-object Object -eq "LBN406 in Draco" |
        #Select-Object -Last 1 |
        #where-object Object -eq "M64 - Black Eye Galaxy Widefield" |
        group-object Instrument,SetTemp,Gain,Offset,Exposure |
        foreach-object {
            $lights = $_.Group
            $x=$lights[0]

            $instrument=$x.Instrument
            $gain=[decimal]$x.Gain
            $offset=[decimal]$x.Offset
            $exposure=[decimal]$x.Exposure
            $ccdTemp = [decimal]$x.CCDTemp
            $setTemp=[decimal]$x.SetTemp
            $masterDark = $DarkLibrary | where-object {
                $dark = $_
                ($dark.Instrument-eq $instrument) -and
                ($dark.Gain-eq $gain) -and
                ($dark.Offset-eq $offset) -and
                ($dark.Exposure-eq $exposure) -and
                #($dark.SetTemp -eq $setTemp)
                ([Math]::abs($dark.SetTemp - $ccdTemp) -lt 3)
            } | select-object -first 1

            if(-not $masterDark){
                Write-Warning "Unable to process $($lights.Count) images: No master dark available for $instrument at Gain=$gain Offset=$offset Exposure=$exposure (s) and SetTemp $setTemp"
            }else {
                Write-Host "Master dark available for $instrument at Gain=$gain Offset=$offset Exposure=$exposure (s) and SetTemp $setTemp"
                $lights |
                    group-object Filter,FocalLength |
                    foreach-object {
                        $filter = $_.Group[0].Filter
                        $focalLength=$_.Group[0].FocalLength
                        #$masterFlat = "E:\Astrophotography\$($focalLength)mm\Flats\20220227.MasterFlatCal.$filter.xisf"
                        $masterFlat = "E:\Astrophotography\$($focalLength)mm\Flats\20220425.MasterFlatCal.$filter.F4.Short2.xisf"
                        

                        if($masterFlat -and (-not (test-path $masterFlat))) {
                            Write-Warning "Skipping $($_.Group.Count) frames at ($focalLength)mm with filter $filter. Reason: No master flat was found."
                        }
                        else{

                            Write-Host "Sorting $($_.Group.Count) frames at ($focalLength)mm with filter $filter"
                            Write-Host " Dark: $($masterDark.Path)"
                            Write-Host " Flat: $($masterFlat)"

                            Invoke-LightFrameSorting `
                                -XisfStats ($_.Group) -ArchiveDirectory $ArchiveDirectory `
                                -MasterDark ($masterDark.Path) `
                                -MasterFlat $masterFlat `
                                -OutputPath $CalibratedOutput `
                                -PixInsightSlot 201 `
                                -OutputPedestal 50 `
                                -Verbose `
                                -AfterImagesCalibrated {
                                    param($LightFrames)

                                    $Last = $LightFrames | select-object -Last 1
                                    $LightFrames | foreach-object {
                                        $LightFrame = $_
                                        $ThumbnailData=$null
                                        
                                        if($LightFrame -eq $Last){
                                            try{
                                                $OutputFolder = Join-Path $CalibratedOutput $LightFrame.Object
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
                                            write-warning $_.Exception.ToString()
                                            throw
                                        }
                                    }
                                }

                        }
                    }
            }
        }
    write-host "waiting for next set of data..."
    Wait-Event -Timeout 60
}