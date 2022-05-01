
Import-module "C:\Program Files\N.I.N.A. - Nighttime Imaging 'N' Astronomy\NINA.Astrometry.dll"

if($null -eq $apiKey){
    $apiKey = Get-Credential -Message "Supply username and apiKey from url: https://staging.lightbucket.co/api_credentials"
}

# $lastSub = Get-ChildItem "D:\Backups\Camera\Dropoff\NINA" *.xisf |
#     sort-object LastWriteTime -Descending |
#     Select-Object -First 1 
# $stats= $lastSub|Get-XisfFitsStats
# $headers=$lastSub|Get-XisfHeader

cls

$authString="$($apiKey.UserName):$($apiKey.GetNetworkCredential().Password)"
$basicToken = [Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes($authString))

    Get-XisfLightFrames -Path "E:\Astrophotography\1000mm" -Recurse -UseCache -SkipOnError |
    Where-Object {-not $_.HasTokensInPath(@("reject","process","testing","clouds","draft","cloudy","_ez_LS_","drizzle"))} |
    Where-Object {$_.LocalDate -and $_.LocalDate -gt [DateTime]::Today.AddMonths(-1)} |
    #Where-Object {$_.LocalDate -and $_.LocalDate -gt [DateTime]("2020-10-01") -and $_.LocalDate -lt [DateTime]("2020-10-31") } |
    Where-Object {-not $_.IsIntegratedFile()} |
    sort-object LocalDate |
    foreach-object {
        $stats= $_
        $headers=$_.Path|Get-Item|Get-XisfHeader
        $ra = [NINA.Astrometry.AstroUtil]::HMSToDegrees(($headers.xisf.Image.FITSKeyword |? name -eq "OBJCTRA").value)
        $dec = [NINA.Astrometry.AstroUtil]::DMSToDegrees(($headers.xisf.Image.FITSKeyword |? name -eq "OBJCTDEC").value)
        $request = new-object psobject -Property @{
            image= @{
                filter_name= $stats.Filter
                duration= $stats.Exposure
                gain= $stats.Gain
                offset= $stats.Offset
                binning= "1x1"
                captured_at= $stats.LocalDate.ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss')
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
        Invoke-WebRequest `
             -Uri "https://staging.lightbucket.co/api/image_capture_complete" `
             -Method Post `
             -Body $request `
             -ContentType "application/json" `
             -Headers @{
                 Authorization="Basic $basicToken"
             } 
    }
