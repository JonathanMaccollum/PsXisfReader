    #import-module $PSScriptRoot/PsXisfReader.psd1 -Force
    if (-not (get-module psxisfreader)){import-module psxisfreader}
    $ErrorActionPreference="STOP"

    $DropoffLocation = "D:\Backups\Camera\Dropoff\NINA"
    $ArchiveDirectory="E:\Astrophotography\50mm"
    

    Get-ChildItem $DropoffLocation Timelapse*.xisf |
        foreach-object { try{ $_ | Get-XisfFitsStats -ErrorAction Continue}catch{} }|
        where-object Instrument -eq "ZWO ASI071MC Pro" |
        where-object ImageType -eq "LIGHT" |
        where-object FocalLength -eq "50" |
        group-object Object,Exposure |
        foreach-object {
            $batch = $_.Group
            $Object = $batch[0].Object
            $Exposure = $batch[0].Exposure

            $destination = join-path $ArchiveDirectory $Object "$($Exposure)s" 
            [IO.Directory]::CreateDirectory($destination) >> $null
            $itemsMoved = $batch.Path | 
                foreach-object { move-item -Path $_.FullName -Destination $destination -PassThru }
            
            #Cos cor
            #Invoke-PiDebayer -PixInsightSlot 202 -Images $itemsMoved -OutputPath S:\PixInsight\Debayered -CfaPattern "RGGB"
            #HT Fixed Stretch
            #Sample Format 16bit
            #Save as Tiff
        }
        
