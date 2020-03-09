import-module $PSScriptRoot/pixinsightpreprocessing.psm1 -Force
import-module $PSScriptRoot/PsXisfReader.psm1


$PixInsightSlot=200
$DropoffLocation = "D:\Backups\Camera\Dropoff\NINA"
$ArchiveDirectory="D:\Backups\Camera\Astrophotography"
$CalibratedFlatsOutput=Get-Item "S:\PixInsight\CalibratedFlats"
$CalibratedLightsOutput=Get-Item "S:\PixInsight\Calibrated"

$DarkLibrary = (Get-ChildItem D:\Backups\Camera\ASI071mc-pro "*Dark*.xisf")
$DarkLibrary += (Get-ChildItem 'D:\Backups\Camera\2019\Dark Library' "*.xisf")
$DarkLibrary = $DarkLibrary | Foreach-Object {Get-XisfFitsStats -Path $_ }

$FilesToProcess=
    Get-ChildItem $DropoffLocation "*.xisf" |
    foreach-object { Get-XisfFitsStats -Path $_ } |
    group-object ImageType

($FilesToProcess|where-object Name -eq "DARKFLAT").Group |
    group-object Filter,FocalLength |
    foreach-object {
        $images=$_.Group
        $filter=$images[0].Filter.Trim()
        $focalLength=($images[0].FocalLength).TrimEnd('mm')+'mm'
        $flatDate = ([DateTime]$images[0].LocalDate).Date.ToString('yyyyMMdd')

        Write-Host "Processing Dark Flats for $filter at $focalLength"

        $targetDirectory = "$ArchiveDirectory\$focalLength\Flats"
        $masterDark = "$targetDirectory\$($flatDate).MasterDarkFlat.$($filter).xisf"
        if(-not (Test-Path $masterDark)) {
            $darkFlats = $images | foreach-object {$_.Path}
            Invoke-PiDarkIntegration `
                -Images $darkFlats `
                -PixInsightSlot $PixInsightSlot `
                -OutputFile $masterDark `
                -Verbose
        }
        $images | Foreach-Object {Move-Item -Path $_.Path -Destination "$targetDirectory\$flatDate"}
    }

Function Add-Calibrated
{
    [CmdletBinding()]
    param (
        [Parameter()]$File,
        [Parameter()][System.IO.DirectoryInfo]$CalibratedPath
    )
    $input = $File.Path
    $calibrated = Join-Path ($CalibratedPath.FullName) ($input.Name.Replace($input.Extension,"")+"_c.xisf")
    Add-Member -InputObject $_ -Name "Calibrated" -MemberType NoteProperty -Value $calibrated -Force
    $File
}

($FilesToProcess|where-object Name -eq "FLAT").Group |
    group-object Filter,FocalLength |
    foreach-object {
        $flats=$_.Group
        $filter=$flats[0].Filter.Trim()
        $focalLength=($flats[0].FocalLength).TrimEnd('mm')+'mm'
        $flatDate = ([DateTime]$flats[0].LocalDate).Date.ToString('yyyyMMdd')

        Write-Host "Processing Flats for $filter at $focalLength"

        $targetDirectory = "$ArchiveDirectory\$focalLength\Flats"
        $masterDark = "$targetDirectory\$flatDate.MasterDarkFlat.$filter.xisf"
        if(Test-Path $masterDark) {
            $masterCalibratedFlat = "$targetDirectory\$flatDate.MasterFlatCal.$filter.xisf"
            if(-not (Test-Path $masterCalibratedFlat)) {
                $flats | foreach-object{
                    $input = $_.Path
                    $output = Join-Path ($CalibratedFlatsOutput.FullName) ($input.Name.Replace($input.Extension,"")+"_c.xisf")
                    Add-Member -InputObject $_ -Name "CalibratedFlat" -MemberType NoteProperty -Value $output -Force
                }
                $toCalibrate = $flats | where-object {-not (Test-Path $_.CalibratedFlat)} | foreach-object {$_.Path}
                if($toCalibrate){
                    Invoke-PiFlatCalibration `
                        -Images $toCalibrate `
                        -MasterDark $masterDark `
                        -OutputPath $CalibratedFlatsOutput `
                        -PixInsightSlot $PixInsightSlot -Verbose
                }
                $calibratedFlats = $flats|foreach-object {Get-Item $_.CalibratedFlat}
                if($flats) {
                    Invoke-PiFlatIntegration `
                        -Images $calibratedFlats `
                        -PixInsightSlot $PixInsightSlot `
                        -OutputFile $masterCalibratedFlat `
                        -Verbose
                }
            }
        }

        $masterNoCalFlat = "$TargetDirectory\$FlatDate.MasterFlatNoCal.$($filter).xisf"
        $toIntegrate = $flats|foreach-object {$_.Path}
        if(-not (Test-Path $masterNoCalFlat)) {
            Invoke-PiFlatIntegration `
                -Images $toIntegrate `
                -PixInsightSlot $PixInsightSlot `
                -OutputFile $masterNoCalFlat `
                -Verbose
        }
        $flats | Foreach-Object {Move-Item -Path $_.Path -Destination "$targetDirectory\$flatDate"}
    }

Function Invoke-LightFrameSorting
{
    [CmdletBinding()]
    param (
        [Object[]]$Lights,
        [String]$Filter,
        [int]$FocalLength,
        [int]$Exposure,
        [System.IO.FileInfo]$MasterDark,
        [System.IO.FileInfo]$MasterFlat,
        [System.IO.DirectoryInfo]$OutputPath,
        [int]$PixInsightSlot
    )

    $Lights | 
        where-object Filter -eq $Filter |
        where-object Exposure -eq $Exposure |
        where-object FocalLength -eq $FocalLength |
        Group-Object Object |
        foreach-object {
            $object = $_.Name

            $toCalibrate = $_.Group |
                foreach-object {Add-Calibrated -File $_ -CalibratedPath $CalibratedLightsOutput } |
                where-object {-not (Test-Path $_.Calibrated)} | foreach-object {$_.Path}
            if($toCalibrate){
                Invoke-PiLightCalibration `
                    -Images $toCalibrate `
                    -MasterDark $MasterDark `
                    -MasterFlat $MasterFlat `
                    -OutputPath $CalibratedLightsOutput `
                    -PixInsightSlot $PixInsightSlot -Verbose
            }

            $objectNameParts = $object.Split(' ')
            $archive = Join-Path $ArchiveDirectory "$($FocalLength)mm" -AdditionalChildPath $object
            if(-not (test-path $archive)){
                $archive = Join-Path $ArchiveDirectory "$($FocalLength)mm" -AdditionalChildPath ($objectNameParts[0] + ' '+ $objectNameParts[1])
            }
            if(-not (test-path $archive)){
                $archive = Join-Path $ArchiveDirectory "$($FocalLength)mm" -AdditionalChildPath ($objectNameParts[0])
            }
            $archive = Join-Path $archive "$Exposure.00"
            [System.IO.Directory]::CreateDirectory($archive) >> $null

            $_.Group | Foreach-Object {Move-Item -Path $_.Path -Destination $archive}
        }
}


Invoke-LightFrameSorting `
    -Lights (($FilesToProcess|where-object Name -eq "Light").Group) `
    -MasterDark "D:\Backups\Camera\ASI071mc-pro\MasterDark.NINA.240s.-15C_x20.Gain90Offset65.NoCalibration.xisf" `
    -MasterFlat "D:\Backups\Camera\2019\20191130 - 135mm 071mc L3\MasterFlat.L3.xisf" `
    -Filter "L3" `
    -FocalLength 135 `
    -Exposure 240 `
    -OutputPath $CalibratedLightsOutput `
    -PixInsightSlot $PixInsightSlot `
    -Verbose

Invoke-LightFrameSorting `
    -Lights (($FilesToProcess|where-object Name -eq "Light").Group) `
    -MasterDark "D:\Backups\Camera\2019\Dark Library\20190313.60x120s.masterdark.-10c.90g.65o.nobias.xisf" `
    -MasterFlat "D:\Backups\Camera\2019\20191130 - 135mm 071mc L3\MasterFlat.L3.xisf" `
    -Filter "L3" `
    -FocalLength 135 `
    -Exposure 120 `
    -OutputPath $CalibratedLightsOutput `
    -PixInsightSlot $PixInsightSlot `
    -Verbose
Invoke-LightFrameSorting `
    -Lights (($FilesToProcess|where-object Name -eq "Light").Group) `
    -MasterDark "D:\Backups\Camera\2019\Dark Library\20190313.60x30s.masterdark.-10c.90g.65o.nobias.integration.xisf" `
    -MasterFlat "D:\Backups\Camera\2019\20191130 - 135mm 071mc L3\MasterFlat.L3.xisf" `
    -Filter "L3" `
    -FocalLength 135 `
    -Exposure 30 `
    -OutputPath $CalibratedLightsOutput `
    -PixInsightSlot $PixInsightSlot `
    -Verbose
