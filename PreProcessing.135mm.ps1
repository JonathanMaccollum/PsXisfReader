import-module $PSScriptRoot/PsXisfReader.psd1 -Force

$ErrorActionPreference="STOP"

$DropoffLocation = "D:\Backups\Camera\Dropoff\NINA"
$ArchiveDirectory="E:\Astrophotography"
$CalibratedOutput = "E:\PixInsightLT\Calibrated"

Invoke-DarkFrameSorting `
    -DropoffLocation $DropoffLocation `
    -ArchiveDirectory $ArchiveDirectory `
    -PixInsightSlot 200 `
    -Verbose
Invoke-DarkFlatFrameSorting `
    -DropoffLocation $DropoffLocation `
    -ArchiveDirectory $ArchiveDirectory `
    -PixInsightSlot 200
Invoke-FlatFrameSorting `
    -DropoffLocation $DropoffLocation `
    -ArchiveDirectory $ArchiveDirectory `
    -CalibratedFlatsOutput "E:\PixInsightLT\CalibratedFlats" `
    -PixInsightSlot 200

$DarkLibraryFiles =
    Get-ChildItem "E:\Astrophotography\DarkLibrary\ZWO ASI071MC Pro" "*MasterDark.Gain.*.Offset.*.*C*x*s.xisf" -File |
    Get-XisfFitsStats |
    % {
        # note: expecting file names with the pattern: "*MasterDark.Gain.___.Offset.___.-15C.__x__s.xisf"
        $x=$_
        $parts=$x.Path.Name.Split(".")

        $gain=$parts[$parts.IndexOf("Gain")+1]
        $offset=$parts[$parts.IndexOf("Offset")+1]
        $temp=$parts[$parts.IndexOf("Offset")+2]
        $exposure=$parts[$parts.IndexOf("Offset")+3]
        if(-not $x.SetTemp){
            $x.SetTemp=[decimal]$temp.TrimEnd("C")
        }
        if(-not $x.Gain) {
            $x.Gain=[decimal]$gain
        }
        if(-not $x.Offset){
            $x.Offset=[decimal]$offset
        }
        if(-not $x.Exposure){
            $x.Exposure=[decimal]$exposure.Split('x')[1].TrimEnd("s")
        }
        $x
    }

$DarkLibrary=($DarkLibraryFiles|group-object Instrument,Gain,Offset,Exposure,SetTemp|foreach-object {
    $instrument=$_.Group[0].Instrument
    $gain=$_.Group[0].Gain
    $offset=$_.Group[0].Offset
    $exposure=$_.Group[0].Exposure
    $setTemp=$_.Group[0].SetTemp
    
    $dark=$_.Group | sort-object {(Get-Item $_.Path).LastModifiedDate} -Descending | select-object -First 1
    new-object psobject -Property @{
        Instrument=$instrument
        Gain=$gain
        Offset=$offset
        Exposure=$exposure
        SetTemp=$setTemp
        Path=$dark.Path
    }
})
Get-ChildItem $DropoffLocation *.xisf |
    Get-XisfFitsStats | 
    where-object Instrument -eq "ZWO ASI071MC Pro" |
    where-object ImageType -eq "LIGHT" |
    group-object Instrument,SetTemp,Gain,Offset,Exposure |
    foreach-object {
        $lights = $_.Group
        $x=$lights[0]

        $instrument=$x.Instrument
        $gain=[decimal]$x.Gain
        $offset=[decimal]$x.Offset
        $exposure=[decimal]$x.Exposure
        $setTemp=[decimal]$x.SetTemp
        $masterDark = $DarkLibrary | where-object {
            $dark = $_
            ($dark.Instrument-eq $instrument) -and
            ($dark.Gain-eq $gain) -and
            ($dark.Offset-eq $offset) -and
            ($dark.Exposure-eq $exposure) -and
            ($dark.SetTemp-eq $setTemp)
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
                    #$masterFlat = "E:\Astrophotography\$($focalLength)mm\Flats\20200508.MasterFlatCal.$filter.xisf"
                    $masterFlat = "D:\Backups\Camera\2019\20191130 - 135mm 071mc L3\MasterFlat.L3.xisf"

                    if(-not (test-path $masterFlat)) {
                        Write-Warning "Skipping $($_.Group.Count) frames at ($focalLength)mm with filter $filter. Reason: No master flat was found."
                    }
                    else{

                        Write-Host "Sorting $($_.Group.Count) frames at ($focalLength)mm with filter $filter"

                        Invoke-LightFrameSorting `
                            -XisfStats ($_.Group) -ArchiveDirectory $ArchiveDirectory `
                            -MasterDark ($masterDark.Path) `
                            -MasterFlat $masterFlat `
                            -OutputPath $CalibratedOutput `
                            -PixInsightSlot 200 `
                            -OutputPedestal 0 `
                            -Verbose
                    }
                }
        }
    }





exit
Invoke-LightFrameSorting `
    -DropoffLocation $DropoffLocation -ArchiveDirectory $ArchiveDirectory `
    -MasterDark "D:\Backups\Camera\2019\Dark Library\20190313.60x30s.masterdark.-10c.90g.65o.nobias.integration.xisf" `
    -MasterFlat "D:\Backups\Camera\2019\20191130 - 135mm 071mc L3\MasterFlat.L3.xisf" `
    -Filter "L3" -FocalLength 135 -Exposure 30 -Gain 90 -Offset 65 -SetTemp -10 `
    -OutputPath $CalibratedOutput `
    -PixInsightSlot 200
    
Invoke-LightFrameSorting `
    -DropoffLocation $DropoffLocation -ArchiveDirectory $ArchiveDirectory `
    -MasterDark "D:\Backups\Camera\2019\Dark Library\20190313.60x60s.masterdark.-10c.90g.65o.nobias.xisf" `
    -MasterFlat "D:\Backups\Camera\2019\20191130 - 135mm 071mc L3\MasterFlat.L3.xisf" `
    -Filter "L3" -FocalLength 135 -Exposure 60 -Gain 90 -Offset 65 -SetTemp -10 `
    -OutputPath $CalibratedOutput `
    -PixInsightSlot 200

Invoke-LightFrameSorting `
    -DropoffLocation $DropoffLocation -ArchiveDirectory $ArchiveDirectory `
    -MasterDark "D:\Backups\Camera\2019\Dark Library\20190313.60x120s.masterdark.-10c.90g.65o.nobias.xisf" `
    -MasterFlat "D:\Backups\Camera\2019\20191130 - 135mm 071mc L3\MasterFlat.L3.xisf" `
    -Filter "L3" -FocalLength 135 -Exposure 120 -Gain 90 -Offset 65 -SetTemp -10 `
    -OutputPath $CalibratedOutput `
    -PixInsightSlot 200
    
Invoke-LightFrameSorting `
    -DropoffLocation $DropoffLocation -ArchiveDirectory $ArchiveDirectory `
    -MasterDark "D:\Backups\Camera\ASI071mc-pro\MasterDark.NINA.240s.-15C_x20.Gain90Offset65.NoCalibration.xisf" `
    -MasterFlat "D:\Backups\Camera\2019\20191130 - 135mm 071mc L3\MasterFlat.L3.xisf" `
    -OutputPath $CalibratedOutput `
    -Filter "L3" -FocalLength 135 -Exposure 240 -Gain 90 -Offset 65 -SetTemp -15 `
    -PixInsightSlot 200