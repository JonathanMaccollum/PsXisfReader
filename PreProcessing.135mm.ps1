import-module $PSScriptRoot/PsXisfReader.psm1 -Force
import-module $PSScriptRoot/pixinsightpreprocessing.psm1 -Force
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