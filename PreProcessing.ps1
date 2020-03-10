import-module $PSScriptRoot/pixinsightpreprocessing.psm1 -Force
import-module $PSScriptRoot/PsXisfReader.psm1


$DropoffLocation = "D:\Backups\Camera\Dropoff\NINA"
$ArchiveDirectory="D:\Backups\Camera\Astrophotography"
Invoke-DarkFlatFrameSorting `
    -DropoffLocation $DropoffLocation `
    -ArchiveDirectory $ArchiveDirectory `
    -PixInsightSlot 200
Invoke-FlatFrameSorting `
    -DropoffLocation $DropoffLocation `
    -ArchiveDirectory $ArchiveDirectory `
    -CalibratedFlatsOutput "S:\PixInsight\CalibratedFlats" `
    -PixInsightSlot 200
Invoke-LightFrameSorting `
    -DropoffLocation $DropoffLocation -ArchiveDirectory $ArchiveDirectory `
    -MasterDark "D:\Backups\Camera\ASI071mc-pro\MasterDark.NINA.240s.-15C_x20.Gain90Offset65.NoCalibration.xisf" `
    -MasterFlat "D:\Backups\Camera\2019\20191130 - 135mm 071mc L3\MasterFlat.L3.xisf" `
    -OutputPath "S:\PixInsight\Calibrated" `
    -Filter "L3" -FocalLength 135 -Exposure 240 -Gain 90 -Offset 65 `
    -PixInsightSlot 200
Invoke-LightFrameSorting `
    -DropoffLocation $DropoffLocation -ArchiveDirectory $ArchiveDirectory `
    -MasterDark "D:\Backups\Camera\2019\Dark Library\20190313.60x120s.masterdark.-10c.90g.65o.nobias.xisf" `
    -MasterFlat "D:\Backups\Camera\2019\20191130 - 135mm 071mc L3\MasterFlat.L3.xisf" `
    -Filter "L3" -FocalLength 135 -Exposure 120 -Gain 90 -Offset 65 `
    -OutputPath "S:\PixInsight\Calibrated" `
    -PixInsightSlot 200
Invoke-LightFrameSorting `
    -DropoffLocation $DropoffLocation -ArchiveDirectory $ArchiveDirectory `
    -MasterDark "D:\Backups\Camera\2019\Dark Library\20190313.60x30s.masterdark.-10c.90g.65o.nobias.integration.xisf" `
    -MasterFlat "D:\Backups\Camera\2019\20191130 - 135mm 071mc L3\MasterFlat.L3.xisf" `
    -Filter "L3" -FocalLength 135 -Exposure 30 -Gain 90 -Offset 65 `
    -OutputPath "S:\PixInsight\Calibrated" `
    -PixInsightSlot 200


Invoke-LightFrameSorting `
    -DropoffLocation $DropoffLocation -ArchiveDirectory $ArchiveDirectory `
    -MasterDark "D:\Backups\Camera\ASI183mm-Pro\20190909\25x60s.-15C.G111.O8.masterdark.nocalibration.xisf" `
    -MasterFlat "D:\Backups\Camera\Astrophotography\1000mm\Flats\20200306.MasterFlatCal.R.xisf" `
    -Filter "R" -FocalLength 1000 -Exposure 60 -Gain 111 -Offset 8 `
    -OutputPath "S:\PixInsight\Calibrated" `
    -PixInsightSlot 200
Invoke-LightFrameSorting `
    -DropoffLocation $DropoffLocation -ArchiveDirectory $ArchiveDirectory `
    -MasterDark "D:\Backups\Camera\ASI183mm-Pro\20190909\25x60s.-15C.G111.O8.masterdark.nocalibration.xisf" `
    -MasterFlat "D:\Backups\Camera\Astrophotography\1000mm\Flats\20200306.MasterFlatCal.G.xisf" `
    -Filter "G" -FocalLength 1000 -Exposure 60 -Gain 111 -Offset 8  `
    -OutputPath "S:\PixInsight\Calibrated" `
    -PixInsightSlot 200
Invoke-LightFrameSorting `
    -DropoffLocation $DropoffLocation -ArchiveDirectory $ArchiveDirectory `
    -MasterDark "D:\Backups\Camera\ASI183mm-Pro\20190909\25x60s.-15C.G111.O8.masterdark.nocalibration.xisf" `
    -MasterFlat "D:\Backups\Camera\Astrophotography\1000mm\Flats\20200306.MasterFlatCal.B.xisf" `
    -Filter "B" -FocalLength 1000 -Exposure 60 -Gain 111 -Offset 8  `
    -OutputPath "S:\PixInsight\Calibrated" `
    -PixInsightSlot 200
Invoke-LightFrameSorting `
    -DropoffLocation $DropoffLocation -ArchiveDirectory $ArchiveDirectory `
    -MasterDark "D:\Backups\Camera\ASI183mm-Pro\20190525\60x30s.-10C.G111.O8.masterdark.nocalibration.xisf" `
    -MasterFlat "D:\Backups\Camera\Astrophotography\1000mm\Flats\20200306.MasterFlatCal.L.xisf" `
    -Filter "L" -FocalLength 1000 -Exposure 30 -Gain 111 -Offset 8  `
    -OutputPath "S:\PixInsight\Calibrated" `
    -PixInsightSlot 200

$('Ha','L','R','G','B')|foreach-object {
    Invoke-LightFrameSorting `
        -DropoffLocation $DropoffLocation -ArchiveDirectory $ArchiveDirectory `
        -MasterDark "D:\Backups\Camera\ASI183mm-Pro\20191220\30x240s.-15C.G111.O8.masterdark.nocalibration.xisf" `
        -MasterFlat "D:\Backups\Camera\Astrophotography\1000mm\Flats\20200306.MasterFlatCal.$_.xisf" `
        -Filter $_ -FocalLength 1000 -Exposure 240 -Gain 111 -Offset 8  `
        -OutputPath "S:\PixInsight\Calibrated" `
        -PixInsightSlot 200
}

$('Ha','L','R','G','B')|foreach-object {
    Invoke-LightFrameSorting `
        -DropoffLocation $DropoffLocation -ArchiveDirectory $ArchiveDirectory `
        -MasterDark "D:\Backups\Camera\ASI183mm-Pro\20191220\30x120s.-15C.G111.O8.masterdark.nocalibration.xisf" `
        -MasterFlat "D:\Backups\Camera\Astrophotography\1000mm\Flats\20200306.MasterFlatCal.$_.xisf" `
        -Filter $_ -FocalLength 1000 -Exposure 120 -Gain 111 -Offset 8  `
        -OutputPath "S:\PixInsight\Calibrated" `
        -PixInsightSlot 200
}

$('Ha','Oiii','Sii')|foreach-object {
    Invoke-LightFrameSorting `
        -DropoffLocation $DropoffLocation -ArchiveDirectory $ArchiveDirectory `
        -MasterDark "D:\Backups\Camera\ASI183mm-Pro\20191220\74x600s.-15C.G53.O10.masterdark.nocalibration.xisf" `
        -MasterFlat "D:\Backups\Camera\Astrophotography\1000mm\Flats\20200306.MasterFlatCal.$_.xisf" `
        -Filter $_ -FocalLength 1000 -Exposure 600 -Gain 53 -Offset 10  `
        -OutputPath "S:\PixInsight\Calibrated" `
        -PixInsightSlot 200
}
