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
    -MasterDark "D:\Backups\Camera\ASI071mc-pro\MasterDark.NINA.240s.-15C_x20.Gain90Offset65.NoCalibration.xisf" `
    -MasterFlat "D:\Backups\Camera\2019\20191130 - 135mm 071mc L3\MasterFlat.L3.xisf" `
    -OutputPath $CalibratedOutput `
    -Filter "L3" -FocalLength 135 -Exposure 240 -Gain 90 -Offset 65 -SetTemp -15 `
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
    -MasterDark "D:\Backups\Camera\2019\Dark Library\20190313.60x30s.masterdark.-10c.90g.65o.nobias.integration.xisf" `
    -MasterFlat "D:\Backups\Camera\2019\20191130 - 135mm 071mc L3\MasterFlat.L3.xisf" `
    -Filter "L3" -FocalLength 135 -Exposure 30 -Gain 90 -Offset 65 -SetTemp -10 `
    -OutputPath $CalibratedOutput `
    -PixInsightSlot 200

$('L','R','G','B')|foreach-object {
    Invoke-LightFrameSorting `
    -DropoffLocation $DropoffLocation -ArchiveDirectory $ArchiveDirectory `
    -MasterDark "E:\Astrophotography\DarkLibrary\ZWO ASI183MM Pro\20200328.MasterDark.-15C.60x15s.xisf" `
    -MasterFlat "E:\Astrophotography\1000mm\Flats\20200324.MasterFlatCal.$_.xisf" `
    -Filter $_ -FocalLength 1000 -Exposure 15 -Gain 111 -Offset 8 -SetTemp -15  `
    -OutputPath $CalibratedOutput `
    -PixInsightSlot 200 `
    -Verbose
}
$('L','R','G','B')|foreach-object {
    Invoke-LightFrameSorting `
    -DropoffLocation $DropoffLocation -ArchiveDirectory $ArchiveDirectory `
    -MasterDark "E:\Astrophotography\DarkLibrary\ZWO ASI183MM Pro\20200329.MasterDark.-5C.60x60s.xisf" `
    -MasterFlat "E:\Astrophotography\1000mm\Flats\20200324.MasterFlatCal.$_.xisf" `
    -Filter $_ -FocalLength 1000 -Exposure 60 -Gain 111 -Offset 8 -SetTemp -5  `
    -OutputPath $CalibratedOutput `
    -PixInsightSlot 200 `
    -Verbose
}
$('L','R','G','B')|foreach-object {
    Invoke-LightFrameSorting `
    -DropoffLocation $DropoffLocation -ArchiveDirectory $ArchiveDirectory `
    -MasterDark "D:\Backups\Camera\ASI183mm-Pro\20190525\60x30s.-10C.G111.O8.masterdark.nocalibration.xisf" `
    -MasterFlat "E:\Astrophotography\1000mm\Flats\20200324.MasterFlatCal.$_.xisf" `
    -Filter $_ -FocalLength 1000 -Exposure 30 -Gain 111 -Offset 8 -SetTemp -10  `
    -OutputPath $CalibratedOutput `
    -PixInsightSlot 200
}
$('L','R','G','B')|foreach-object {
    Invoke-LightFrameSorting `
    -DropoffLocation $DropoffLocation -ArchiveDirectory $ArchiveDirectory `
    -MasterDark "D:\Backups\Camera\ASI183mm-Pro\20190909\25x60s.-15C.G111.O8.masterdark.nocalibration.xisf" `
    -MasterFlat "E:\Astrophotography\1000mm\Flats\20200324.MasterFlatCal.$_.xisf" `
    -Filter $_ -FocalLength 1000 -Exposure 60 -Gain 111 -Offset 8 -SetTemp -15 `
    -OutputPath $CalibratedOutput `
    -PixInsightSlot 200
}
$('L','R','G','B')|foreach-object {
    Invoke-LightFrameSorting `
        -DropoffLocation $DropoffLocation -ArchiveDirectory $ArchiveDirectory `
        -MasterDark "D:\Backups\Camera\ASI183mm-Pro\20191220\30x240s.-15C.G111.O8.masterdark.nocalibration.xisf" `
        -MasterFlat "E:\Astrophotography\1000mm\Flats\20200324.MasterFlatCal.$_.xisf" `
        -Filter $_ -FocalLength 1000 -Exposure 240 -Gain 111 -Offset 8 -SetTemp -15  `
        -OutputPath $CalibratedOutput `
        -PixInsightSlot 200
}
$('L','R','G','B')|foreach-object {
    Invoke-LightFrameSorting `
        -DropoffLocation $DropoffLocation -ArchiveDirectory $ArchiveDirectory `
        -MasterDark "D:\Backups\Camera\ASI183mm-Pro\20191220\30x120s.-15C.G111.O8.masterdark.nocalibration.xisf" `
        -MasterFlat "E:\Astrophotography\1000mm\Flats\20200324.MasterFlatCal.$_.xisf" `
        -Filter $_ -FocalLength 1000 -Exposure 120 -Gain 111 -Offset 8 -SetTemp -15  `
        -OutputPath $CalibratedOutput `
        -PixInsightSlot 200
}
$('Ha','Oiii','Sii')|foreach-object {
    Invoke-LightFrameSorting `
        -DropoffLocation $DropoffLocation -ArchiveDirectory $ArchiveDirectory `
        -MasterDark "D:\Backups\Camera\ASI183mm-Pro\20191220\74x600s.-15C.G53.O10.masterdark.nocalibration.xisf" `
        -MasterFlat "E:\Astrophotography\1000mm\Flats\20200306.MasterFlatCal.$_.xisf" `
        -Filter $_ -FocalLength 1000 -Exposure 600 -Gain 53 -Offset 10 -SetTemp -15  `
        -OutputPath $CalibratedOutput `
        -PixInsightSlot 200
}
$('Ha')|foreach-object {
    Invoke-LightFrameSorting `
        -DropoffLocation $DropoffLocation -ArchiveDirectory $ArchiveDirectory `
        -MasterDark "D:\Backups\Camera\ASI183mm-Pro\20191220\30x240s.-15C.G111.O8.masterdark.nocalibration.xisf" `
        -MasterFlat "E:\Astrophotography\1000mm\Flats\20200306.MasterFlatCal.$_.xisf" `
        -Filter $_ -FocalLength 1000 -Exposure 240 -Gain 111 -Offset 8 -OutputPedestal 150 -SetTemp -15 `
        -OutputPath $CalibratedOutput `
        -PixInsightSlot 200
}