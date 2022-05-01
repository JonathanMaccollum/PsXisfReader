import-module PsXisfReader -force

get-childitem "E:\Astrophotography\BiasLibrary\QHYCCD-Cameras-Capture (ASCOM)\20210215" "*.xisf"
    | foreach-object {
        $p=$_|Get-XisfProperties
    } 