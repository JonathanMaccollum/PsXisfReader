$manif=Invoke-WebRequest https://nighttime-imaging.eu/wp-json/nina/v1/plugins/manifest

$manif | 
    ConvertFrom-Json |
    #group-object Name
    #where-object Name -eq "Three Point Polar Alignment" |
    where-object Name -eq "Hocus Focus" |
    Format-List Installer