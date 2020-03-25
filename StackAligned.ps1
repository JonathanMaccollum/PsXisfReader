import-module $PSScriptRoot/pixinsightpreprocessing.psm1 -Force
import-module $PSScriptRoot/PsXisfReader.psm1 -Force

$target="E:\Astrophotography\1000mm\NGC2805"
$AlignedOutputPath = "S:\PixInsight\Aligned"

$aligned = Get-XisfFile -Path $AlignedOutputPath |
    where-object object -like "'NGC*2805*" 
write-host "Aligned Stats"
$aligned |
    group-object Filter,Exposure | 
    foreach-object {
        $filter=$_.Values[0]
        $exposure=$_.Values[1]
        new-object psobject -Property @{
            Filter=$filter
            Exposure=$exposure
            Exposures=$_.Group.Count
            ExposureTime=([TimeSpan]::FromSeconds($_.Group.Count*$exposure))
            Images=$_.Group
        } } |
    Sort-Object Filter |
    Format-Table Filter,Exposures,Exposure,ExposureTime,TopWeight

$aligned |
    group-object Filter | 
    foreach-object {
        $filter=$_.Values[0]
        new-object psobject -Property @{
            Filter=$filter
            Images=$_.Group
        } } |
    ForEach-Object {
        $filter = $_.Filter
        $outputFileName = $_.Images[0].Object.Trim("'")
        $outputFileName+=".$($filter.Trim("'"))"            
        $images = $_.Images
        $images | group-object Exposure | foreach-object {
            $exposure=$_.Group[0].Exposure;
            $outputFileName+=".$($_.Group.Count)x$($exposure)s"
        }
        $outputFileName+=".xisf"
        write-host $outputFileName
        $ref = $images | sort-object {[decimal]::Parse($_.SSWeight)} -Descending | select-object -first 1
        write-host ($ref.Path.Name)
        $toStack = $_.Images | sort-object {
            $x = $_
            ($x.Path.Name) -ne ($ref.Path.Name)
        }
        $outputFile = Join-Path $target $outputFileName
        if(-not (test-path $outputFile)) {
            Invoke-PiLightIntegration `
                -Images ($toStack|foreach-object {$_.Path}) `
                -OutputFile $outputFile `
                -KeepOpen `
                -PixInsightSlot 200
        }
    }

