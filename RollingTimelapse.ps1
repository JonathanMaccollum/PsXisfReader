import-module PsXisfReader
$VerbosePreference="CONTINUE"
$subs = Get-ChildItem `
    -Path D:\Backups\Camera\Dropoff\NINA `
    -Filter "20210604.FirefliesAtopTrees.CloudyNight_D1_LIGHT_*.xisf" |
    Sort-Object CreationTime

$pairwiseSetCount=10
$subCount=$subs.Count
$sets = @((-$pairwiseSetCount+1)..($subCount-1))|foreach-object{
    $i=$_
    $set = @($i..($i+$pairwiseSetCount-1)) | where-object {$_ -ge 0} | where-object {$_ -lt $subCount}
    Write-Output -NoEnumerate -InputObject $set
} 

$x=0
$integrationDefinitions = $sets | foreach-object {
    $toIntegrate = $subs[$_]
    if($toIntegrate.Count -ge 3){
        $outputFile = "S:\PixInsight\Timelapse\$x.xisf"
        if(-not (Test-Path $outputFile)){
            $definition = 
                Invoke-PiLightIntegration `
                    -PixInsightSlot 201 `
                    -Rejection "NoRejection" `
                    -Combination "Maximum" `
                    -OutputFile $outputFile `
                    -Images $toIntegrate `
                    -Normalization "NoNormalization" `
                    -GenerateRejectionMaps:$false `
                    -GenerateDrizzleData:$false `
                    -EvaluateNoise:$false `
                    -OutputDefinitionOnly
            new-object psobject -Property @{
                Script = $definition
                OutputFile = $outputFile
            }
        }
    }
    $x+=1
}

Invoke-PIIntegrationBatchScript `
    -KeepOpen `
    -PixInsightSlot 201 `
    -Verbose `
    -ScriptsWithOutputFile $integrationDefinitions