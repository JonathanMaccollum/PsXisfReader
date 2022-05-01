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
[System.IO.Directory]::CreateDirectory("F:\PixInsightLT\TimelapseEvery$pairwiseSetCount\")
$x=0
$integrationDefinitions = $sets | foreach-object {
    $toIntegrate = $subs[$_]
    if($toIntegrate.Count -ge 3){
        $outputFile = "F:\PixInsightLT\TimelapseEvery$pairwiseSetCount\$x.xisf"
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

($sets | where-object Count -ge 3) | measure-object 

@(2..2838)|where-object {
    -not (Test-Path "S:\PixInsight\Corrected\$($_)_cc.xisf")
}


@(2..2838)|where-object {
    -not (Test-Path "F:\PixInsightLT\TimelapseEvery$pairwiseSetCount\$($_).xisf")
}


@(2..2838)|where-object {
    -not (Test-Path "S:\PixInsight\Debayered\$($_)_cc_d.xisf")
}
<#
$toStack = @(2..2838) |
foreach-object {
    if($_%$pairwiseSetCount -eq 0){
        Get-Item "S:\PixInsight\Corrected\$($_)_cc.xisf"
    }
}
Invoke-PiLightIntegration `
    -PixInsightSlot 201 `
    -Rejection "NoRejection" `
    -Combination "Maximum" `
    -Images $toStack `
    -Normalization "NoNormalization" `
    -GenerateRejectionMaps:$false `
    -GenerateDrizzleData:$false `
    -EvaluateNoise:$false `
    -OutputFile "E:\Astrophotography\40mm\20210604.FirefliesAtopTrees.CloudyNight\$($toStack.Count)xMaxOf8x10s.cc.integrated.xisf" `
    -KeepOpen
#>





$subs = Get-ChildItem `
    -Path S:\PixInsight\Corrected\FirefliesAtopTrees `
    -Filter "*.xisf" |
    Sort-Object {[int]::Parse($_.Name.Split("_")[0])}

$pairwiseSetCount=3
$subCount=$subs.Count
$sets = @((-$pairwiseSetCount+1)..($subCount-1))|foreach-object{
    $i=$_
    $set = @($i..($i+$pairwiseSetCount-1)) | where-object {$_ -ge 0} | where-object {$_ -lt $subCount}
    Write-Output -NoEnumerate -InputObject $set
} 
[System.IO.Directory]::CreateDirectory("F:\PixInsightLT\TimelapseEvery$pairwiseSetCount\")
$x=0
$integrationDefinitions = $sets | foreach-object {
    $toIntegrate = $subs[$_]
    if($toIntegrate.Count -ge 3){
        $outputFile = "F:\PixInsightLT\TimelapseEvery$pairwiseSetCount\$x.xisf"
        if(-not (Test-Path $outputFile)){
            $definition = 
                Invoke-PiLightIntegration `
                    -PixInsightSlot 201 `
                    -Rejection "NoRejection" `
                    -Combination "Average" `
                    -OutputFile $outputFile `
                    -Images $toIntegrate `
                    -Normalization "NoNormalization" `
                    -GenerateRejectionMaps:$false `
                    -GenerateDrizzleData:$false `
                    -EvaluateNoise:$false `
                    -OutputDefinitionOnly `
                    -WeightMode "DontCare"
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
