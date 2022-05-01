import-module PsXisfReader

Function ConvertTo-CalibrationInfo
{
    param(
        [string[]]$History
    )

    $calibrationHeaders = $History |
        where-object {$_.StartsWith("ImageCalibration.")} |
        foreach-object { $_.Substring(17) } |
        foreach-object {
            $parts = $_.Split(": ")
            $key = $parts[0]
            $value = [string]::Join(": ", $parts[1..($parts.Count-1)])
            new-object psobject -Property @{
                Key=$key
                Value = $value
            }
        }

    $results = new-object psobject -Property @{
    }

    $calibrationHeaders | 
        where-object Key -ne "inputHints" |
        foreach-object {
            Add-Member -MemberType NoteProperty -InputObject $results -Name ($_.Key) -Value ($_.Value)
        }

    $results
}

Get-ChildItem -Path "F:\PixInsightLT\Calibrated" -Directory |
    ForEach-Object {
        $calibrationLogFile = join-path $_.FullName "CalibrationLog.json"
        if(-not (test-path $calibrationLogFile)){
            $calibrationLog = 
                Get-XisfLightFrames -Path ($_) -UseCache -UseErrorCache -SkipOnError -Recurse |
                    foreach-object {
                        $calibrated = $_
                        $nameWithoutExtension = [System.IO.Path]::GetFileNameWithoutExtension($calibrated.Path.FullName)
                        $extension = $calibrated.Path.Extension
                        $parentFileName=$nameWithoutExtension
                        do{
                            $parentFileName = $parentFileName.Substring(0,$parentFileName.Length-2)
                        }while($parentFileName.EndsWith("_c"))        
                        $parentFileName+=$extension
                        $calibratedOn = $calibrated.Path.LastWriteTime
                        $CalibrationInfo = ConvertTo-CalibrationInfo -History $calibrated.History
                
                        add-member -InputObject $calibrated -MemberType NoteProperty -Name "CalibrationInfo" -Value $CalibrationInfo 
                        add-member -InputObject $calibrated -MemberType NoteProperty -Name "FileName" -Value $parentFileName
                        add-member -InputObject $calibrated -MemberType NoteProperty -Name "CalibratedOn" -Value $calibratedOn 
                
                        $calibrated.Path = $null
                        $calibrated.History = $null
                        $calibrated
                    }
            $calibrationLog | ConvertTo-Json -Depth 5 | Out-File $calibrationLogFile -Force -Verbose
        }        
    }
