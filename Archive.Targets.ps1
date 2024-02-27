if (-not (get-module psxisfreader)){import-module psxisfreader}

@(
    #40
    #50
    #90
    #135
    #350
    #950
    #1000
    #4000
    )|ForEach-Object{

        $focalLength=$_

        $topLevelFolder = "P:\Astrophotography\$($focalLength)mm"
        $rejectedFilesToDelete = get-childitem $topLevelFolder -Directory |
            foreach-object{
                $targetFolder = $_.FullName
                $xisfFiles = Get-ChildItem -Path $targetFolder -Recurse -Filter *.xisf -File
                $rejectionFolder = join-path $targetFolder "Rejection"
                if(test-path $rejectionFolder){
                    $rejects = Get-ChildItem -Path $rejectionFolder -Recurse -Filter *.xisf -File
                    $rejectFileNames = $rejects.Name
                    $xisfFiles | 
                        where-object {$_.FullName -notin $rejects.FullName} |
                        group-object Name | 
                        where-object {$rejectFileNames -contains ($_.Name)} |
                        foreach-object {
                            $_.Group
                        } |
                        foreach-object {
                            $originalFile = $_
                            $rejectedFile = join-path $rejectionFolder $originalFile.Name
                            if(test-path $rejectedFile){
                                write-warning "File $($originalFile.FullName) also exists in rejection folder..."
                                $originalFile
                            }
                            else{
                                write-warning "File $($originalFile.FullName) also exists in child of rejection folder."
                            }
                        }
                }
            }
        if($rejectedFilesToDelete){
            write-warning "Deleting $($rejectedFilesToDelete.Count) files."        
            $rejectedFilesToDelete|foreach-object {
                #Remove-Item $_.FullName
            }
        }
    }