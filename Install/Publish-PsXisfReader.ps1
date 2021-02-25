Push-Location $PSScriptRoot
try{
    Import-Module ../PsXisfReader.psd1 -force -Verbose
    if(-not $PublishApiKey){
        $PublishApiKey = Read-Host -Prompt "Specify NuGetApiKey" -MaskInput
    }
    if(Test-ModuleManifest -Path "../PsXisfReader.psd1" -Verbose){
        Publish-Module -Path "../" -NuGetApiKey  $PublishApiKey -Verbose
    }    
}
finally{
    Pop-Location
}
