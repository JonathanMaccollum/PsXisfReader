Push-Location $PSScriptRoot
try{
    Import-Module ../Modules/PsXisfReader/PsXisfReader.psd1 -force -Verbose
    if(-not $PublishApiKey){
        $PublishApiKey = Read-Host -Prompt "Specify NuGetApiKey" -MaskInput
    }
    if(Test-ModuleManifest -Path "../Modules/PsXisfReader/PsXisfReader.psd1" -Verbose){
        Publish-Module -Path "../Modules/PsXisfReader/" -NuGetApiKey  $PublishApiKey -Verbose
    }    
}
finally{
    Pop-Location
}
Update-Module PsXisfReader