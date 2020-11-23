[CmdletBinding()]
param (
    $SourceFolder = $PSScriptRoot
)

# Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force <= Not required in PS 7?
Set-PSRepository -Name PSGallery -InstallationPolicy Trusted

if (-not (Get-Module PSDepend -ListAvailable)) {
    Install-Module PSDepend -Repository (Get-PSRepository)0].Name -Scope CurrentUser
}

Push-Location $SourceFolder -StackName BuildScript
Invoke-PSDepend -Path $SourceFolder -Confirm:$false
Pop-Location -StackName BuildScript
