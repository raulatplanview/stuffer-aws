<# STUFFER PATH #>
$stufferDirectory = $MyInvocation.MyCommand.Path | Split-Path

<# PRIMARY CUSTOMER INFORMATION #>
$type = Read-Host "Select a build type: 1-InPlace 2-New Build/Upgrade"
$customerCode = Read-Host "Enter the customer code"
$option = Read-Host "Select a region: 1-Frankfurt or 2-Sydney"

. "$($stufferDirectory)\Locations\AWS.ps1" 
