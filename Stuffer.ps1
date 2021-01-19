<# ADMIN #>
$admin = "rreyes"

<# STUFFER PATH #>
$stufferDirectory = $MyInvocation.MyCommand.Path | Split-Path

<# TO LOCATION #>
$stufferType = Read-Host "Where is the environment located: 1-AWS or 2-US/EU"

if ($stufferType -eq 1) {
    
    $customerCode = Read-Host "Enter the customer code"
    $option = Read-Host "Select a region: 1-Frankfurt or 2-Sydney"

    . "$($stufferDirectory)\Locations\AWS.ps1" $stufferDirectory $admin $customerCode $option

}
else {

    $option = Read-Host "Select a region: 1-SG or 2-LN"
    $customerName = Read-Host "Enter the customer OU name"
    $customerCode = Read-Host "Enter the customer code"

    . "$($stufferDirectory)\Locations\US-EU.ps1" $stufferDirectory $admin $option $customerName $customerCode
}
