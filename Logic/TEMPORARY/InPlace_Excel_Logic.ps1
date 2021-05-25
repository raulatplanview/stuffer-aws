#Import-Module \\scripthost\modules\pvadmin
Import-Module SQLSERVER
Import-Module F5-LTM

<# COPYING EXCEL TEMPLATE #>
Get-ChildItem -Path "C:\Users\$($awsAdmin)\Planview, Inc\E1 Build Cutover - Documents\Customer Builds\1_FolderTemplate\18" -Filter "InPlace*" | Copy-Item -Destination "C:\Users\$($awsAdmin)\Desktop"
$excelFilePath = Get-ChildItem -Path "C:\Users\$($awsAdmin)\Desktop\" -Filter "InPlace*" | ForEach-Object {$_.FullName}

<# EXCEL OBJECT #>
$excel = New-Object -ComObject Excel.Application
$excelfile = $excel.Workbooks.Open($excelFilePath)
$buildData = $excelfile.sheets.item("MasterConfig")

<# COMMON FIELDS #>
# AWS BUILD #
$buildData.Cells.Item(18,2)= "False" 

# SPLIT TIER #
$buildData.Cells.Item(19,2)= "False"

# PRODUCTION SERVER COUNT #
$buildData.Cells.Item(23,2)= ($productionComputers.Count - 1)

# SANDBOX SERVER COUNT #
$buildData.Cells.Item(23,3)= ($sandboxComputers.Count - 1) 

# CUSTOMER CODE #
$buildData.Cells.Item(10,2)= $customerCode.ToUpper()

# DATACENTER LOCATION #
$buildData.Cells.Item(9,2)= $dataCenterLocation

# AD OU NAME (Pulled from Production Environment) #
$AD_OU = $environmentsMaster[0][1][0][2].Value
$buildData.Cells.Item(14,2)= $AD_OU

# SAASINFO LINK #
$buildData.Cells.Item(3,2)= "N/A"

<# MAIN LOOP #>
for ($x=0; $x -lt $environmentsMaster.Length; $x++) {
    
    if ($environmentsMaster[$x][0] -eq $slot1) { 
        Write-Host ":::::::: $($environmentsMaster[$x][0]) Environment ::::::::" -Foregroundcolor Yellow
        
        $webServerCount = 0
        for ($y=1; $y -lt $environmentsMaster[$x].Length; $y++) {     
            
            if ($environmentsMaster[$x][$y][0][6].Value.Substring(($environmentsMaster[$x][$y][0][6].Value.Length - 5), 3) -eq "app" -or $environmentsMaster[$x][$y][0][6].Value.Substring(($environmentsMaster[$x][$y][0][6].Value.Length - 5), 3) -eq "ctm") {
            
                ##########################
                # PRODUCTION APP SERVER 
                ##########################
                if ($environmentsMaster[$x][$y][0][6].Value.Substring(($environmentsMaster[$x][$y][0][6].Value.Length - 5), 3) -eq "app") {
                    Write-Host "THIS IS THE PRODUCTION APP SERVER" -ForegroundColor Cyan

                    <# CPU/RAM #>
                    Write-Host "Server CPU and RAM" -ForegroundColor Red
                    Write-Host "Server Name: $($environmentsMaster[$x][$y][0][6].Value)"
                    Write-Host "Server CPUs: $(@($environmentsMaster[$x][$y][2][2]).Count)"
                    Write-Host "Server RAM: $([MATH]::Round((($environmentsMaster[$x][$y][2][1].MaxCapacity) / 1000000),2))"

                    <# HARDDRIVES #> 
                    Write-Host "Disks and Disk Capacity" -ForegroundColor Red
                    $diskResize = "Yes"
                    $hdStringArray = ""
                    foreach ($hd in $environmentsMaster[$x][$y][2][0]) {
                        $hdString = "$($hd.DeviceID) $([MATH]::Round($hd.Size / 1GB,2))gb"
                        $hdStringArray += "$($hdString)`n"
                        Write-Host $hdString  
                        if (([MATH]::Round($hd.Size / 1GB,2)) -gt 60) {
                            $diskResize = "No"  
                        }
                    }
                    Write-Host "Standard Size Disks (less than 60GB): $($diskResize)"

                    <# AWS METADATA #>

                        # INSTANCE SIZE #
                        Write-Host "Instance Size" -ForegroundColor Red
                        $instanceSize = $environmentsMaster[$x][$y][1][2]
                        Write-Host $instanceSize
                        
                        # AVAILABILITY ZONE #
                        Write-Host "Availability Zone" -ForegroundColor Red
                        $availabilityZone = $environmentsMaster[$x][$y][1][3]
                        Write-Host $availabilityZone

                        # IP ADDRESS #
                        Write-Host "IP Address" -ForegroundColor Red
                        $ipAddress = $environmentsMaster[$x][$y][1][4]
                        Write-Host $ipAddress

                        # INSTANCE ID #
                        Write-Host "Instance ID" -ForegroundColor Red
                        $instanceID = $environmentsMaster[$x][$y][1][0]
                        Write-Host $instanceID
                    
                    <# SCHEDULED TASKS #>
                    Write-Host "Scheduled Tasks on Server" -ForegroundColor Red
                    $tasks = Invoke-Command -ComputerName "$($environmentsMaster[$x][$y][0][6].Value)$($domain)" -ScriptBlock {
                        Get-ScheduledTask -TaskPath "\" | Select-Object -Property TaskName, LastRunTime | Where-Object TaskName -notlike "Op*" 
                    } -Credential $credentials

                    $task_array = ""
                    foreach ($task in $tasks){
                        Write-Host "Task Name: $($task.TaskName)"
                        $task_array += "$($task.TaskName)`n"
                    }
                    
                    <# OPEN SUITE #>
                    Write-Host "OpenSuite" -ForegroundColor Red
                    $opensuite = Invoke-Command -ComputerName "$($environmentsMaster[$x][$y][0][6].Value)$($domain)" -Credential $credentials -ScriptBlock {
                        if ((Test-Path -Path "C:\ProgramData\Actian" -PathType Container) -And (Test-Path -Path "F:\Planview\Interfaces\OpenSuite" -PathType Container)) {

                            $software = "*Actian*";
                            $installed = (Get-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*" | Where-Object { $_.DisplayName -like $software }) -ne $null

                            if ($installed) {
                                return "Yes"
                            }
                            
                        } else {
                            return "No"
                        }
                    }
                    Write-Host "OpenSuite Detected: $($opensuite)"

                    <# CONNECTORS #>
                    Write-Host "Connectors" -ForegroundColor Red
                    $PRMAdapter = "False"
                    $PPAdapter = "False"
                    $LKAdapter = "False"

                    $PRMini = Invoke-Command -ComputerName "$($environmentsMaster[$x][$y][0][6].Value)$($domain)" -Credential $credentials -ScriptBlock {

                        if (Test-Path -Path "F:\Planview\midtier\webserver\objects\PRM_Adapter_Config.ini" -PathType leaf){

                            $PRMiniContent = (Get-Content -Path "F:\Planview\midtier\webserver\objects\PRM_Adapter_Config.ini" | Where-Object {$_.length -ne 0}) | Where-Object {$_.substring(0,1) -ne ";"}

                            $masterArray = @()
                            $databaseName = ""
                            foreach($x in $PRMiniContent){

                                $valuePair = New-Object PSObject
                                Add-Member -InputObject $valuePair -MemberType NoteProperty -Name Database -Value ""
                                Add-Member -InputObject $valuePair -MemberType NoteProperty -Name Key -Value ""

                                if ($x -match "^\["){

                                    $databaseName = $x.substring(1, ($x.length - 2)) 

                                    $valuePair.Database = $databaseName
                                    $valuePair.Key = $databaseName
                                    $masterArray += $valuePair
                                
                                
                                } else {       
                                
                                    $valuePair.Database = $databaseName
                                    $valuePair.Key = $x 
                                    $masterArray += $valuePair
                                
                                }           
                            }
                        
                            return $masterArray
                        
                        } else {

                            return 0

                        }
                    }

                    if ($PRMini -ne '0'){

                        $PRMAdapter = "True"

                        Write-Host "PRM Adapter Found"
                        $LKKey = ($PRMini | Where-Object {$_.Database -like "*PROD"} | Where-Object {$_.Key -like "use_prm_leankit*"}).Key
                        $PPKey = ($PRMini | Where-Object {$_.Database -like "*PROD"} | Where-Object {$_.Key -like "use_prm_projectplace*"}).Key

                        if ($LKKey -like "*true*"){
                            $LKAdapter = "True"
                            Write-Host "LeanKit Connector --- True"
                        } else {
                            Write-Host "LeanKit Connector --- False"
                        }

                        if ($PPKey -like "*true*"){
                            $PPAdapter = "True"
                            Write-Host "ProjectPlace Connector --- True"
                        } else {
                            Write-Host "ProjectPlace Connector --- False"
                        }

                    } else {

                        Write-Host "PRM Adapter not found"

                        $LegacyPPAdapter = Invoke-Command -ComputerName "$($environmentsMaster[$x][$y][0][6].Value)$($domain)" -Credential $credentials -ScriptBlock {
                            Test-Path -Path "F:\Planview\midtier\webserver\objects\ProjectPlace_Config.ini" -PathType leaf
                        }

                        $PPAdapter = $LegacyPPAdapter
                        Write-Host "ProjectPlace (Legacy install) --- $($PPAdapter)"

                    }
                    

                    <# EXCEL LOGIC AND VARIABLES#>
                    $buildData.Cells.Item(52,2)= "$($environmentsMaster[$x][$y][0][6].Value)"
                    $buildData.Cells.Item(52,3)= "$(@($environmentsMaster[$x][$y][2][2]).Count)"
                    $buildData.Cells.Item(52,4)= "$([MATH]::Round((($environmentsMaster[$x][$y][2][1].MaxCapacity) / 1000000),2))"
                    $buildData.Cells.Item(52,5)= $hdStringArray
                    $buildData.Cells.Item(52,6)= $diskResize
                    $buildData.Cells.Item(52,7)= $task_array

                    <# AWS-SPECIFIC VARIABLES #>
                    $buildData.Cells.Item(52,8)= $instanceSize
                    $buildData.Cells.Item(52,9)= $availabilityZone
                    $buildData.Cells.Item(52,10)= $ipAddress
                    $buildData.Cells.Item(52,11)= $instanceID

                    $buildData.Cells.Item(31,2)= $PPAdapter
                    $buildData.Cells.Item(32,2)= $LKAdapter
                    $buildData.Cells.Item(37,2)= $PRMAdapter

                    $buildData.Cells.Item(36,2)= $opensuite

                    Write-Host "`n" -ForegroundColor Red
                } 

                ##################################
                # PRODUCTION CTM SERVER (Troux) 
                ##################################
                elseif ($environmentsMaster[$x][$y][0][6].Value.Substring(($environmentsMaster[$x][$y][0][6].Value.Length - 5), 3) -eq "ctm") {
                    Write-Host "THIS IS THE PRODUCTION TROUX SERVER" -ForegroundColor Cyan

                    <# CPU/RAM #>
                    Write-Host "Server CPU and RAM" -ForegroundColor Red
                    Write-Host "Server Name: $($environmentsMaster[$x][$y][0][6].Value)"
                    Write-Host "Server CPUs: $(@($environmentsMaster[$x][$y][2][2]).Count)"
                    Write-Host "Server RAM: $([MATH]::Round((($environmentsMaster[$x][$y][2][1].MaxCapacity) / 1000000),2))"     

                    <# HARDDRIVES #>
                    Write-Host "Disks and Disk Capacity" -ForegroundColor Red
                    $diskResize = "Yes"
                    $hdStringArray = ""
                    foreach ($hd in $environmentsMaster[$x][$y][2][0]) {
                        $hdString = "$($hd.DeviceID) $([MATH]::Round($hd.Size / 1GB,2))gb"
                        $hdStringArray += "$($hdString)`n"
                        Write-Host $hdString  
                        if (([MATH]::Round($hd.Size / 1GB,2)) -gt 60) {
                            $diskResize = "No"  
                        }
                    }
                    Write-Host "Standard Size Disks (less than 60GB): $($diskResize)"       

                    <# AWS METADATA #>

                        # INSTANCE SIZE #
                        Write-Host "Instance Size" -ForegroundColor Red
                        $instanceSize = $environmentsMaster[$x][$y][1][2]
                        Write-Host $instanceSize
                        
                        # AVAILABILITY ZONE #
                        Write-Host "Availability Zone" -ForegroundColor Red
                        $availabilityZone = $environmentsMaster[$x][$y][1][3]
                        Write-Host $availabilityZone

                        # IP ADDRESS #
                        Write-Host "IP Address" -ForegroundColor Red
                        $ipAddress = $environmentsMaster[$x][$y][1][4]
                        Write-Host $ipAddress

                        # INSTANCE ID #
                        Write-Host "Instance ID" -ForegroundColor Red
                        $instanceID = $environmentsMaster[$x][$y][1][0]
                        Write-Host $instanceID

                    <# SCHEDULED TASKS #>
                    Write-Host "Scheduled Tasks on Server" -ForegroundColor Red
                    $tasks = Invoke-Command -ComputerName "$($environmentsMaster[$x][$y][0][6].Value)$($domain)" -ScriptBlock {
                        Get-ScheduledTask -TaskPath "\" | Select-Object -Property TaskName, LastRunTime | Where-Object TaskName -notlike "Op*" 
                    } -Credential $credentials

                    $task_array = ""
                    foreach ($task in $tasks){
                        Write-Host "Task Name: $($task.TaskName)"
                        $task_array += "$($task.TaskName)`n"
                    }

                    <# EXCEL LOGIC AND VARIABLES#>
                    $buildData.Cells.Item(53,2)= "$($environmentsMaster[$x][$y][0][6].Value)"
                    $buildData.Cells.Item(53,3)= "$(@($environmentsMaster[$x][$y][2][2]).Count)"
                    $buildData.Cells.Item(53,4)= "$([MATH]::Round((($environmentsMaster[$x][$y][2][1].MaxCapacity) / 1000000),2))"
                    $buildData.Cells.Item(53,5)= $hdStringArray
                    $buildData.Cells.Item(53,6)= $diskResize
                    $buildData.Cells.Item(53,7)= $task_array

                    <# AWS-SPECIFIC VARIABLES #>
                    $buildData.Cells.Item(53,8)= $instanceSize
                    $buildData.Cells.Item(53,9)= $availabilityZone
                    $buildData.Cells.Item(53,10)= $ipAddress
                    $buildData.Cells.Item(53,11)= $instanceID

                    Write-Host "`n" -ForegroundColor Red  
                }
        }
        
            ##########################
            # PRODUCTION WEB SERVER 
            ##########################
            elseif ($environmentsMaster[$x][$y][0][6].Value.Substring(($environmentsMaster[$x][$y][0][6].Value.Length - 5), 3) -eq "web") {

                Write-Host "THIS IS THE PRODUCTION WEB SERVER" -ForegroundColor Cyan

                <# CPU/RAM #>
                Write-Host "Server CPU and RAM" -ForegroundColor Red
                Write-Host "Server Name: $($environmentsMaster[$x][$y][0][6].Value)"
                Write-Host "Server CPUs: $(@($environmentsMaster[$x][$y][2][2]).Count)"
                Write-Host "Server RAM: $([MATH]::Round((($environmentsMaster[$x][$y][2][1].MaxCapacity) / 1000000),2))"

                <# HARDDRIVES #>
                Write-Host "Disks and Disk Capacity" -ForegroundColor Red
                $diskResize = "Yes"
                $hdStringArray = ""
                foreach ($hd in $environmentsMaster[$x][$y][2][0]) {
                    $hdString = "$($hd.DeviceID) $([MATH]::Round($hd.Size / 1GB,2))gb"
                    $hdStringArray += "$($hdString)`n"
                    Write-Host $hdString  
                    if (([MATH]::Round($hd.Size / 1GB,2)) -gt 60) {
                        $diskResize = "No"  
                    }
                }
                Write-Host "Standard Size Disks (less than 60GB): $($diskResize)"

                <# AWS METADATA #>

                        # INSTANCE SIZE #
                        Write-Host "Instance Size" -ForegroundColor Red
                        $instanceSize = $environmentsMaster[$x][$y][1][2]
                        Write-Host $instanceSize
                        
                        # AVAILABILITY ZONE #
                        Write-Host "Availability Zone" -ForegroundColor Red
                        $availabilityZone = $environmentsMaster[$x][$y][1][3]
                        Write-Host $availabilityZone

                        # IP ADDRESS #
                        Write-Host "IP Address" -ForegroundColor Red
                        $ipAddress = $environmentsMaster[$x][$y][1][4]
                        Write-Host $ipAddress

                        # INSTANCE ID #
                        Write-Host "Instance ID" -ForegroundColor Red
                        $instanceID = $environmentsMaster[$x][$y][1][0]
                        Write-Host $instanceID

                <# SCHEDULED TASKS #>
                Write-Host "Scheduled Tasks on Server" -ForegroundColor Red
                $tasks = Invoke-Command -ComputerName "$($environmentsMaster[$x][$y][0][6].Value)$($domain)" -ScriptBlock {
                    Get-ScheduledTask -TaskPath "\" | Select-Object -Property TaskName, LastRunTime | Where-Object TaskName -notlike "Op*" 
                } -Credential $credentials

                $task_array = ""
                foreach ($task in $tasks){
                    Write-Host "Task Name: $($task.TaskName)"
                    $task_array += "$($task.TaskName)`n"
                }

                <# CURRENT VERSION #>
                Write-Host "Current Environment Version" -ForegroundColor Red
                $crVersion = Invoke-Command -ComputerName "$($environmentsMaster[$x][$y][0][6].Value)$($domain)" -Credential $credentials -ScriptBlock {
                    Get-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\Planview\WebServerPlatform"
                }
                Write-Host $crVersion.CrVersion

                <# MAJOR VERSION #>
                Write-Host "Major Version" -ForegroundColor Red
                $majorVersion = $crVersion.CrVersion.Split('.')[0]
                $majorVersion

                # NEW RELIC #
                Write-Host "New Relic" -ForegroundColor Red
                $newRelic = Invoke-Command -ComputerName "$($environmentsMaster[$x][$y][0][6].Value)$($domain)" -Credential $credentials -ScriptBlock {
                    if (Test-Path -Path "C:\ProgramData\New Relic" -PathType Container ) {
                        Write-Host "New Relic has been detected on this server"
                        return "Yes"
                    } else {
                        Write-Host "New Relic was not detected on this server"
                        return "No"
                    }
                }

                    # GET WEB CONFIG #
                    $webConfig = Invoke-Command -ComputerName "$($environmentsMaster[$x][$y][0][6].Value)$($domain)" -Credential $credentials -ScriptBlock {
                        return Get-Content -Path "F:\Planview\MidTier\ODataService\Web.config"
                    }
                    $webConfig = [xml] $webConfig

                <# PRODUCTION URL #>
                Write-Host "Production URL" -ForegroundColor Red
                $environmentURL = $webConfig.configuration.appSettings.add | Where-Object {$_.key -eq "PveUrl"} | Select-Object -Property value
                Write-Host $environmentURL.value

                <# DNS ALIAS #>
                Write-Host "Production DNS Alias" -ForegroundColor Red            
                $dnsAlias = ($environmentURL.value.Split('//')[2]).Split('.')[0] 
                $dnsAlias

                <# REPORT FARM URL #>
                Write-Host "Report Farm URL" -ForegroundColor Red
                $reportfarmURL = $webConfig.configuration.appSettings.add | Where-Object {$_.key -eq "Report_Server_Web_Service_URL"} | Select-Object -Property value
                Write-Host $reportfarmURL.value

                <# ENCRYPTED PVMASTER PASSWORD #>
                Write-Host "Encrypted PVMaster Password" -ForegroundColor Red
                $encryptedPVMasterPassword = $webConfig.configuration.appSettings.add | Where-Object {$_.key -eq "PveUserPassword"} | Select-Object -Property value
                Write-Host $encryptedPVMasterPassword.value
                
                <# UNENCRYPTED PVMASTER PASSWORD #>
                Write-Host "Unencrypted PVMaster Password" -ForegroundColor Red
                $unencryptedPVMasterPassword = Invoke-PassUtil -InputString $encryptedPVMasterPassword.value -Deobfuscation
                Write-Host $unencryptedPVMasterPassword

                <# IP RESTRICTIONS #>
                Write-Host "IP Restrictions on F5" -ForegroundColor Red
                Write-Host "Currently Not Automated"
                $IPRestrictions = "Not Automated"
<#                    
                    # Authentication on the F5 #
                    $websession =  New-Object Microsoft.PowerShell.Commands.WebRequestSession
                    $jsonbody = @{username = $f5Credentials.UserName ; password = $f5Credentials.GetNetworkCredential().Password; loginProviderName='tmos'} | ConvertTo-Json
                    $authResponse = Invoke-RestMethodOverride -Method Post -Uri "https://$($f5ip)/mgmt/shared/authn/login" -Credential $f5Credentials -Body $jsonbody -ContentType 'application/json'
                    $token = $authResponse.token.token
                    $websession.Headers.Add('X-F5-Auth-Token', $Token)

                    # Calling data-group REST endpoint and parsing IPRestrictions list #
                    $IPRestrictionsList = (Invoke-RestMethod  -Uri "https://$($f5ip)/mgmt/tm/ltm/data-group/internal" -WebSession $websession).Items | 
                        Where-Object {$_.name -eq "IPRestrictions"} | Select-Object -Property records

                    foreach ($record in $IPRestrictionsList.records) {
                        if ($record.name -eq "$($dnsAlias).pvcloud.com") {
                            $IPRestrictions = "Yes"
                            Write-Host "IP restrctions were found for $($dnsAlias).pvcloud.com"
                        }
                    }

                    if ($IPRestrictions -eq "No") {
                        Write-Host "No IP restrictions found for $($dnsAlias).pvcloud.com"
                    }
#>
                <# EXCEL LOGIC AND VARIABLES#>
                $buildData.Cells.Item(25,2)= $crVersion.CrVersion
                $buildData.Cells.Item(35,2)= $IPRestrictions
                $buildData.Cells.Item(13,2)= $majorVersion
                $buildData.Cells.Item(17,2)= $encryptedPVMasterPassword.value
                $buildData.Cells.Item(16,2)= $unencryptedPVMasterPassword
                $buildData.Cells.Item(1,2)= $environmentURL.value
                $buildData.Cells.Item(7,2)= $dnsAlias
                $buildData.Cells.Item(46,2)= $reportfarmURL.value
                $buildData.Cells.Item(15,2)= $newRelic

                if ($webServerCount -gt 0){
                    $buildData.Cells.Item(58 + ($webServerCount - 1),2)= "$($environmentsMaster[$x][$y][0][6].Value)"
                    $buildData.Cells.Item(58 + ($webServerCount - 1),3)= "$(@($environmentsMaster[$x][$y][2][2]).Count)"
                    $buildData.Cells.Item(58 + ($webServerCount - 1),4)= "$([MATH]::Round((($environmentsMaster[$x][$y][2][1].MaxCapacity) / 1000000),2))"
                    $buildData.Cells.Item(58 + ($webServerCount - 1),5)= $hdStringArray
                    $buildData.Cells.Item(58 + ($webServerCount - 1),6)= $diskResize
                    $buildData.Cells.Item(58 + ($webServerCount - 1),7)= $task_array

                    <# AWS-SPECIFIC VARIABLES #>
                    $buildData.Cells.Item(58 + ($webServerCount - 1),8)= $instanceSize
                    $buildData.Cells.Item(58 + ($webServerCount - 1),9)= $availabilityZone
                    $buildData.Cells.Item(58 + ($webServerCount - 1),10)= $ipAddress
                    $buildData.Cells.Item(58 + ($webServerCount - 1),11)= $instanceID
                }
                else {
                    $buildData.Cells.Item(51,2)= "$($environmentsMaster[$x][$y][0][6].Value)"
                    $buildData.Cells.Item(51,3)= "$(@($environmentsMaster[$x][$y][2][2]).Count)"
                    $buildData.Cells.Item(51,4)= "$([MATH]::Round((($environmentsMaster[$x][$y][2][1].MaxCapacity) / 1000000),2))"
                    $buildData.Cells.Item(51,5)= $hdStringArray
                    $buildData.Cells.Item(51,6)= $diskResize
                    $buildData.Cells.Item(51,7)= $task_array

                    <# AWS-SPECIFIC VARIABLES #>
                    $buildData.Cells.Item(51,8)= $instanceSize
                    $buildData.Cells.Item(51,9)= $availabilityZone
                    $buildData.Cells.Item(51,10)= $ipAddress
                    $buildData.Cells.Item(51,11)= $instanceID
                }
                
                $webServerCount++
                $buildData.Cells.Item(24,2)= $webServerCount

                
                Write-Host "`n" -ForegroundColor Red  

        }

            ##########################
            # PRODUCTION SAS SERVER 
            ##########################
            elseif ($environmentsMaster[$x][$y][0][6].Value.Substring(($environmentsMaster[$x][$y][0][6].Value.Length - 5), 3) -eq "sas") {
                Write-Host "THIS IS THE PRODUCTION SAS SERVER" -ForegroundColor Cyan

                <# CPU/RAM #>
                Write-Host "Server CPU and RAM" -ForegroundColor Red
                Write-Host "Server Name: $($environmentsMaster[$x][$y][0][6].Value)"
                Write-Host "Server CPUs: $(@($environmentsMaster[$x][$y][2][2]).Count)"
                Write-Host "Server RAM: $([MATH]::Round((($environmentsMaster[$x][$y][2][1].MaxCapacity) / 1000000),2))"

                <# HARDDRIVES #>
                Write-Host "Disks and Disk Capacity" -ForegroundColor Red
                $diskResize = "Yes"
                $hdStringArray = ""
                foreach ($hd in $environmentsMaster[$x][$y][2][0]) {
                    $hdString = "$($hd.DeviceID) $([MATH]::Round($hd.Size / 1GB,2))gb"
                    $hdStringArray += "$($hdString)`n"
                    Write-Host $hdString  
                    if (([MATH]::Round($hd.Size / 1GB,2)) -gt 60) {
                        $diskResize = "No"  
                    }
                }
                Write-Host "Standard Size Disks (less than 60GB): $($diskResize)"

                <# AWS METADATA #>

                        # INSTANCE SIZE #
                        Write-Host "Instance Size" -ForegroundColor Red
                        $instanceSize = $environmentsMaster[$x][$y][1][2]
                        Write-Host $instanceSize
                        
                        # AVAILABILITY ZONE #
                        Write-Host "Availability Zone" -ForegroundColor Red
                        $availabilityZone = $environmentsMaster[$x][$y][1][3]
                        Write-Host $availabilityZone

                        # IP ADDRESS #
                        Write-Host "IP Address" -ForegroundColor Red
                        $ipAddress = $environmentsMaster[$x][$y][1][4]
                        Write-Host $ipAddress

                        # INSTANCE ID #
                        Write-Host "Instance ID" -ForegroundColor Red
                        $instanceID = $environmentsMaster[$x][$y][1][0]
                        Write-Host $instanceID

                <# SCHEDULED TASKS #>
                Write-Host "Scheduled Tasks on Server" -ForegroundColor Red
                $tasks = Invoke-Command -ComputerName "$($environmentsMaster[$x][$y][0][6].Value)$($domain)" -ScriptBlock {
                    Get-ScheduledTask -TaskPath "\" | Select-Object -Property TaskName, LastRunTime | Where-Object TaskName -notlike "Op*" 
                } -Credential $credentials

                $task_array = ""
                foreach ($task in $tasks){
                    Write-Host "Task Name: $($task.TaskName)"
                    $task_array += "$($task.TaskName)`n"
                }
                
                <# EXCEL LOGIC AND VARIABLES#>
                $buildData.Cells.Item(55,2)= "$($environmentsMaster[$x][$y][0][6].Value)"
                $buildData.Cells.Item(55,3)= "$(@($environmentsMaster[$x][$y][2][2]).Count)"
                $buildData.Cells.Item(55,4)= "$([MATH]::Round((($environmentsMaster[$x][$y][2][1].MaxCapacity) / 1000000),2))"
                $buildData.Cells.Item(55,5)= $hdStringArray
                $buildData.Cells.Item(55,6)= $diskResize
                $buildData.Cells.Item(55,7)= $task_array

                <# AWS-SPECIFIC VARIABLES #>
                $buildData.Cells.Item(55,8)= $instanceSize
                $buildData.Cells.Item(55,9)= $availabilityZone
                $buildData.Cells.Item(55,10)= $ipAddress
                $buildData.Cells.Item(55,11)= $instanceID

                Write-Host "`n" -ForegroundColor Red  
        }

            ##########################
            # PRODUCTION SQL SERVER 
            ##########################
            elseif ($environmentsMaster[$x][$y][0][6].Value.Substring(($environmentsMaster[$x][$y][0][6].Value.Length - 5), 3) -eq "sql") {
                Write-Host "THIS IS THE PRODUCTION SQL SERVER" -ForegroundColor Cyan

                <# CPU/RAM #>
                Write-Host "Server CPU and RAM" -ForegroundColor Red
                Write-Host "Server Name: $($environmentsMaster[$x][$y][0][6].Value)"
                Write-Host "Server CPUs: $(@($environmentsMaster[$x][$y][2][2]).Count)"
                Write-Host "Server RAM: $([MATH]::Round((($environmentsMaster[$x][$y][2][1].MaxCapacity) / 1000000),2))"

                <# HARDDRIVES #>
                Write-Host "Disks and Disk Capacity" -ForegroundColor Red
                $diskResize = "Yes"
                $hdStringArray = ""
                foreach ($hd in $environmentsMaster[$x][$y][2][0]) {
                    $hdString = "$($hd.DeviceID) $([MATH]::Round($hd.Size / 1GB,2))gb"
                    $hdStringArray += "$($hdString)`n"
                    Write-Host $hdString  
                    if (([MATH]::Round($hd.Size / 1GB,2)) -gt 60) {
                        $diskResize = "No"  
                    }
                }
                Write-Host "Standard Size Disks (less than 60GB): $($diskResize)"

                <# AWS METADATA #>

                        # INSTANCE SIZE #
                        Write-Host "Instance Size" -ForegroundColor Red
                        $instanceSize = $environmentsMaster[$x][$y][1][2]
                        Write-Host $instanceSize
                        
                        # AVAILABILITY ZONE #
                        Write-Host "Availability Zone" -ForegroundColor Red
                        $availabilityZone = $environmentsMaster[$x][$y][1][3]
                        Write-Host $availabilityZone

                        # IP ADDRESS #
                        Write-Host "IP Address" -ForegroundColor Red
                        $ipAddress = $environmentsMaster[$x][$y][1][4]
                        Write-Host $ipAddress

                        # INSTANCE ID #
                        Write-Host "Instance ID" -ForegroundColor Red
                        $instanceID = $environmentsMaster[$x][$y][1][0]
                        Write-Host $instanceID

                <# SCHEDULED TASKS #>
                Write-Host "Scheduled Tasks on Server" -ForegroundColor Red
                $tasks = Invoke-Command -ComputerName "$($environmentsMaster[$x][$y][0][6].Value)$($domain)" -ScriptBlock {
                    Get-ScheduledTask -TaskPath "\" | Select-Object -Property TaskName, LastRunTime | Where-Object TaskName -notlike "Op*" 
                } -Credential $credentials

                $task_array = ""
                foreach ($task in $tasks){
                    Write-Host "Task Name: $($task.TaskName)"
                    $task_array += "$($task.TaskName)`n"
                }

                <# MAINTENANCE DAY #>
                Write-Host "Maintenance Day" -ForegroundColor Red
                $maintenanceDay = $environmentsMaster[$x][$y][0][4].Value
                Write-Host $maintenanceDay

                <# DATABASE PROPERTIES #>
                Write-Host "Database Properties" -ForegroundColor Red
                $sqlSession = New-PSSession -ComputerName "$($environmentsMaster[$x][$y][0][6].Value)$($domain)" -Credential $credentials

                    # ALL DATABASES (NAMES AND SIZES in MB)
                    $mainDatabase = ""
                    Write-Host "Listing All Databases and Sizes (in MB)" -ForegroundColor Cyan
                    $all_databases = Invoke-Command -Session $sqlSession -ScriptBlock { 
                        param ($server, $mainDatabase)        
                        Invoke-Sqlcmd -Query "SELECT d.name,
                        ROUND(SUM(mf.size) * 8 / 1024, 0) Size_MB
                        FROM sys.master_files mf
                        INNER JOIN sys.databases d ON d.database_id = mf.database_id
                        WHERE d.database_id > 4 -- Skip system databases
                        GROUP BY d.name
                        ORDER BY d.name" -ServerInstance $server.Name 
                    } -ArgumentList $environmentsMaster[$x][$y][0][6].Value, $mainDatabase
                    foreach ($database in $all_databases) {
                        Write-Host "$($database.name) ---- $($database.Size_MB) MB"
                        if (($database.name -like "*PROD") -or ($database.name -like "*DEV*")) {
                            if ($database.name -notlike "DM*") {
                                $mainDatabase = $database.name
                            }
                        }
                    }
                    Write-Host "$($mainDatabase) is the main database!`n" -Foregroundcolor green

                    # MAXDOP/THRESHOLD
                    Write-Host "Identifying MaxDOP/Threshold..." -ForegroundColor Cyan
                    $database_maxdop_threshold = Invoke-Command  -Session $sqlSession -ScriptBlock { 
                        param ($server)
                        Invoke-Sqlcmd -Query "SELECT name, value, [description] FROM sys.configurations WHERE name like
                        '%parallel%' ORDER BY name OPTION (RECOMPILE);" -ServerInstance $server.Name
                    } -ArgumentList $environmentsMaster[$x][$y][0][6].Value
                    $maxdop = $database_maxdop_threshold | Where-Object {$_.name -like "cost*"} | Select-Object -property value
                    $cost_threshold = $database_maxdop_threshold | Where-Object {$_.name -like "max*"} | Select-Object -property value
                    Write-Host "Max DOP --- $($maxdop.value) MB"
                    Write-Host "Cost Threshold --- $($cost_threshold.value) MB"
                    
                    # MIN/MAX MEMORY
                    Write-Host "Identifying MIN/MAX Memory..." -ForegroundColor Cyan
                    $database_memory = Invoke-Command -Session $sqlSession -ScriptBlock { 
                        param ($server)
                        Invoke-Sqlcmd -Query "SELECT name, value, [description] FROM sys.configurations WHERE name like
                        '%server memory%' ORDER BY name OPTION (RECOMPILE);" -ServerInstance $server.Name
                    } -ArgumentList $environmentsMaster[$x][$y][0][6].Value 
                    $database_memory_max = $database_memory | where-Object {$_.name -like "max*"} | Select-Object -property value
                    $database_memory_min = $database_memory | where-Object {$_.name -like "min*"} | Select-Object -property value
                    Write-Host "Max Server Memory --- $($database_memory_max.value) MB"
                    Write-Host "Min Server Memory --- $($database_memory_min.value) MB"
                    
                    # DATABASE ENCRYPTION
                    Write-Host "Identifying Database Encryption..." -ForegroundColor Cyan
                    $database_encryption = Invoke-Command -Session $sqlSession -ScriptBlock { 
                        param ($server)
                        Invoke-Sqlcmd -Query "SELECT
                        db.name,
                        db.is_encrypted
                        FROM
                        sys.databases db
                        LEFT OUTER JOIN sys.dm_database_encryption_keys dm
                            ON db.database_id = dm.database_id;
                        GO" -ServerInstance $server 
                    } -ArgumentList $environmentsMaster[$x][$y][0][6].Value
                    $dbEncryption = $database_encryption | Where-Object {$_.name -eq $mainDatabase}
                    Write-Host "$($dbEncryption.name) --- $($dbEncryption.is_encrypted)"

                    # DATABASE SIZE (MAIN)
                    Write-Host "Calculating Database Size" -ForegroundColor Cyan
                    $database_dbSize = Invoke-Command -Session $sqlSession -ScriptBlock { 
                        param ($server, $mainDatabase)        
                        Invoke-Sqlcmd -Query "USE $($mainDatabase)
                        GO
                        exec sp_spaceused
                        GO" -ServerInstance $server.Name 
                    } -ArgumentList $environmentsMaster[$x][$y][0][6].Value, $mainDatabase
                    Write-Host "$($database_dbSize.database_name) --- $($database_dbSize.database_size)"

                    # CUSTOM MODELS
                    Write-Host "Calculating Custom Models..." -ForegroundColor Cyan
                    $database_custom_models = Invoke-Command -Session $sqlSession -ScriptBlock { 
                        param ($server, $mainDatabase)        
                        Invoke-Sqlcmd -Query "USE $($mainDatabase);
                        SELECT * FROM ip.olap_properties 
                        WHERE bism_ind ='N' 
                        AND olap_obj_name 
                        NOT like 'PVE%'" -ServerInstance $server.Name 
                    } -ArgumentList $environmentsMaster[$x][$y][0][6].Value, $mainDatabase | Select-Object -property olap_obj_name
                    foreach ($model in $database_custom_models.olap_obj_name) {
                        Write-Host $model
                    }          

                    # INTERFACES
                    Write-Host "Identifying Interfaces..." -ForegroundColor Cyan
                    $database_interfaces = Invoke-Command -Session $sqlSession -ScriptBlock { 
                        param ($server, $mainDatabase)        
                        Invoke-Sqlcmd -Query "USE $($mainDatabase);
                        SELECT
                        s.description JobStreamName,
                        j.description JobName,
                        j.job_order JobOrder,
                        j.job_id JobID,
                        p.name ParamName,
                        p.param_value ParamValue,
                        MIN(r.last_started) JobLastStarted,
                        MAX(r.last_finished) JobLastFinished,
                        MAX(CONVERT(CHAR(8), DATEADD(S,DATEDIFF(S,r.last_started,r.last_finished),'1900-1-1'),8)) Duration
                        FROM ip.job_stream_job j
                        INNER JOIN ip.job_stream s
                        ON j.job_stream_id = s.job_stream_id
                        INNER JOIN ip.job_stream_schedule ss
                        ON ss.job_stream_id = s.job_stream_id
                        INNER JOIN ip.job_run_status r
                        ON s.job_stream_id = r.job_stream_id
                        LEFT JOIN ip.job_param p
                        ON j.job_id = p.job_id
                        WHERE P.Name = 'Command'
                        GROUP BY
                        s.description,
                        j.description,
                        j.job_order,
                        j.job_id,
                        p.name,
                        p.param_value;" -ServerInstance $server.Name 
                    } -ArgumentList $environmentsMaster[$x][$y][0][6].Value, $mainDatabase
                    $database_interfaces.ParamValue

                    # LICENSE COUNT
                    Write-Host "Calculating License Count..." -ForegroundColor Cyan
                    $database_license_count = Invoke-Command -Session $sqlSession -ScriptBlock { 
                        param ($server, $mainDatabase)        
                        Invoke-Sqlcmd -Query "USE $($mainDatabase);
                        SELECT
                        LicenseRole,
                        COUNT(UserName) UserCount,
                        r.seats LicenseCount
                        FROM (
                        SELECT
                        s1.description LicenseRole,
                        s1.structure_code LicenseCode,
                        u.active_ind Active,
                        u.full_name UserName
                        FROM ip.ip_user u
                        INNER JOIN ip.structure s
                        ON u.role_code = s.structure_code
                        INNER JOIN ip.structure s1
                        ON s.father_code = s1.structure_code
                        WHERE u.active_ind = 'Y'
                        ) q
                        INNER JOIN ip.ip_role r
                        ON q.LicenseCode = r.role_code
                        GROUP BY
                        LicenseRole,
                        LicenseCode,
                        r.seats" -ServerInstance $server.Name 
                    } -ArgumentList $environmentsMaster[$x][$y][0][6].Value, $mainDatabase
                    $licenseProperties = $database_license_count | Select-Object -Property LicenseRole,LicenseCount
                    $totalLicenseCount = 0
                    foreach ($license in $licenseProperties){
                        Write-Output "$($license.LicenseRole): $($license.LicenseCount)"
                        $totalLicenseCount += $license.LicenseCount
                    }
                    Write-Output "Total License Count: $($totalLicenseCount)"
                    
                    # PROGRESSING WEB VERSION
                    Write-Host "Identifying Progressing Web Version..." -ForegroundColor Cyan
                    $database_progressing_web_version = Invoke-Command -Session $sqlSession -ScriptBlock { 
                        param ($server, $mainDatabase)        
                        Invoke-Sqlcmd -Query "USE $($mainDatabase); SELECT TOP 1 sub_release 
                        FROM ip.pv_version 
                        WHERE release = 'PROGRESSING_WEB'
                        ORDER BY seq DESC;" -ServerInstance $server.Name 
                    } -ArgumentList $environmentsMaster[$x][$y][0][6].Value, $mainDatabase
                    $database_progressing_web_version.sub_release 

                <# EXCEL LOGIC AND VARIABLES#>
                $buildData.Cells.Item(11,2)= $environmentsMaster[$x][$y][0][6].Value.Substring(($environmentsMaster[$x][$y][0][6].Value.Length - 2), 2)
                $buildData.Cells.Item(44,2)= $database_dbSize.database_size
                $buildData.Cells.Item(43,2)= $database_memory_max.value
                $buildData.Cells.Item(42,2)= $database_memory_min.value

                $buildData.Cells.Item(54,2)= "$($environmentsMaster[$x][$y][0][6].Value)"
                $buildData.Cells.Item(54,3)= "$(@($environmentsMaster[$x][$y][2][2]).Count)"
                $buildData.Cells.Item(54,4)= "$([MATH]::Round((($environmentsMaster[$x][$y][2][1].MaxCapacity) / 1000000),2))"
                $buildData.Cells.Item(54,5)= $hdStringArray
                $buildData.Cells.Item(54,6)= $diskResize
                $buildData.Cells.Item(54,7)= $task_array

                <# AWS-SPECIFIC VARIABLES #>
                $buildData.Cells.Item(54,8)= $instanceSize
                $buildData.Cells.Item(54,9)= $availabilityZone
                $buildData.Cells.Item(54,10)= $ipAddress
                $buildData.Cells.Item(54,11)= $instanceID

                $buildData.Cells.Item(26,2)= $database_progressing_web_version.sub_release

                $buildData.Cells.Item(28,2)= $database_custom_models.Count
                $modelCount = 0;
                foreach ($model in $database_custom_models.olap_obj_name){
                    $buildData.Cells.Item(91, (2 + $modelCount))= $model
                    $modelCount++
                }

                $databaseCount = 0
                foreach ($database in $all_databases) {
                    $buildData.Cells.Item(99, (2 + $databaseCount))= $database.name
                    $buildData.Cells.Item(100, (2 + $databaseCount))= $database.Size_MB
                    $databaseCount++
                }

                $buildData.Cells.Item(30,2)= $database_interfaces.ParamValue.Count
                $interfaceCount = 0
                foreach ($interface in $database_interfaces.ParamValue) {
                    $buildData.Cells.Item(95, (2 + $interfaceCount))= $interface
                    $interfaceCount++
                }

                $buildData.Cells.Item(41,2)= $dbEncryption.is_encrypted
                $buildData.Cells.Item(22,2)= $totalLicenseCount
                $buildData.Cells.Item(40,2)= $cost_threshold.value            
                $buildData.Cells.Item(39,2)= $maxdop.value
                $buildData.Cells.Item(105,2)= $maintenanceDay
                
                Remove-PSSession -Session $sqlSession

                Write-Host "`n" -ForegroundColor Red  
        }

            ##########################
            # PRODUCTION PVE SERVER 
            ##########################
            elseif ($environmentsMaster[$x][$y][0][6].Value.Substring(($environmentsMaster[$x][$y][0][6].Value.Length - 5), 3) -eq "pve") {
                Write-Host "THIS IS THE PRODUCTION PVE SERVER" -ForegroundColor Cyan

                <# CPU/RAM #>
                Write-Host "Server CPU and RAM" -ForegroundColor Red
                Write-Host "Server Name: $($environmentsMaster[$x][$y][0][6].Value)"
                Write-Host "Server CPUs: $(@($environmentsMaster[$x][$y][2][2]).Count)"
                Write-Host "Server RAM: $([MATH]::Round((($environmentsMaster[$x][$y][2][1].MaxCapacity) / 1000000),2))"  

                <# HARDDRIVES #>
                Write-Host "Disks and Disk Capacity" -ForegroundColor Red
                $diskResize = "Yes"
                $hdStringArray = ""
                foreach ($hd in $environmentsMaster[$x][$y][2][0]) {
                    $hdString = "$($hd.DeviceID) $([MATH]::Round($hd.Size / 1GB,2))gb"
                    $hdStringArray += "$($hdString)`n"
                    Write-Host $hdString  
                    if (([MATH]::Round($hd.Size / 1GB,2)) -gt 60) {
                        $diskResize = "No"  
                    }
                }
                Write-Host "Standard Size Disks (less than 60GB): $($diskResize)"

                <# AWS METADATA #>

                        # INSTANCE SIZE #
                        Write-Host "Instance Size" -ForegroundColor Red
                        $instanceSize = $environmentsMaster[$x][$y][1][2]
                        Write-Host $instanceSize
                        
                        # AVAILABILITY ZONE #
                        Write-Host "Availability Zone" -ForegroundColor Red
                        $availabilityZone = $environmentsMaster[$x][$y][1][3]
                        Write-Host $availabilityZone

                        # IP ADDRESS #
                        Write-Host "IP Address" -ForegroundColor Red
                        $ipAddress = $environmentsMaster[$x][$y][1][4]
                        Write-Host $ipAddress

                        # INSTANCE ID #
                        Write-Host "Instance ID" -ForegroundColor Red
                        $instanceID = $environmentsMaster[$x][$y][1][0]
                        Write-Host $instanceID

                <# SCHEDULED TASKS #>
                Write-Host "Scheduled Tasks on Server" -ForegroundColor Red
                $tasks = Invoke-Command -ComputerName "$($environmentsMaster[$x][$y][0][6].Value)$($domain)" -ScriptBlock {
                    Get-ScheduledTask -TaskPath "\" | Select-Object -Property TaskName, LastRunTime | Where-Object TaskName -notlike "Op*" 
                } -Credential $credentials

                $task_array = ""
                foreach ($task in $tasks){
                    Write-Host "Task Name: $($task.TaskName)"
                    $task_array += "$($task.TaskName)`n"
                }

                <# CURRENT VERSION #>
                Write-Host "Current Environment Version" -ForegroundColor Red
                $crVersion = Invoke-Command -ComputerName "$($environmentsMaster[$x][$y][0][6].Value)$($domain)" -Credential $credentials -ScriptBlock {
                    Get-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\Planview\WebServerPlatform"
                }
                Write-Host $crVersion.CrVersion

                <# MAJOR VERSION #>
                Write-Host "Major Version" -ForegroundColor Red
                $majorVersion = $crVersion.CrVersion.Split('.')[0]
                $majorVersion

                <# OPEN SUITE #>
                Write-Host "OpenSuite" -ForegroundColor Red
                $opensuite = Invoke-Command -ComputerName "$($environmentsMaster[$x][$y][0][6].Value)$($domain)" -Credential $credentials -ScriptBlock {
                    if ((Test-Path -Path "C:\ProgramData\Actian" -PathType Container) -And (Test-Path -Path "F:\Planview\Interfaces\OpenSuite" -PathType Container)) {

                        $software = "*Actian*";
                        $installed = (Get-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*" | Where { $_.DisplayName -like $software }) -ne $null

                        if ($installed) {
                            return "Yes"
                        }
                        
                    } else {
                        return "No"
                    }
                }
                Write-Host "OpenSuite Detected: $($opensuite)"

                <# CONNECTORS #>
                Write-Host "Connectors" -ForegroundColor Red
                $PRMAdapter = "False"
                $PPAdapter = "False"
                $LKAdapter = "False"

                $PRMini = Invoke-Command -ComputerName "$($environmentsMaster[$x][$y][0][6].Value)$($domain)" -Credential $credentials -ScriptBlock {

                    if (Test-Path -Path "F:\Planview\midtier\webserver\objects\PRM_Adapter_Config.ini" -PathType leaf){

                        $PRMiniContent = (Get-Content -Path "F:\Planview\midtier\webserver\objects\PRM_Adapter_Config.ini" | Where-Object {$_.length -ne 0}) | Where-Object {$_.substring(0,1) -ne ";"}

                        $masterArray = @()
                        $databaseName = ""
                        foreach($x in $PRMiniContent){

                            $valuePair = New-Object PSObject
                            Add-Member -InputObject $valuePair -MemberType NoteProperty -Name Database -Value ""
                            Add-Member -InputObject $valuePair -MemberType NoteProperty -Name Key -Value ""

                            if ($x -match "^\["){

                                $databaseName = $x.substring(1, ($x.length - 2)) 

                                $valuePair.Database = $databaseName
                                $valuePair.Key = $databaseName
                                $masterArray += $valuePair
                            
                            
                            } else {       
                            
                                $valuePair.Database = $databaseName
                                $valuePair.Key = $x 
                                $masterArray += $valuePair
                            
                            }           
                        }
                    
                        return $masterArray
                    
                    } else {

                        return 0

                    }
                }

                if ($PRMini -ne '0'){

                    $PRMAdapter = "True"

                    Write-Host "PRM Adapter Found"
                    $LKKey = ($PRMini | Where-Object {$_.Database -like "*PROD"} | Where-Object {$_.Key -like "use_prm_leankit*"}).Key
                    $PPKey = ($PRMini | Where-Object {$_.Database -like "*PROD"} | Where-Object {$_.Key -like "use_prm_projectplace*"}).Key

                    if ($LKKey -like "*true*"){
                        $LKAdapter = "True"
                        Write-Host "LeanKit Connector --- True"
                    } else {
                        Write-Host "LeanKit Connector --- False"
                    }

                    if ($PPKey -like "*true*"){
                        $PPAdapter = "True"
                        Write-Host "ProjectPlace Connector --- True"
                    } else {
                        Write-Host "ProjectPlace Connector --- False"
                    }

                } else {

                    Write-Host "PRM Adapter not found"

                    $LegacyPPAdapter = Invoke-Command -ComputerName "$($environmentsMaster[$x][$y][0][6].Value)$($domain)" -Credential $credentials -ScriptBlock {
                        Test-Path -Path "F:\Planview\midtier\webserver\objects\ProjectPlace_Config.ini" -PathType leaf
                    }

                    $PPAdapter = $LegacyPPAdapter
                    Write-Host "ProjectPlace (Legacy install) --- $($PPAdapter)"

                }

                # NEW RELIC #
                Write-Host "New Relic" -ForegroundColor Red
                $newRelic = Invoke-Command -ComputerName "$($environmentsMaster[$x][$y][0][6].Value)$($domain)" -Credential $credentials -ScriptBlock {
                    if (Test-Path -Path "C:\ProgramData\New Relic" -PathType Container ) {
                        Write-Host "New Relic has been detected on this server"
                        return "Yes"
                    } else {
                        Write-Host "New Relic was not detected on this server"
                        return "No"
                    }
                }

                    # GET WEB CONFIG #
                    $webConfig = Invoke-Command -ComputerName "$($environmentsMaster[$x][$y][0][6].Value)$($domain)" -Credential $credentials -ScriptBlock {
                        return Get-Content -Path "F:\Planview\MidTier\ODataService\Web.config"
                    }
                    $webConfig = [xml] $webConfig

                <# PRODUCTION URL #>
                Write-Host "Production URL" -ForegroundColor Red
                $environmentURL = $webConfig.configuration.appSettings.add | Where-Object {$_.key -eq "PveUrl"} | Select-Object -Property value
                Write-Host $environmentURL.value

                <# DNS ALIAS #>
                Write-Host "Production DNS Alias" -ForegroundColor Red            
                $dnsAlias = ($environmentURL.value.Split('//')[2]).Split('.')[0] 
                $dnsAlias

                <# REPORT FARM URL #>
                Write-Host "Report Farm URL" -ForegroundColor Red
                $reportfarmURL = $webConfig.configuration.appSettings.add | Where-Object {$_.key -eq "Report_Server_Web_Service_URL"} | Select-Object -Property value
                Write-Host $reportfarmURL.value

                <# ENCRYPTED PVMASTER PASSWORD #>
                Write-Host "Encrypted PVMaster Password" -ForegroundColor Red
                $encryptedPVMasterPassword = $webConfig.configuration.appSettings.add | Where-Object {$_.key -eq "PveUserPassword"} | Select-Object -Property value
                Write-Host $encryptedPVMasterPassword.value
                
                <# UNENCRYPTED PVMASTER PASSWORD #>
                Write-Host "Unencrypted PVMaster Password" -ForegroundColor Red
                $unencryptedPVMasterPassword = Invoke-PassUtil -InputString $encryptedPVMasterPassword.value -Deobfuscation
                Write-Host $unencryptedPVMasterPassword

                <# IP RESTRICTIONS #>
                Write-Host "IP Restrictions on F5" -ForegroundColor Red
                Write-Host "Currently Not Automated"
                $IPRestrictions = "Not Automated"
<#                    
                    # Authentication on the F5 #
                    $websession =  New-Object Microsoft.PowerShell.Commands.WebRequestSession
                    $jsonbody = @{username = $f5Credentials.UserName ; password = $f5Credentials.GetNetworkCredential().Password; loginProviderName='tmos'} | ConvertTo-Json
                    $authResponse = Invoke-RestMethodOverride -Method Post -Uri "https://$($f5ip)/mgmt/shared/authn/login" -Credential $f5Credentials -Body $jsonbody -ContentType 'application/json'
                    $token = $authResponse.token.token
                    $websession.Headers.Add('X-F5-Auth-Token', $Token)

                    # Calling data-group REST endpoint and parsing IPRestrictions list #
                    $IPRestrictionsList = (Invoke-RestMethod  -Uri "https://$($f5ip)/mgmt/tm/ltm/data-group/internal" -WebSession $websession).Items | 
                        Where-Object {$_.name -eq "IPRestrictions"} | Select-Object -Property records

                    foreach ($record in $IPRestrictionsList.records) {
                        if ($record.name -eq "$($dnsAlias).pvcloud.com") {
                            $IPRestrictions = "Yes"
                            Write-Host "IP restrctions were found for $($dnsAlias).pvcloud.com"
                        }
                    }

                    if ($IPRestrictions -eq "No") {
                        Write-Host "No IP restrictions found for $($dnsAlias).pvcloud.com"
                    }
#>                
                <# EXCEL LOGIC AND VARIABLES#>
                $webServerCount++
                $buildData.Cells.Item(24,2)= $webServerCount
                $buildData.Cells.Item(1,2)= $environmentURL.value
                $buildData.Cells.Item(7,2)= $dnsAlias
                $buildData.Cells.Item(46,2)= $reportfarmURL.value
                $buildData.Cells.Item(15,2)= $newRelic
                $buildData.Cells.Item(35,2)= $IPRestrictions
                $buildData.Cells.Item(36,2)= $opensuite
                $buildData.Cells.Item(25,2)= $crVersion.CrVersion
                $buildData.Cells.Item(13,2)= $majorVersion
                $buildData.Cells.Item(19,2)= "True"
                $buildData.Cells.Item(17,2)= $encryptedPVMasterPassword.value
                $buildData.Cells.Item(16,2)= $unencryptedPVMasterPassword
                

                $buildData.Cells.Item(31,2)= $PPAdapter
                $buildData.Cells.Item(32,2)= $LKAdapter
                $buildData.Cells.Item(37,2)= $PRMAdapter
                
                

                $buildData.Cells.Item(56,2)= "$($environmentsMaster[$x][$y][0][6].Value)"
                $buildData.Cells.Item(56,3)= "$(@($environmentsMaster[$x][$y][2][2]).Count)"
                $buildData.Cells.Item(56,4)= "$([MATH]::Round((($environmentsMaster[$x][$y][2][1].MaxCapacity) / 1000000),2))"
                $buildData.Cells.Item(56,5)= $hdStringArray
                $buildData.Cells.Item(56,6)= $diskResize
                $buildData.Cells.Item(56,7)= $task_array

                <# AWS-SPECIFIC VARIABLES #>
                $buildData.Cells.Item(56,8)= $instanceSize
                $buildData.Cells.Item(56,9)= $availabilityZone
                $buildData.Cells.Item(56,10)= $ipAddress
                $buildData.Cells.Item(56,11)= $instanceID

                Write-Host "`n" -ForegroundColor Red  
            }

        }

    } 
    
    if ($environmentsMaster[$x][0] -eq $slot2) { 
        Write-Host ":::::::: $($environmentsMaster[$x][0]) Environment ::::::::" -Foregroundcolor Yellow
        
        $webServerCount = 0
        for ($y=1; $y -lt $environmentsMaster[$x].Length; $y++) {        

            if ($environmentsMaster[$x][$y][0][6].Value.Substring(($environmentsMaster[$x][$y][0][6].Value.Length - 5), 3) -eq "app" -or $environmentsMaster[$x][$y][0][6].Value.Substring(($environmentsMaster[$x][$y][0][6].Value.Length - 5), 3) -eq "ctm") {
        
                #######################
                # SANDBOX APP SERVER 
                #######################
                if ($environmentsMaster[$x][$y][0][6].Value.Substring(($environmentsMaster[$x][$y][0][6].Value.Length - 5), 3) -eq "app") {
                    Write-Host "THIS IS THE SANDBOX APP SERVER" -ForegroundColor Cyan
        
                    <# CPU/RAM #>
                    Write-Host "Server CPU and RAM" -ForegroundColor Red
                    Write-Host "Server Name: $($environmentsMaster[$x][$y][0][6].Value)"
                    Write-Host "Server CPUs: $(@($environmentsMaster[$x][$y][2][2]).Count)"
                    Write-Host "Server RAM: $([MATH]::Round((($environmentsMaster[$x][$y][2][1].MaxCapacity) / 1000000),2))"
        
                    <# HARDDRIVES #>
                    Write-Host "Disks and Disk Capacity" -ForegroundColor Red
                    $diskResize = "Yes"
                    $hdStringArray = ""
                    foreach ($hd in $environmentsMaster[$x][$y][2][0]) {
                        $hdString = "$($hd.DeviceID) $([MATH]::Round($hd.Size / 1GB,2))gb"
                        $hdStringArray += "$($hdString)`n"
                        Write-Host $hdString  
                        if (([MATH]::Round($hd.Size / 1GB,2)) -gt 60) {
                            $diskResize = "No"  
                        }
                    }
                    Write-Host "Standard Size Disks (less than 60GB): $($diskResize)"
        
                    <# AWS METADATA #>

                        # INSTANCE SIZE #
                        Write-Host "Instance Size" -ForegroundColor Red
                        $instanceSize = $environmentsMaster[$x][$y][1][2]
                        Write-Host $instanceSize
                        
                        # AVAILABILITY ZONE #
                        Write-Host "Availability Zone" -ForegroundColor Red
                        $availabilityZone = $environmentsMaster[$x][$y][1][3]
                        Write-Host $availabilityZone

                        # IP ADDRESS #
                        Write-Host "IP Address" -ForegroundColor Red
                        $ipAddress = $environmentsMaster[$x][$y][1][4]
                        Write-Host $ipAddress

                        # INSTANCE ID #
                        Write-Host "Instance ID" -ForegroundColor Red
                        $instanceID = $environmentsMaster[$x][$y][1][0]
                        Write-Host $instanceID
        
                    <# SCHEDULED TASKS #>
                    Write-Host "Scheduled Tasks on Server" -ForegroundColor Red
                    $tasks = Invoke-Command -ComputerName "$($environmentsMaster[$x][$y][0][6].Value)$($domain)" -ScriptBlock {
                        Get-ScheduledTask -TaskPath "\" | Select-Object -Property TaskName, LastRunTime | Where-Object TaskName -notlike "Op*" 
                    } -Credential $credentials

                    $task_array = ""
                    foreach ($task in $tasks){
                        Write-Host "Task Name: $($task.TaskName)"
                        $task_array += "$($task.TaskName)`n"
                    }
                    
                    <# OPEN SUITE #>
                    Write-Host "OpenSuite" -ForegroundColor Red
                    $opensuite = Invoke-Command -ComputerName "$($environmentsMaster[$x][$y][0][6].Value)$($domain)" -Credential $credentials -ScriptBlock {
                        if ((Test-Path -Path "C:\ProgramData\Actian" -PathType Container) -And (Test-Path -Path "F:\Planview\Interfaces\OpenSuite" -PathType Container)) {
        
                            $software = "*Actian*";
                            $installed = (Get-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*" | Where { $_.DisplayName -like $software }) -ne $null
        
                            if ($installed) {
                                return "Yes"
                            }
                            
                        } else {
                            return "No"
                        }
                    }
                    Write-Host "OpenSuite Detected: $($opensuite)"
        
                    <# CONNECTORS #>
                    Write-Host "Connectors" -ForegroundColor Red
                    $PRMAdapter = "False"
                    $PPAdapter = "False"
                    $LKAdapter = "False"

                    $PRMini = Invoke-Command -ComputerName "$($environmentsMaster[$x][$y][0][6].Value)$($domain)" -Credential $credentials -ScriptBlock {

                        if (Test-Path -Path "F:\Planview\midtier\webserver\objects\PRM_Adapter_Config.ini" -PathType leaf){

                            $PRMiniContent = (Get-Content -Path "F:\Planview\midtier\webserver\objects\PRM_Adapter_Config.ini" | Where-Object {$_.length -ne 0}) | Where-Object {$_.substring(0,1) -ne ";"}

                            $masterArray = @()
                            $databaseName = ""
                            foreach($x in $PRMiniContent){

                                $valuePair = New-Object PSObject
                                Add-Member -InputObject $valuePair -MemberType NoteProperty -Name Database -Value ""
                                Add-Member -InputObject $valuePair -MemberType NoteProperty -Name Key -Value ""

                                if ($x -match "^\["){

                                    $databaseName = $x.substring(1, ($x.length - 2)) 

                                    $valuePair.Database = $databaseName
                                    $valuePair.Key = $databaseName
                                    $masterArray += $valuePair
                                
                                
                                } else {       
                                
                                    $valuePair.Database = $databaseName
                                    $valuePair.Key = $x 
                                    $masterArray += $valuePair
                                
                                }           
                            }
                        
                            return $masterArray
                        
                        } else {

                            return 0

                        }
                    }

                    if ($PRMini -ne '0'){

                        $PRMAdapter = "True"

                        Write-Host "PRM Adapter Found"
                        $LKKey = ($PRMini | Where-Object {$_.Database -like "*SANDBOX*"} | Where-Object {$_.Key -like "use_prm_leankit*"}).Key
                        $PPKey = ($PRMini | Where-Object {$_.Database -like "*SANDBOX*"} | Where-Object {$_.Key -like "use_prm_projectplace*"}).Key

                        if ($LKKey -like "*true*"){
                            $LKAdapter = "True"
                            Write-Host "LeanKit Connector --- True"
                        } else {
                            Write-Host "LeanKit Connector --- False"
                        }

                        if ($PPKey -like "*true*"){
                            $PPAdapter = "True"
                            Write-Host "ProjectPlace Connector --- True"
                        } else {
                            Write-Host "ProjectPlace Connector --- False"
                        }

                    } else {

                        Write-Host "PRM Adapter not found"

                        $LegacyPPAdapter = Invoke-Command -ComputerName "$($environmentsMaster[$x][$y][0][6].Value)$($domain)" -Credential $credentials -ScriptBlock {
                            Test-Path -Path "F:\Planview\midtier\webserver\objects\ProjectPlace_Config.ini" -PathType leaf
                        }

                        $PPAdapter = $LegacyPPAdapter
                        Write-Host "ProjectPlace (Legacy install) --- $($PPAdapter)"

                    }
                    
                    <# EXCEL LOGIC AND VARIABLES#>
                    $buildData.Cells.Item(72,2)= "$($environmentsMaster[$x][$y][0][6].Value)"
                    $buildData.Cells.Item(72,3)= "$(@($environmentsMaster[$x][$y][2][2]).Count)"
                    $buildData.Cells.Item(72,4)= "$([MATH]::Round((($environmentsMaster[$x][$y][2][1].MaxCapacity) / 1000000),2))"
                    $buildData.Cells.Item(72,5)= $hdStringArray
                    $buildData.Cells.Item(72,6)= $diskResize
                    $buildData.Cells.Item(72,7)= $task_array

                    <# AWS-SPECIFIC VARIABLES #>
                    $buildData.Cells.Item(72,8)= $instanceSize
                    $buildData.Cells.Item(72,9)= $availabilityZone
                    $buildData.Cells.Item(72,10)= $ipAddress
                    $buildData.Cells.Item(72,11)= $instanceID
        
                    $buildData.Cells.Item(31,3)= $PPAdapter
                    $buildData.Cells.Item(32,3)= $LKAdapter
                    $buildData.Cells.Item(37,3)= $PRMAdapter
                    
                    $buildData.Cells.Item(36,3)= $opensuite
        
                    Write-Host "`n" -ForegroundColor Red
                }
        
                ###############################
                # SANDBOX CTM SERVER (Troux) 
                ###############################
                elseif ($environmentsMaster[$x][$y][0][6].Value.Substring(($environmentsMaster[$x][$y][0][6].Value.Length - 5), 3) -eq "ctm") {
                    Write-Host "THIS IS THE SANDBOX TROUX SERVER" -ForegroundColor Cyan
        
                    <# CPU/RAM #>
                    Write-Host "Server CPU and RAM" -ForegroundColor Red
                    Write-Host "Server Name: $($environmentsMaster[$x][$y][0][6].Value)"
                    Write-Host "Server CPUs: $(@($environmentsMaster[$x][$y][2][2]).Count)"
                    Write-Host "Server RAM: $([MATH]::Round((($environmentsMaster[$x][$y][2][1].MaxCapacity) / 1000000),2))"    
        
                    <# HARDDRIVES #>
                    Write-Host "Disks and Disk Capacity" -ForegroundColor Red
                    $diskResize = "Yes"
                    $hdStringArray = ""
                    foreach ($hd in $environmentsMaster[$x][$y][2][0]) {
                        $hdString = "$($hd.DeviceID) $([MATH]::Round($hd.Size / 1GB,2))gb"
                        $hdStringArray += "$($hdString)`n"
                        Write-Host $hdString  
                        if (([MATH]::Round($hd.Size / 1GB,2)) -gt 60) {
                            $diskResize = "No"  
                        }
                    }
                    Write-Host "Standard Size Disks (less than 60GB): $($diskResize)"       
        
                    <# AWS METADATA #>

                        # INSTANCE SIZE #
                        Write-Host "Instance Size" -ForegroundColor Red
                        $instanceSize = $environmentsMaster[$x][$y][1][2]
                        Write-Host $instanceSize
                        
                        # AVAILABILITY ZONE #
                        Write-Host "Availability Zone" -ForegroundColor Red
                        $availabilityZone = $environmentsMaster[$x][$y][1][3]
                        Write-Host $availabilityZone

                        # IP ADDRESS #
                        Write-Host "IP Address" -ForegroundColor Red
                        $ipAddress = $environmentsMaster[$x][$y][1][4]
                        Write-Host $ipAddress

                        # INSTANCE ID #
                        Write-Host "Instance ID" -ForegroundColor Red
                        $instanceID = $environmentsMaster[$x][$y][1][0]
                        Write-Host $instanceID
        
                    <# SCHEDULED TASKS #>
                    Write-Host "Scheduled Tasks on Server" -ForegroundColor Red
                    $tasks = Invoke-Command -ComputerName "$($environmentsMaster[$x][$y][0][6].Value)$($domain)" -ScriptBlock {
                        Get-ScheduledTask -TaskPath "\" | Select-Object -Property TaskName, LastRunTime | Where-Object TaskName -notlike "Op*" 
                    } -Credential $credentials

                    $task_array = ""
                    foreach ($task in $tasks){
                        Write-Host "Task Name: $($task.TaskName)"
                        $task_array += "$($task.TaskName)`n"
                    }
                    
                    <# EXCEL LOGIC AND VARIABLES#>
                    $buildData.Cells.Item(73,2)= "$($environmentsMaster[$x][$y][0][6].Value)"
                    $buildData.Cells.Item(73,3)= "$(@($environmentsMaster[$x][$y][2][2]).Count)"
                    $buildData.Cells.Item(73,4)= "$([MATH]::Round((($environmentsMaster[$x][$y][2][1].MaxCapacity) / 1000000),2))"
                    $buildData.Cells.Item(73,5)= $hdStringArray
                    $buildData.Cells.Item(73,6)= $diskResize
                    $buildData.Cells.Item(73,7)= $task_array

                    <# AWS-SPECIFIC VARIABLES #>
                    $buildData.Cells.Item(73,8)= $instanceSize
                    $buildData.Cells.Item(73,9)= $availabilityZone
                    $buildData.Cells.Item(73,10)= $ipAddress
                    $buildData.Cells.Item(73,11)= $instanceID
        
                    Write-Host "`n" -ForegroundColor Red  
                }
            }
        
            #######################
            # SANDBOX WEB SERVER 
            #######################
            elseif ($environmentsMaster[$x][$y][0][6].Value.Substring(($environmentsMaster[$x][$y][0][6].Value.Length - 5), 3) -eq "web") {
                Write-Host "THIS IS THE SANDBOX WEB SERVER" -ForegroundColor Cyan
        
                <# CPU/RAM #>
                Write-Host "Server CPU and RAM" -ForegroundColor Red
                Write-Host "Server Name: $($environmentsMaster[$x][$y][0][6].Value)"
                Write-Host "Server CPUs: $(@($environmentsMaster[$x][$y][2][2]).Count)"
                Write-Host "Server RAM: $([MATH]::Round((($environmentsMaster[$x][$y][2][1].MaxCapacity) / 1000000),2))"
        
                <# HARDDRIVES #>
                Write-Host "Disks and Disk Capacity" -ForegroundColor Red
                $diskResize = "Yes"
                $hdStringArray = ""
                foreach ($hd in $environmentsMaster[$x][$y][2][0]) {
                    $hdString = "$($hd.DeviceID) $([MATH]::Round($hd.Size / 1GB,2))gb"
                    $hdStringArray += "$($hdString)`n"
                    Write-Host $hdString  
                    if (([MATH]::Round($hd.Size / 1GB,2)) -gt 60) {
                        $diskResize = "No"  
                    }
                }
                Write-Host "Standard Size Disks (less than 60GB): $($diskResize)"
        
                <# AWS METADATA #>

                        # INSTANCE SIZE #
                        Write-Host "Instance Size" -ForegroundColor Red
                        $instanceSize = $environmentsMaster[$x][$y][1][2]
                        Write-Host $instanceSize
                        
                        # AVAILABILITY ZONE #
                        Write-Host "Availability Zone" -ForegroundColor Red
                        $availabilityZone = $environmentsMaster[$x][$y][1][3]
                        Write-Host $availabilityZone

                        # IP ADDRESS #
                        Write-Host "IP Address" -ForegroundColor Red
                        $ipAddress = $environmentsMaster[$x][$y][1][4]
                        Write-Host $ipAddress

                        # INSTANCE ID #
                        Write-Host "Instance ID" -ForegroundColor Red
                        $instanceID = $environmentsMaster[$x][$y][1][0]
                        Write-Host $instanceID
        
                <# SCHEDULED TASKS #>
                Write-Host "Scheduled Tasks on Server" -ForegroundColor Red
                $tasks = Invoke-Command -ComputerName "$($environmentsMaster[$x][$y][0][6].Value)$($domain)" -ScriptBlock {
                    Get-ScheduledTask -TaskPath "\" | Select-Object -Property TaskName, LastRunTime | Where-Object TaskName -notlike "Op*" 
                } -Credential $credentials

                $task_array = ""
                foreach ($task in $tasks){
                    Write-Host "Task Name: $($task.TaskName)"
                    $task_array += "$($task.TaskName)`n"
                }
        
                <# CURRENT VERSION #>
                Write-Host "Current Environment Version" -ForegroundColor Red
                $crVersion = Invoke-Command -ComputerName "$($environmentsMaster[$x][$y][0][6].Value)$($domain)" -Credential $credentials -ScriptBlock {
                    Get-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\Planview\WebServerPlatform"
                }
                Write-Host $crVersion.CrVersion
        
                <# MAJOR VERSION #>
                Write-Host "Major Version" -ForegroundColor Red
                $majorVersion = $crVersion.CrVersion.Split('.')[0]
                $majorVersion
        
                # NEW RELIC #
                Write-Host "New Relic" -ForegroundColor Red
                $newRelic = Invoke-Command -ComputerName "$($environmentsMaster[$x][$y][0][6].Value)$($domain)" -Credential $credentials -ScriptBlock {
                    if (Test-Path -Path "C:\ProgramData\New Relic" -PathType Container ) {
                        Write-Host "New Relic has been detected on this server"
                        return "Yes"
                    } else {
                        Write-Host "New Relic was not detected on this server"
                        return "No"
                    }
                }
        
                    # GET WEB CONFIG #
                    $webConfig = Invoke-Command -ComputerName "$($environmentsMaster[$x][$y][0][6].Value)$($domain)" -Credential $credentials -ScriptBlock {
                        return Get-Content -Path "F:\Planview\MidTier\ODataService\Web.config"
                    }
                    $webConfig = [xml] $webConfig
        
                <# SANDBOX URL #>
                Write-Host "Sandbox URL" -ForegroundColor Red
                $environmentURL = $webConfig.configuration.appSettings.add | Where-Object {$_.key -eq "PveUrl"} | Select-Object -Property value
                Write-Host $environmentURL.value
        
                <# DNS ALIAS #>
                Write-Host "Sandbox DNS Alias" -ForegroundColor Red            
                $dnsAlias = ($environmentURL.value.Split('//')[2]).Split('.')[0] 
                $dnsAlias
        
                <# REPORT FARM URL #>
                Write-Host "Report Farm URL" -ForegroundColor Red
                $reportfarmURL = $webConfig.configuration.appSettings.add | Where-Object {$_.key -eq "Report_Server_Web_Service_URL"} | Select-Object -Property value
                Write-Host $reportfarmURL.value
        
                <# IP RESTRICTIONS #>
                Write-Host "IP Restrictions on F5" -ForegroundColor Red
                Write-Host "Currently Not Automated"
                $IPRestrictions = "Not Automated"
<#                    
                    # Authentication on the F5 #
                    $websession =  New-Object Microsoft.PowerShell.Commands.WebRequestSession
                    $jsonbody = @{username = $f5Credentials.UserName ; password = $f5Credentials.GetNetworkCredential().Password; loginProviderName='tmos'} | ConvertTo-Json
                    $authResponse = Invoke-RestMethodOverride -Method Post -Uri "https://$($f5ip)/mgmt/shared/authn/login" -Credential $f5Credentials -Body $jsonbody -ContentType 'application/json'
                    $token = $authResponse.token.token
                    $websession.Headers.Add('X-F5-Auth-Token', $Token)
        
                    # Calling data-group REST endpoint and parsing IPRestrictions list #
                    $IPRestrictionsList = (Invoke-RestMethod  -Uri "https://$($f5ip)/mgmt/tm/ltm/data-group/internal" -WebSession $websession).Items | 
                        Where-Object {$_.name -eq "IPRestrictions"} | Select-Object -Property records
        
                    foreach ($record in $IPRestrictionsList.records) {
                        if ($record.name -eq "$($dnsAlias).pvcloud.com") {
                            $IPRestrictions = "Yes"
                            Write-Host "IP restrctions were found for $($dnsAlias).pvcloud.com"
                        }
                    }
        
                    if ($IPRestrictions -eq "No") {
                        Write-Host "No IP restrictions found for $($dnsAlias).pvcloud.com"
                    }
#>        
                <# EXCEL LOGIC AND VARIABLES#>
                $buildData.Cells.Item(25,3)= $crVersion.CrVersion
                $buildData.Cells.Item(35,3)= $IPRestrictions
                $buildData.Cells.Item(2,2)= $environmentURL.value
                $buildData.Cells.Item(8,2)= $dnsAlias
                
        
                if ($webServerCount -gt 0){
                    $buildData.Cells.Item(78 + ($webServerCount - 1),2)= "$($environmentsMaster[$x][$y][0][6].Value)"
                    $buildData.Cells.Item(78 + ($webServerCount - 1),3)= "$(@($environmentsMaster[$x][$y][2][2]).Count)"
                    $buildData.Cells.Item(78 + ($webServerCount - 1),4)= "$([MATH]::Round((($environmentsMaster[$x][$y][2][1].MaxCapacity) / 1000000),2))"
                    $buildData.Cells.Item(78 + ($webServerCount - 1),5)= $hdStringArray
                    $buildData.Cells.Item(78 + ($webServerCount - 1),6)= $diskResize
                    $buildData.Cells.Item(78 + ($webServerCount - 1),7)= $task_array

                    <# AWS-SPECIFIC VARIABLES #>
                    $buildData.Cells.Item(78 + ($webServerCount - 1),8)= $instanceSize
                    $buildData.Cells.Item(78 + ($webServerCount - 1),9)= $availabilityZone
                    $buildData.Cells.Item(78 + ($webServerCount - 1),10)= $ipAddress
                    $buildData.Cells.Item(78 + ($webServerCount - 1),11)= $instanceID
                }
                else {
                    $buildData.Cells.Item(71,2)= "$($environmentsMaster[$x][$y][0][6].Value)"
                    $buildData.Cells.Item(71,3)= "$(@($environmentsMaster[$x][$y][2][2]).Count)"
                    $buildData.Cells.Item(71,4)= "$([MATH]::Round((($environmentsMaster[$x][$y][2][1].MaxCapacity) / 1000000),2))"
                    $buildData.Cells.Item(71,5)= $hdStringArray
                    $buildData.Cells.Item(71,6)= $diskResize
                    $buildData.Cells.Item(71,7)= $task_array

                    <# AWS-SPECIFIC VARIABLES #>
                    $buildData.Cells.Item(71,8)= $instanceSize
                    $buildData.Cells.Item(71,9)= $availabilityZone
                    $buildData.Cells.Item(71,10)= $ipAddress
                    $buildData.Cells.Item(71,11)= $instanceID
                }
        
                $webServerCount++
                $buildData.Cells.Item(24,3)= $webServerCount
        
                Write-Host "`n" -ForegroundColor Red  
        
            }
        
            #######################
            # SANDBOX SAS SERVER 
            #######################
            elseif ($environmentsMaster[$x][$y][0][6].Value.Substring(($environmentsMaster[$x][$y][0][6].Value.Length - 5), 3) -eq "sas") {
                Write-Host "THIS IS THE SANDBOX SAS SERVER" -ForegroundColor Cyan
        
                <# CPU/RAM #>
                Write-Host "Server CPU and RAM" -ForegroundColor Red
                Write-Host "Server Name: $($environmentsMaster[$x][$y][0][6].Value)"
                Write-Host "Server CPUs: $(@($environmentsMaster[$x][$y][2][2]).Count)"
                Write-Host "Server RAM: $([MATH]::Round((($environmentsMaster[$x][$y][2][1].MaxCapacity) / 1000000),2))"
        
                <# HARDDRIVES #>
                Write-Host "Disks and Disk Capacity" -ForegroundColor Red
                $diskResize = "Yes"
                $hdStringArray = ""
                foreach ($hd in $environmentsMaster[$x][$y][2][0]) {
                    $hdString = "$($hd.DeviceID) $([MATH]::Round($hd.Size / 1GB,2))gb"
                    $hdStringArray += "$($hdString)`n"
                    Write-Host $hdString  
                    if (([MATH]::Round($hd.Size / 1GB,2)) -gt 60) {
                        $diskResize = "No"  
                    }
                }
                Write-Host "Standard Size Disks (less than 60GB): $($diskResize)"
        
                <# AWS METADATA #>

                        # INSTANCE SIZE #
                        Write-Host "Instance Size" -ForegroundColor Red
                        $instanceSize = $environmentsMaster[$x][$y][1][2]
                        Write-Host $instanceSize
                        
                        # AVAILABILITY ZONE #
                        Write-Host "Availability Zone" -ForegroundColor Red
                        $availabilityZone = $environmentsMaster[$x][$y][1][3]
                        Write-Host $availabilityZone

                        # IP ADDRESS #
                        Write-Host "IP Address" -ForegroundColor Red
                        $ipAddress = $environmentsMaster[$x][$y][1][4]
                        Write-Host $ipAddress

                        # INSTANCE ID #
                        Write-Host "Instance ID" -ForegroundColor Red
                        $instanceID = $environmentsMaster[$x][$y][1][0]
                        Write-Host $instanceID
        
                <# SCHEDULED TASKS #>
                Write-Host "Scheduled Tasks on Server" -ForegroundColor Red
                $tasks = Invoke-Command -ComputerName "$($environmentsMaster[$x][$y][0][6].Value)$($domain)" -ScriptBlock {
                    Get-ScheduledTask -TaskPath "\" | Select-Object -Property TaskName, LastRunTime | Where-Object TaskName -notlike "Op*" 
                } -Credential $credentials

                $task_array = ""
                foreach ($task in $tasks){
                    Write-Host "Task Name: $($task.TaskName)"
                    $task_array += "$($task.TaskName)`n"
                }
                
                <# EXCEL LOGIC AND VARIABLES#>
                $buildData.Cells.Item(75,2)= "$($environmentsMaster[$x][$y][0][6].Value)"
                $buildData.Cells.Item(75,3)= "$(@($environmentsMaster[$x][$y][2][2]).Count)"
                $buildData.Cells.Item(75,4)= "$([MATH]::Round((($environmentsMaster[$x][$y][2][1].MaxCapacity) / 1000000),2))"
                $buildData.Cells.Item(75,5)= $hdStringArray
                $buildData.Cells.Item(75,6)= $diskResize
                $buildData.Cells.Item(75,7)= $task_array

                <# AWS-SPECIFIC VARIABLES #>
                $buildData.Cells.Item(75,8)= $instanceSize
                $buildData.Cells.Item(75,9)= $availabilityZone
                $buildData.Cells.Item(75,10)= $ipAddress
                $buildData.Cells.Item(75,11)= $instanceID
        
                Write-Host "`n" -ForegroundColor Red  
            }
        
            #######################
            # SANDBOX SQL SERVER 
            #######################
            elseif ($environmentsMaster[$x][$y][0][6].Value.Substring(($environmentsMaster[$x][$y][0][6].Value.Length - 5), 3) -eq "sql") {
                Write-Host "THIS IS THE SANDBOX SQL SERVER" -ForegroundColor Cyan
        
                <# CPU/RAM #>
                Write-Host "Server CPU and RAM" -ForegroundColor Red
                Write-Host "Server Name: $($environmentsMaster[$x][$y][0][6].Value)"
                Write-Host "Server CPUs: $(@($environmentsMaster[$x][$y][2][2]).Count)"
                Write-Host "Server RAM: $([MATH]::Round((($environmentsMaster[$x][$y][2][1].MaxCapacity) / 1000000),2))"
        
                <# HARDDRIVES #>
                Write-Host "Disks and Disk Capacity" -ForegroundColor Red
                $diskResize = "Yes"
                $hdStringArray = ""
                foreach ($hd in $environmentsMaster[$x][$y][2][0]) {
                    $hdString = "$($hd.DeviceID) $([MATH]::Round($hd.Size / 1GB,2))gb"
                    $hdStringArray += "$($hdString)`n"
                    Write-Host $hdString  
                    if (([MATH]::Round($hd.Size / 1GB,2)) -gt 60) {
                        $diskResize = "No"  
                    }
                }
                Write-Host "Standard Size Disks (less than 60GB): $($diskResize)"
        
                <# AWS METADATA #>

                        # INSTANCE SIZE #
                        Write-Host "Instance Size" -ForegroundColor Red
                        $instanceSize = $environmentsMaster[$x][$y][1][2]
                        Write-Host $instanceSize
                        
                        # AVAILABILITY ZONE #
                        Write-Host "Availability Zone" -ForegroundColor Red
                        $availabilityZone = $environmentsMaster[$x][$y][1][3]
                        Write-Host $availabilityZone

                        # IP ADDRESS #
                        Write-Host "IP Address" -ForegroundColor Red
                        $ipAddress = $environmentsMaster[$x][$y][1][4]
                        Write-Host $ipAddress

                        # INSTANCE ID #
                        Write-Host "Instance ID" -ForegroundColor Red
                        $instanceID = $environmentsMaster[$x][$y][1][0]
                        Write-Host $instanceID
        
                <# SCHEDULED TASKS #>
                Write-Host "Scheduled Tasks on Server" -ForegroundColor Red
                $tasks = Invoke-Command -ComputerName "$($environmentsMaster[$x][$y][0][6].Value)$($domain)" -ScriptBlock {
                    Get-ScheduledTask -TaskPath "\" | Select-Object -Property TaskName, LastRunTime | Where-Object TaskName -notlike "Op*" 
                } -Credential $credentials

                $task_array = ""
                foreach ($task in $tasks){
                    Write-Host "Task Name: $($task.TaskName)"
                    $task_array += "$($task.TaskName)`n"
                }

                <# MAINTENANCE DAY #>
                Write-Host "Maintenance Day" -ForegroundColor Red
                $maintenanceDay = $environmentsMaster[$x][$y][0][4].Value
                Write-Host $maintenanceDay
                
                <# DATABASE PROPERTIES #>
                Write-Host "Database Properties" -ForegroundColor Red
                $sqlSession = New-PSSession -ComputerName "$($environmentsMaster[$x][$y][0][6].Value)$($domain)" -Credential $credentials
                    
                    # ALL DATABASES (NAMES AND SIZES in MB)
                    $mainDatabase = ""
                    Write-Host "Listing All Databases and Sizes (in MB)" -ForegroundColor Cyan
                    $all_databases = Invoke-Command -Session $sqlSession -ScriptBlock { 
                        param ($server, $mainDatabase)        
                        Invoke-Sqlcmd -Query "SELECT d.name,
                        ROUND(SUM(mf.size) * 8 / 1024, 0) Size_MB
                        FROM sys.master_files mf
                        INNER JOIN sys.databases d ON d.database_id = mf.database_id
                        WHERE d.database_id > 4 -- Skip system databases
                        GROUP BY d.name
                        ORDER BY d.name" -ServerInstance $server.Name 
                    } -ArgumentList $environmentsMaster[$x][$y][0][6].Value, $mainDatabase
                    foreach ($database in $all_databases) {
                        Write-Host "$($database.name) ---- $($database.Size_MB) MB"
                        if (($database.name -like "*SANDBOX1") -or ($database.name -like "*DEV*")) {
                            if ($database.name -notlike "DM*") {
                                $mainDatabase = $database.name
                            }
                        }
                    }
                    Write-Host "$($mainDatabase) is the main database!`n" -Foregroundcolor green

                    # MAXDOP/THRESHOLD
                    Write-Host "Identifying MaxDOP/Threshold..." -ForegroundColor Cyan
                    $database_maxdop_threshold = Invoke-Command  -Session $sqlSession -ScriptBlock { 
                        param ($server)
                        Invoke-Sqlcmd -Query "SELECT name, value, [description] FROM sys.configurations WHERE name like
                        '%parallel%' ORDER BY name OPTION (RECOMPILE);" -ServerInstance $server.Name
                    } -ArgumentList $environmentsMaster[$x][$y][0][6].Value
                    $maxdop = $database_maxdop_threshold | Where-Object {$_.name -like "cost*"} | Select-Object -property value
                    $cost_threshold = $database_maxdop_threshold | Where-Object {$_.name -like "max*"} | Select-Object -property value
                    Write-Host "Max DOP --- $($maxdop.value) MB"
                    Write-Host "Cost Threshold --- $($cost_threshold.value) MB"
                    
                    # MIN/MAX MEMORY
                    Write-Host "Identifying MIN/MAX Memory..." -ForegroundColor Cyan
                    $database_memory = Invoke-Command -Session $sqlSession -ScriptBlock { 
                        param ($server)
                        Invoke-Sqlcmd -Query "SELECT name, value, [description] FROM sys.configurations WHERE name like
                        '%server memory%' ORDER BY name OPTION (RECOMPILE);" -ServerInstance $server.Name
                    } -ArgumentList $environmentsMaster[$x][$y][0][6].Value 
                    $database_memory_max = $database_memory | where-Object {$_.name -like "max*"} | Select-Object -property value
                    $database_memory_min = $database_memory | where-Object {$_.name -like "min*"} | Select-Object -property value
                    Write-Host "Max Server Memory --- $($database_memory_max.value) MB"
                    Write-Host "Min Server Memory --- $($database_memory_min.value) MB"
                    
                    # DATABASE ENCRYPTION
                    Write-Host "Identifying Database Encryption..." -ForegroundColor Cyan
                    $database_encryption = Invoke-Command -Session $sqlSession -ScriptBlock { 
                        param ($server)
                        Invoke-Sqlcmd -Query "SELECT
                        db.name,
                        db.is_encrypted
                        FROM
                        sys.databases db
                        LEFT OUTER JOIN sys.dm_database_encryption_keys dm
                            ON db.database_id = dm.database_id;
                        GO" -ServerInstance $server 
                    } -ArgumentList $environmentsMaster[$x][$y][0][6].Value
                    $dbEncryption = $database_encryption | Where-Object {$_.name -eq $mainDatabase}
                    Write-Host "$($dbEncryption.name) --- $($dbEncryption.is_encrypted)"

                    # DATABASE SIZE (MAIN)
                    Write-Host "Calculating Database Size" -ForegroundColor Cyan
                    $database_dbSize = Invoke-Command -Session $sqlSession -ScriptBlock { 
                        param ($server, $mainDatabase)        
                        Invoke-Sqlcmd -Query "USE $($mainDatabase)
                        GO
                        exec sp_spaceused
                        GO" -ServerInstance $server.Name 
                    } -ArgumentList $environmentsMaster[$x][$y][0][6].Value, $mainDatabase
                    Write-Host "$($database_dbSize.database_name) --- $($database_dbSize.database_size)"
        
                    # CUSTOM MODELS
                    Write-Host "Calculating Custom Models..." -ForegroundColor Cyan
                    $database_custom_models = Invoke-Command -Session $sqlSession -ScriptBlock { 
                        param ($server, $mainDatabase)        
                        Invoke-Sqlcmd -Query "USE $($mainDatabase);
                        SELECT COUNT(*) FROM ip.olap_properties 
                        WHERE bism_ind ='N' 
                        AND olap_obj_name 
                        NOT like 'PVE%'" -ServerInstance $server.Name 
                    } -ArgumentList $environmentsMaster[$x][$y][0][6].Value, $mainDatabase | Select-Object -property olap_obj_name
                    foreach ($model in $database_custom_models.olap_obj_name) {
                        Write-Host $model
                    }  
                    
                    # INTERFACES
                    Write-Host "Identifying Interfaces..." -ForegroundColor Cyan
                    $database_interfaces = Invoke-Command -Session $sqlSession -ScriptBlock { 
                        param ($server, $mainDatabase)        
                        Invoke-Sqlcmd -Query "USE $($mainDatabase);
                        SELECT
                        s.description JobStreamName,
                        j.description JobName,
                        j.job_order JobOrder,
                        j.job_id JobID,
                        p.name ParamName,
                        p.param_value ParamValue,
                        MIN(r.last_started) JobLastStarted,
                        MAX(r.last_finished) JobLastFinished,
                        MAX(CONVERT(CHAR(8), DATEADD(S,DATEDIFF(S,r.last_started,r.last_finished),'1900-1-1'),8)) Duration
                        FROM ip.job_stream_job j
                        INNER JOIN ip.job_stream s
                        ON j.job_stream_id = s.job_stream_id
                        INNER JOIN ip.job_stream_schedule ss
                        ON ss.job_stream_id = s.job_stream_id
                        INNER JOIN ip.job_run_status r
                        ON s.job_stream_id = r.job_stream_id
                        LEFT JOIN ip.job_param p
                        ON j.job_id = p.job_id
                        WHERE P.Name = 'Command'
                        GROUP BY
                        s.description,
                        j.description,
                        j.job_order,
                        j.job_id,
                        p.name,
                        p.param_value;" -ServerInstance $server.Name 
                    } -ArgumentList $environmentsMaster[$x][$y][0][6].Value, $mainDatabase
                    $database_interfaces.ParamValue  
                    
                    # LICENSE COUNT
                    Write-Host "Calculating License Count..." -ForegroundColor Cyan
                    $database_license_count = Invoke-Command -Session $sqlSession -ScriptBlock { 
                        param ($server, $mainDatabase)        
                        Invoke-Sqlcmd -Query "USE $($mainDatabase);
                        SELECT
                        LicenseRole,
                        COUNT(UserName) UserCount,
                        r.seats LicenseCount
                        FROM (
                        SELECT
                        s1.description LicenseRole,
                        s1.structure_code LicenseCode,
                        u.active_ind Active,
                        u.full_name UserName
                        FROM ip.ip_user u
                        INNER JOIN ip.structure s
                        ON u.role_code = s.structure_code
                        INNER JOIN ip.structure s1
                        ON s.father_code = s1.structure_code
                        WHERE u.active_ind = 'Y'
                        ) q
                        INNER JOIN ip.ip_role r
                        ON q.LicenseCode = r.role_code
                        GROUP BY
                        LicenseRole,
                        LicenseCode,
                        r.seats" -ServerInstance $server.Name 
                    } -ArgumentList $environmentsMaster[$x][$y][0][6].Value, $mainDatabase
                    $licenseProperties = $database_license_count | Select-Object -Property LicenseRole,LicenseCount
                    $totalLicenseCount = 0
                    foreach ($license in $licenseProperties){
                        Write-Output "$($license.LicenseRole): $($license.LicenseCount)"
                        $totalLicenseCount += $license.LicenseCount
                    }
                    Write-Output "Total License Count: $($totalLicenseCount)"
        
                    # PROGRESSING WEB VERSION
                    Write-Host "Identifying Progressing Web Version..." -ForegroundColor Cyan
                    $database_progressing_web_version = Invoke-Command -Session $sqlSession -ScriptBlock { 
                        param ($server, $mainDatabase)        
                        Invoke-Sqlcmd -Query "USE $($mainDatabase); SELECT TOP 1 sub_release 
                        FROM ip.pv_version 
                        WHERE release = 'PROGRESSING_WEB'
                        ORDER BY seq DESC;" -ServerInstance $server.Name 
                    } -ArgumentList $environmentsMaster[$x][$y][0][6].Value, $mainDatabase
                    $database_progressing_web_version.sub_release
        
                <# EXCEL LOGIC AND VARIABLES#>
                $buildData.Cells.Item(12,2)= $environmentsMaster[$x][$y][0][6].Value.Substring(($environmentsMaster[$x][$y][0][6].Value.Length - 2), 2)
                $buildData.Cells.Item(44,3)= $database_dbSize.database_size
                $buildData.Cells.Item(43,3)= $database_memory_max.value
                $buildData.Cells.Item(42,3)= $database_memory_min.value
        
                $buildData.Cells.Item(74,2)= "$($environmentsMaster[$x][$y][0][6].Value)"
                $buildData.Cells.Item(74,3)= "$(@($environmentsMaster[$x][$y][2][2]).Count)"
                $buildData.Cells.Item(74,4)= "$([MATH]::Round((($environmentsMaster[$x][$y][2][1].MaxCapacity) / 1000000),2))"
                $buildData.Cells.Item(74,5)= $hdStringArray
                $buildData.Cells.Item(74,6)= $diskResize
                $buildData.Cells.Item(74,7)= $task_array

                <# AWS-SPECIFIC VARIABLES #>
                $buildData.Cells.Item(74,8)= $instanceSize
                $buildData.Cells.Item(74,9)= $availabilityZone
                $buildData.Cells.Item(74,10)= $ipAddress
                $buildData.Cells.Item(74,11)= $instanceID
                
                $buildData.Cells.Item(26,3)= $database_progressing_web_version.sub_release
        
                $buildData.Cells.Item(28,2)= $database_custom_models.Count
                $modelCount = 0;
                foreach ($model in $database_custom_models.olap_obj_name){
                    $buildData.Cells.Item(92, (2 + $modelCount))= $model
                    $modelCount++
                }
        
                $databaseCount = 0
                foreach ($database in $all_databases) {
                    $buildData.Cells.Item(101, (2 + $databaseCount))= $database.name
                    $buildData.Cells.Item(102, (2 + $databaseCount))= $database.Size_MB
                    $databaseCount++
                }
        
                $buildData.Cells.Item(30,3)= $database_interfaces.ParamValue.Count
                $interfaceCount = 0
                foreach ($interface in $database_interfaces.ParamValue) {
                    $buildData.Cells.Item(96, (2 + $interfaceCount))= $interface
                    $interfaceCount++
                }
        
                $buildData.Cells.Item(41,3)= $dbEncryption.is_encrypted
                $buildData.Cells.Item(22,3)= $totalLicenseCount
                $buildData.Cells.Item(40,3)= $cost_threshold.value
                $buildData.Cells.Item(39,3)= $maxdop.value
                $buildData.Cells.Item(105,3)= $maintenanceDay
        
                Remove-PSSession -Session $sqlSession
        
                Write-Host "`n" -ForegroundColor Red  
            }
        
            #######################
            # SANDBOX PVE SERVER 
            #######################
            elseif ($environmentsMaster[$x][$y][0][6].Value.Substring(($environmentsMaster[$x][$y][0][6].Value.Length - 5), 3) -eq "pve") {
                Write-Host "THIS IS THE SANDBOX PVE SERVER" -ForegroundColor Cyan
        
                <# CPU/RAM #>
                Write-Host "Server CPU and RAM" -ForegroundColor Red
                Write-Host "Server Name: $($environmentsMaster[$x][$y][0][6].Value)"
                Write-Host "Server CPUs: $(@($environmentsMaster[$x][$y][2][2]).Count)"
                Write-Host "Server RAM: $([MATH]::Round((($environmentsMaster[$x][$y][2][1].MaxCapacity) / 1000000),2))"  
        
                <# HARDDRIVES #>
                Write-Host "Disks and Disk Capacity" -ForegroundColor Red
                $diskResize = "Yes"
                $hdStringArray = ""
                foreach ($hd in $environmentsMaster[$x][$y][2][0]) {
                    $hdString = "$($hd.DeviceID) $([MATH]::Round($hd.Size / 1GB,2))gb"
                    $hdStringArray += "$($hdString)`n"
                    Write-Host $hdString  
                    if (([MATH]::Round($hd.Size / 1GB,2)) -gt 60) {
                        $diskResize = "No"  
                    }
                }
                Write-Host "Standard Size Disks (less than 60GB): $($diskResize)"
        
                <# AWS METADATA #>

                        # INSTANCE SIZE #
                        Write-Host "Instance Size" -ForegroundColor Red
                        $instanceSize = $environmentsMaster[$x][$y][1][2]
                        Write-Host $instanceSize
                        
                        # AVAILABILITY ZONE #
                        Write-Host "Availability Zone" -ForegroundColor Red
                        $availabilityZone = $environmentsMaster[$x][$y][1][3]
                        Write-Host $availabilityZone

                        # IP ADDRESS #
                        Write-Host "IP Address" -ForegroundColor Red
                        $ipAddress = $environmentsMaster[$x][$y][1][4]
                        Write-Host $ipAddress

                        # INSTANCE ID #
                        Write-Host "Instance ID" -ForegroundColor Red
                        $instanceID = $environmentsMaster[$x][$y][1][0]
                        Write-Host $instanceID
        
                <# SCHEDULED TASKS #>
                Write-Host "Scheduled Tasks on Server" -ForegroundColor Red
                $tasks = Invoke-Command -ComputerName "$($environmentsMaster[$x][$y][0][6].Value)$($domain)" -ScriptBlock {
                    Get-ScheduledTask -TaskPath "\" | Select-Object -Property TaskName, LastRunTime | Where-Object TaskName -notlike "Op*" 
                } -Credential $credentials

                $task_array = ""
                foreach ($task in $tasks){
                    Write-Host "Task Name: $($task.TaskName)"
                    $task_array += "$($task.TaskName)`n"
                }
        
                <# CURRENT VERSION #>
                Write-Host "Current Environment Version" -ForegroundColor Red
                $crVersion = Invoke-Command -ComputerName "$($environmentsMaster[$x][$y][0][6].Value)$($domain)" -Credential $credentials -ScriptBlock {
                    Get-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\Planview\WebServerPlatform"
                }
                Write-Host $crVersion.CrVersion
        
                <# MAJOR VERSION #>
                Write-Host "Major Version" -ForegroundColor Red
                $majorVersion = $crVersion.CrVersion.Split('.')[0]
                $majorVersion
                
                <# OPEN SUITE #>
                Write-Host "OpenSuite" -ForegroundColor Red
                $opensuite = Invoke-Command -ComputerName "$($environmentsMaster[$x][$y][0][6].Value)$($domain)" -Credential $credentials -ScriptBlock {
                    if ((Test-Path -Path "C:\ProgramData\Actian" -PathType Container) -And (Test-Path -Path "F:\Planview\Interfaces\OpenSuite" -PathType Container)) {
        
                        $software = "*Actian*";
                        $installed = (Get-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*" | Where { $_.DisplayName -like $software }) -ne $null
        
                        if ($installed) {
                            return "Yes"
                        }
                        
                    } else {
                        return "No"
                    }
                }
                Write-Host "OpenSuite Detected: $($opensuite)"
        
                <# CONNECTORS #>
                Write-Host "Connectors" -ForegroundColor Red
                $PRMAdapter = "False"
                $PPAdapter = "False"
                $LKAdapter = "False"

                $PRMini = Invoke-Command -ComputerName "$($environmentsMaster[$x][$y][0][6].Value)$($domain)" -Credential $credentials -ScriptBlock {

                    if (Test-Path -Path "F:\Planview\midtier\webserver\objects\PRM_Adapter_Config.ini" -PathType leaf){

                        $PRMiniContent = (Get-Content -Path "F:\Planview\midtier\webserver\objects\PRM_Adapter_Config.ini" | Where-Object {$_.length -ne 0}) | Where-Object {$_.substring(0,1) -ne ";"}

                        $masterArray = @()
                        $databaseName = ""
                        foreach($x in $PRMiniContent){

                            $valuePair = New-Object PSObject
                            Add-Member -InputObject $valuePair -MemberType NoteProperty -Name Database -Value ""
                            Add-Member -InputObject $valuePair -MemberType NoteProperty -Name Key -Value ""

                            if ($x -match "^\["){

                                $databaseName = $x.substring(1, ($x.length - 2)) 

                                $valuePair.Database = $databaseName
                                $valuePair.Key = $databaseName
                                $masterArray += $valuePair
                            
                            
                            } else {       
                            
                                $valuePair.Database = $databaseName
                                $valuePair.Key = $x 
                                $masterArray += $valuePair
                            
                            }           
                        }
                    
                        return $masterArray
                    
                    } else {

                        return 0

                    }
                }

                if ($PRMini -ne '0'){

                    $PRMAdapter = "True"

                    Write-Host "PRM Adapter Found"
                    $LKKey = ($PRMini | Where-Object {$_.Database -like "*SANDBOX*"} | Where-Object {$_.Key -like "use_prm_leankit*"}).Key
                    $PPKey = ($PRMini | Where-Object {$_.Database -like "*SANDBOX*"} | Where-Object {$_.Key -like "use_prm_projectplace*"}).Key

                    if ($LKKey -like "*true*"){
                        $LKAdapter = "True"
                        Write-Host "LeanKit Connector --- True"
                    } else {
                        Write-Host "LeanKit Connector --- False"
                    }

                    if ($PPKey -like "*true*"){
                        $PPAdapter = "True"
                        Write-Host "ProjectPlace Connector --- True"
                    } else {
                        Write-Host "ProjectPlace Connector --- False"
                    }

                } else {

                    Write-Host "PRM Adapter not found"

                    $LegacyPPAdapter = Invoke-Command -ComputerName "$($environmentsMaster[$x][$y][0][6].Value)$($domain)" -Credential $credentials -ScriptBlock {
                        Test-Path -Path "F:\Planview\midtier\webserver\objects\ProjectPlace_Config.ini" -PathType leaf
                    }

                    $PPAdapter = $LegacyPPAdapter
                    Write-Host "ProjectPlace (Legacy install) --- $($PPAdapter)"

                }
        
                # NEW RELIC #
                Write-Host "New Relic" -ForegroundColor Red
                $newRelic = Invoke-Command -ComputerName "$($environmentsMaster[$x][$y][0][6].Value)$($domain)" -Credential $credentials -ScriptBlock {
                    if (Test-Path -Path "C:\ProgramData\New Relic" -PathType Container ) {
                        Write-Host "New Relic has been detected on this server"
                        return "Yes"
                    } else {
                        Write-Host "New Relic was not detected on this server"
                        return "No"
                    }
                }
        
                    # GET WEB CONFIG #
                    $webConfig = Invoke-Command -ComputerName "$($environmentsMaster[$x][$y][0][6].Value)$($domain)" -Credential $credentials -ScriptBlock {
                        return Get-Content -Path "F:\Planview\MidTier\ODataService\Web.config"
                    }
                    $webConfig = [xml] $webConfig
        
                <# SANDBOX URL #>
                Write-Host "Sandbox URL" -ForegroundColor Red
                $environmentURL = $webConfig.configuration.appSettings.add | Where-Object {$_.key -eq "PveUrl"} | Select-Object -Property value
                Write-Host $environmentURL.value
        
                <# DNS ALIAS #>
                Write-Host "Sandbox DNS Alias" -ForegroundColor Red            
                $dnsAlias = ($environmentURL.value.Split('//')[2]).Split('.')[0] 
                $dnsAlias
        
                <# REPORT FARM URL #>
                Write-Host "Report Farm URL" -ForegroundColor Red
                $reportfarmURL = $webConfig.configuration.appSettings.add | Where-Object {$_.key -eq "Report_Server_Web_Service_URL"} | Select-Object -Property value
                Write-Host $reportfarmURL.value
        
                <# IP RESTRICTIONS #>
                Write-Host "IP Restrictions on F5" -ForegroundColor Red
                Write-Host "Currently Not Automated"
                $IPRestrictions = "Not Automated"
<#                    
                    # Authentication on the F5 #
                    $websession =  New-Object Microsoft.PowerShell.Commands.WebRequestSession
                    $jsonbody = @{username = $f5Credentials.UserName ; password = $f5Credentials.GetNetworkCredential().Password; loginProviderName='tmos'} | ConvertTo-Json
                    $authResponse = Invoke-RestMethodOverride -Method Post -Uri "https://$($f5ip)/mgmt/shared/authn/login" -Credential $f5Credentials -Body $jsonbody -ContentType 'application/json'
                    $token = $authResponse.token.token
                    $websession.Headers.Add('X-F5-Auth-Token', $Token)
        
                    # Calling data-group REST endpoint and parsing IPRestrictions list #
                    $IPRestrictionsList = (Invoke-RestMethod  -Uri "https://$($f5ip)/mgmt/tm/ltm/data-group/internal" -WebSession $websession).Items | 
                        Where-Object {$_.name -eq "IPRestrictions"} | Select-Object -Property records
        
                    foreach ($record in $IPRestrictionsList.records) {
                        if ($record.name -eq "$($dnsAlias).pvcloud.com") {
                            $IPRestrictions = "Yes"
                            Write-Host "IP restrctions were found for $($dnsAlias).pvcloud.com"
                        }
                    }
        
                    if ($IPRestrictions -eq "No") {
                        Write-Host "No IP restrictions found for $($dnsAlias).pvcloud.com"
                    }
#>                
                <# EXCEL LOGIC AND VARIABLES#>
                $webServerCount++
                $buildData.Cells.Item(24,3)= $webServerCount
                $buildData.Cells.Item(2,2)= $environmentURL.value   
                $buildData.Cells.Item(8,2)= $dnsAlias
                $buildData.Cells.Item(35,3)= $IPRestrictions
                $buildData.Cells.Item(36,3)= $opensuite
                $buildData.Cells.Item(25,3)= $crVersion.CrVersion
                $buildData.Cells.Item(19,2)= "True"
        
                $buildData.Cells.Item(76,2)= "$($environmentsMaster[$x][$y][0][6].Value)"
                $buildData.Cells.Item(76,3)= "$(@($environmentsMaster[$x][$y][2][2]).Count)"
                $buildData.Cells.Item(76,4)= "$([MATH]::Round((($environmentsMaster[$x][$y][2][1].MaxCapacity) / 1000000),2))"
                $buildData.Cells.Item(76,5)= $hdStringArray
                $buildData.Cells.Item(76,6)= $diskResize
                $buildData.Cells.Item(76,7)= $task_array

                <# AWS-SPECIFIC VARIABLES #>
                $buildData.Cells.Item(76,8)= $instanceSize
                $buildData.Cells.Item(76,9)= $availabilityZone
                $buildData.Cells.Item(76,10)= $ipAddress
                $buildData.Cells.Item(76,11)= $instanceID
        
                $buildData.Cells.Item(31,3)= $PPAdapter
                $buildData.Cells.Item(32,3)= $LKAdapter
                $buildData.Cells.Item(37,3)= $PRMAdapter
        
                Write-Host "`n" -ForegroundColor Red  
            }
            
            
        }
        
    }

} 

<#
for ($x=0; $x -lt $environmentsMaster.Length; $x++) {
    Write-Host "X changed to $($x)" -Foregroundcolor red
    for ($y=0; $y -lt $environmentsMaster[$x].Length; $y++) {        
        Write-Host "Y changed to $($y)" -Foregroundcolor green
        for ($i=0; $i -lt $environmentsMaster[$x][$y].Length; $i++) {
            Write-Host "i changed to $($i)" -Foregroundcolor yellow
            $environmentsMaster[$x][$y][$i]
        }

    }
} 
#>

<# SAVE AND CLOSE EXCEL #>
$excelfile.Save()
$excelfile.Close()