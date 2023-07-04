$VerbosePreference = "SilentlyContinue"
$InformationPreference = "Continue"
$WarningPreference = "Continue"

# variables configured in form
$username = $form.gridmailuser.userPrincipalName
$devicesToActivate = $form.devicelist.rightToLeft
$devicesToBlock = $form.devicelist.leftToRight

try {
    <#----- Exchange On-Premises: Start -----#>
    # Connect to Exchange
    try {
        $adminSecurePassword = ConvertTo-SecureString -String "$ExchangeAdminPassword" -AsPlainText -Force
        $adminCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ExchangeAdminUsername, $adminSecurePassword
        $sessionOption = New-PSSessionOption -SkipCACheck -SkipCNCheck -SkipRevocationCheck
        $exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $exchangeConnectionUri -Credential $adminCredential -SessionOption $sessionOption -Authentication $ExchangeAuthentication -ErrorAction Stop 
        #-AllowRedirection
        $session = Import-PSSession $exchangeSession -DisableNameChecking -AllowClobber

        Write-Information "Successfully connected to Exchange using the URI [$exchangeConnectionUri]" 
    
        $Log = @{
            Action            = "UpdateResource" # optional. ENUM (undefined = default) 
            System            = "Exchange On-Premise" # optional (free format text) 
            Message           = "Successfully connected to Exchange using the URI [$exchangeConnectionUri]" # required (free format text) 
            IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
            TargetDisplayName = $exchangeConnectionUri # optional (free format text) 
            TargetIdentifier  = $([string]$session.GUID) # optional (free format text) 
        }
        #send result back  
        Write-Information -Tags "Audit" -MessageData $log
    }
    catch {
        Write-Error "Error connecting to Exchange using the URI [$exchangeConnectionUri]. Error: $($_.Exception.Message)"
        $Log = @{
            Action            = "UpdateResource" # optional. ENUM (undefined = default) 
            System            = "Exchange On-Premise" # optional (free format text) 
            Message           = "Failed to connect to Exchange using the URI [$exchangeConnectionUri]." # required (free format text) 
            IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
            TargetDisplayName = $exchangeConnectionUri # optional (free format text) 
            TargetIdentifier  = $([string]$session.GUID) # optional (free format text) 
        }
        #send result back  
        Write-Information -Tags "Audit" -MessageData $log
    }

    if ($devicesToActivate.count -gt 0) {
        try {
            Write-Information "Starting to allow device [$($devicesToActivate.FriendlyName)] for user [$username)]"
            
            foreach ($device in $devicesToActivate) {
                try {
                    Set-CASMailbox -Identity $username -ActiveSyncAllowedDeviceIDs @{ add = $device.DeviceId }

                    Write-Information "Finished allowing $($device.DeviceId) for user [$username]"
                    $Log = @{
                        Action            = "UpdateResource" # optional. ENUM (undefined = default) 
                        System            = "Exchange On-Premise" # optional (free format text) 
                        Message           = "Successfully allowing $($device.DeviceId) for user [$username]" # required (free format text) 
                        IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
                        TargetDisplayName = $username # optional (free format text) 
                        TargetIdentifier  = $($device.DeviceId) # optional (free format text) 
                    }
                    #send result back  
                    Write-Information -Tags "Audit" -MessageData $log       
                }
                catch {
                    Write-Error "Error activating $($device.DeviceId) for user [$username]. Error: $($_.Exception.Message)" 
                    $Log = @{
                        Action            = "UpdateResource" # optional. ENUM (undefined = default) 
                        System            = "Exchange On-Premise" # optional (free format text) 
                        Message           = "Failed to allow [$($device.DeviceId)] for [$username]" # required (free format text) 
                        IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
                        TargetDisplayName = $username # optional (free format text) 
                        TargetIdentifier  = $($device.DeviceId) # optional (free format text) 
                    }
                    #send result back  
                    Write-Information -Tags "Audit" -MessageData $log                    
                }
            }
        }               
        catch {
            Write-Error "Could not allow [$($devicesToActivate.FriendlyName)] for user [$username]. Error: $($_.Exception.Message)"
            $Log = @{
                Action            = "UpdateResource" # optional. ENUM (undefined = default) 
                System            = "Exchange On-Premise" # optional (free format text) 
                Message           = "Failed to allow [$($devicesToActivate.FriendlyName)] for user [$username]" # required (free format text) 
                IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
                TargetDisplayName = $username # optional (free format text) 
                TargetIdentifier  = $($devicesToActivate.DeviceId) # optional (free format text) 
            }
            #send result back  
            Write-Information -Tags "Audit" -MessageData $log            
        }
    }

    if ($devicesToBlock.count -gt 0) {
        try {
            Write-Information "Starting to block device [$($devicesToBlock.FriendlyName)] for user [$username)]"
            
            foreach ($device in $devicesToBlock) {
                try {
                    Set-CASMailbox -Identity $username -ActiveSyncBlockedDeviceIDs  @{ add = $device.DeviceId }

                    Write-Information "Finished blocking $($device.DeviceId) for user [$username]"
                    $Log = @{
                        Action            = "UpdateResource" # optional. ENUM (undefined = default) 
                        System            = "Exchange On-Premise" # optional (free format text) 
                        Message           = "Successfully blocking $($device.DeviceId) for user [$username]" # required (free format text) 
                        IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
                        TargetDisplayName = $username # optional (free format text) 
                        TargetIdentifier  = $($device.DeviceId) # optional (free format text) 
                    }
                    #send result back  
                    Write-Information -Tags "Audit" -MessageData $log       
                }
                catch {
                    Write-Error "Error blocking $($device.DeviceId) for user [$username]. Error: $($_.Exception.Message)" 
                    $Log = @{
                        Action            = "UpdateResource" # optional. ENUM (undefined = default) 
                        System            = "Exchange On-Premise" # optional (free format text) 
                        Message           = "Failed to block [$($device.DeviceId)] for [$username]" # required (free format text) 
                        IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
                        TargetDisplayName = $username # optional (free format text) 
                        TargetIdentifier  = $($device.DeviceId) # optional (free format text) 
                    }
                    #send result back  
                    Write-Information -Tags "Audit" -MessageData $log                    
                }
            }
        }               
        catch {
            Write-Error "Could not block [$($devicesToBlock.FriendlyName)] for user [$username]. Error: $($_.Exception.Message)"
            $Log = @{
                Action            = "UpdateResource" # optional. ENUM (undefined = default) 
                System            = "Exchange On-Premise" # optional (free format text) 
                Message           = "Failed to blocke [$($devicesToBlock.FriendlyName)] for user [$username]" # required (free format text) 
                IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
                TargetDisplayName = $username # optional (free format text) 
                TargetIdentifier  = $($devicesToBlock.DeviceId) # optional (free format text) 
            }
            #send result back  
            Write-Information -Tags "Audit" -MessageData $log            
        }
    }        
} catch {
    Write-Error "Could not manage devices for user [$username]. Error: $($_.Exception.Message)"    
    $Log = @{
        Action            = "UpdateResource" # optional. ENUM (undefined = default) 
        System            = "Exchange On-Premise" # optional (free format text) 
        Message           = "Failed to manage devices for user [$username]." # required (free format text) 
        IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
        TargetDisplayName = $username # optional (free format text) 
        TargetIdentifier  = $username # optional (free format text) 
    }
    #send result back  
    Write-Information -Tags "Audit" -MessageData $log
}
finally {
    # Disconnect from Exchange
    try {
        Remove-PsSession -Session $exchangeSession -Confirm:$false -ErrorAction Stop
        Write-Information "Successfully disconnected from Exchange using the URI [$exchangeConnectionUri]"     
        $Log = @{
            Action            = "UpdateResource" # optional. ENUM (undefined = default) 
            System            = "Exchange On-Premise" # optional (free format text) 
            Message           = "Successfully disconnected from Exchange using the URI [$exchangeConnectionUri]" # required (free format text) 
            IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
            TargetDisplayName = $exchangeConnectionUri # optional (free format text) 
            TargetIdentifier  = $([string]$session.GUID) # optional (free format text) 
        }
        #send result back  
        Write-Information -Tags "Audit" -MessageData $log
    }
    catch {
        Write-Error "Error disconnecting from Exchange.  Error: $($_.Exception.Message)"
        $Log = @{
            Action            = "UpdateResource" # optional. ENUM (undefined = default) 
            System            = "Exchange On-Premise" # optional (free format text) 
            Message           = "Failed to disconnect from Exchange using the URI [$exchangeConnectionUri]." # required (free format text) 
            IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
            TargetDisplayName = $exchangeConnectionUri # optional (free format text) 
            TargetIdentifier  = $([string]$session.GUID) # optional (free format text) 
        }
        #send result back  
        Write-Information -Tags "Audit" -MessageData $log 
    }
    <#----- Exchange On-Premises: End -----#>
}


