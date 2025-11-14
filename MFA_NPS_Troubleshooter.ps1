Clear-Host


$title = @"

****            Welcome to MFA NPS Extension Troubleshooter Tool                ****

**** This Tool will help you to troubleshoot MFA NPS Extension Knows issues     ****
**** Tool Version is 3.5, Make Sure to Visit MS site to get the latest version  ****
**** Thank you for Using MS Products, Microsoft @2025                           ****
"@

$Choice_Options = @"

    (1) Isolate the Cause of the issue: if it's NPS or MFA issue (Export MFA RegKeys, Restart NPS, Test, Import Regkeys, Restart NPS)... 

    (2) All users not able to use MFA NPS Extension (Testing Access to Azure/Create HTML Report) ... 

    (3) Specific User not able to use MFA NPS Extension (Test MFA for specific UPN) ... 
    
    (4) Collect Logs to contact MS support (Enable Logging/Restart NPS/Gather Logs)... 
    
    (E) EXIT SCRIPT

"@


$Cloud_Choice = @"
    (C) Azure Commercial 

    (G) Azure Government 

    (V) Microsoft Azure operated by 21Vianet 

    (E) EXIT SCRIPT

"@


# This function will display messages with consistent styling accross the script
function Write-Message {
    <#
    .DESCRIPTION
    Displays messages with consistent styling and optional inline output.

    .PARAMETER Message
    The text to display.

    .PARAMETER Type
    Defines the message style: Info, Success, Progress, Error, Menu, Section, or Default.

    .PARAMETER BackgroundColor
    Optional background color override.

    .PARAMETER NoNewLine
    Prints message inline without adding a trailing newline.
    #>

    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [ValidateSet('Info', 'Success', 'Progress', 'Error', 'Menu', 'Default')]
        [string]$Type = 'Default',

        [string]$BackgroundColor,

        [switch]$NoNewLine
    )

    # Define color based on message type
    $fg = switch ($Type.ToLower()) {
        'info' { 'Yellow' }
        'success' { 'Green' }
        'progress' { 'Cyan' }
        'error' { 'Red' }
        'Menu' { 'Green' }
        default { 'White' }
    }

    # Add optional prefix based on type
    switch ($Type.ToLower()) {
        'info' { $Message = "[INFO] $Message" }
        'success' { $Message = "[OK] $Message" }
        'error' { $Message = "[ERROR] $Message" }
        'progress' { $Message = "`n[*] $Message" }
    }

    # Add leading newline only if not inline
    $prefix = if (-not $NoNewLine) { "`n" } else { "" }

    # Prepare parameters
    $params = @{
        Object          = "$prefix$Message"
        ForegroundColor = $fg
    }
    if ($BackgroundColor) { $params.BackgroundColor = $BackgroundColor }
    if ($NoNewLine) { $params.NoNewline = $true }

    Write-Host @params
}

##### This function evaluates the need to install MS Graph libraries #####
##### Microsoft 2024 @Miguel Ferreira #####

##### Optimize Library script and adding import to the else if installation successful

Function Manage_Script_Libraries {

    Write-Message "Ensuring required Microsoft Graph modules are installed..." -Type Progress

    # List of required Graph modules
    $RequiredModules = @(
        "Microsoft.Graph.Authentication",
        "Microsoft.Graph.Applications",
        "Microsoft.Graph.Users",
        "Microsoft.Graph.Identity.DirectoryManagement",
        "Microsoft.Graph.Identity.SignIns"
    )

    # Check and install each module
    foreach ($moduleName in $RequiredModules) {
        try {
            if (Get-Module -ListAvailable -Name $moduleName) {
                Write-Message "$moduleName module available" -Type Info
                Import-Module -Name $moduleName -ErrorAction SilentlyContinue
            }
            else {
                Write-Message "Installing $moduleName module..." -Type Progress
                Install-Module -Name $moduleName -Force -ErrorAction Stop
                Import-Module -Name $moduleName -ErrorAction SilentlyContinue
                Write-Message "$moduleName module installed and imported successfully" -Type Success
            }
        }
        catch {
            Write-Message "Failed to install or import $moduleName : $_" -Type Error
        }
    }
}


# centralize all connections here

function Connect-MgGraphEndpoint {

    # Parameter help description
    param(
        [string]$CloudEnvironment
    )
    
    # Validate libraries for ms graph
    Manage_Script_Libraries
    
    # Select cloud environment where GA admin account will sign-in
    Write-Message "Start Entra connection to be established with Global Admin role ..."  -Type progress
        
    if ($CloudEnvironment -eq 'C') { 
        
        Connect-MgGraph -Scopes Domain.Read.All, Application.Read.All -NoWelcome -Environment Global

    }

    if ($CloudEnvironment -eq 'G') { 

        Connect-MgGraph -Scopes Domain.Read.All, Application.Read.All -NoWelcome -Environment USGov

    }

    if ($CloudEnvironment -eq 'V') { 

        Connect-MgGraph -Scopes Domain.Read.All, Application.Read.All -NoWelcome -Environment China

    }
}



################################################################################################


Write-Message "*******************************************************************************************"
Write-Message $title 
Write-Message "*******************************************************************************************"

# Activity or test option
Write-Message "Please Choose one of the tests below: " 
Write-Message $Choice_Options -Type Menu

$Timestamp = "$((Get-Date).ToString("yyyyMMdd_HHmmss"))"
$Choice_Number = ''
$Choice_Number = Read-Host -Prompt "`nBased on which test you need to run, please type 1, 2, 3, 4 or E to exit the test. Then press Enter " 

while ( !($Choice_Number -in @('1', '2', '3', '4', 'E'))) {

    $Choice_Number = Read-Host -Prompt "`nInvalid Option, Based on which test you need to run, please type 1, 2, 3, 4 or E to exit the test. Then press Enter " 

}

# Cloud selection based on the test choice
If ( $Choice_Number.Trim() -in @('2', '3')) { 
    # Decide which cloud environment to be checked
    Write-Message "Please choose one of the cloud environments (Azure Commercial / Azure Government / Microsoft Azure operated by 21Vianet) to evaluate endpoint connectivity and sign-in as GA: " 
    Write-Message $Cloud_Choice -Type Menu

    $Cloud_Choice_Number = ''
    $Cloud_Choice_Number = Read-Host -Prompt "`nBased on which cloud environment you need to evaluate, please type C, G, V or E to exit the test. Then press Enter "

    while ( !($Cloud_Choice_Number.Trim() -in @('C', 'G', 'V', 'E'))) {

        $Cloud_Choice_Number = Read-Host -Prompt "`nInvalid Option, Based on which cloud environment you need to evaluate endpoint connectivity and sign-in as GA,  please type C, G, V or E to exit the test. Then press Enter " 
    }


    # Exit of the cloup Environment is E Eld proceed to check Mg libraries and connect
    if ($Cloud_Choice_Number -eq 'E') {
        Break
    }
    else {
        #Check Required script modules libraries and connect
        Connect-MgGraphEndpoint -CloudEnvironment $Cloud_Choice_Number
        $Global:verifyConnection = Get-MgDomain -ErrorAction SilentlyContinue # This will check if the connection succeeded or not
        if (!$Global:verifyConnection ) {
            Write-Message "Entra connection could not be established. Please verify the connectivity and try again." -Type Error
            Break
        }
    }
}


##### This Function will be run against against MFA NPS Server ######
##### Microsoft 2022 @Ahmad Yasin, Nate Harris (nathar), Will Aftring (wiaftin) ##########

Function Check_Nps_Server_Module {
    param (
        [string]$Cloud_Choice_Number
    )

    # Select cloud environment where GA admin account will sign-in
    # if ($Cloud_Choice_Number -eq 'C') { 
        
    #     Connect-MgGraph -Scopes Domain.Read.All, Application.Read.All -NoWelcome -Environment Global

    # }

    # if ($Cloud_Choice_Number -eq 'G') { 

    #     Connect-MgGraph -Scopes Domain.Read.All, Application.Read.All -NoWelcome -Environment USGov

    # }

    # if ($Cloud_Choice_Number -eq 'V') { 

    #     Connect-MgGraph -Scopes Domain.Read.All, Application.Read.All -NoWelcome -Environment China

    # }

    # Variables 
    $TestStepNumber = 0
    $ErrorActionPreference = 'silentlycontinue'
    $loginAccessResult = 'NA'
    $NotificationaccessResult = 'NA'
    $MFATestVersion = 'NA'
    $MFAVersion = 'NA'
    $NPSServiceStatus = 'NA'
    $SPNExist = 'NA'
    $SPNEnabled = 'NA'
    $FirstSetofReg = 'NA'
    $SecondSetofReg = 'NA'
    $certificateResult = 'NA'
    $ValidCertThumbprint = 'NA'
    $ValidCertThumbprintExpireSoon = 'NA'
    $TimeResult = 'NA'
    $updateResult = 'NA'
    $ListofMissingUpdates = 'NA'
    $objects = @()
    
    ## Variables for endpoints network connectivity
    $TCPLogin = $False
    $TCPAdnotification = $False
    $TCPStrongAuthService = $False
    $TCPCredentials = $False

    $DNSLogin = $False
    $DNSADNotification = $False
    $DNSStrongAuthService = $False
    $DNSCredentials = $False

    $IWRLogin = ""
    $IWRADNotification = ""
    $IWRStrongAuthService = ""
    $IWRCredentials = ""

    
    Write-Message "Validation Entra connection  ..." -Type Info
    $verifyConnection = Get-MgDomain -ErrorAction SilentlyContinue

    if ($null -ne $verifyConnection) {

        Write-Message "Connection established Successfully - Starting the Health Check Process ..." -Type Success

        # Check the accessibility to Azure endpoints based on cloud selection

        if ($Cloud_Choice_Number -eq 'C') { 
        
            $AzureEndpointLogin = "login.microsoftonline.com"
            $AzureEndpointADNotification = "adnotifications.windowsazure.com"
            $AzureEndpointStrongAuthService = "strongauthenticationservice.auth.microsoft.com"

        }

        if ($Cloud_Choice_Number -eq 'G') { 

            $AzureEndpointLogin = "login.microsoftonline.us"
            $AzureEndpointADNotification = "adnotifications.windowsazure.us"
            $AzureEndpointStrongAuthService = "strongauthenticationservice.auth.microsoft.us"

        }

        if ($Cloud_Choice_Number -eq 'V') { 

            $AzureEndpointLogin = "login.chinacloudapi.cn"
            $AzureEndpointADNotification = "adnotifications.windowsazure.cn"
            $AzureEndpointStrongAuthService = "strongauthenticationservice.auth.microsoft.cn"

        }

        # Azure login endpoint
        $AzureEndpointLoginScriptBlock = "Test-NetConnection -ComputerName " + $AzureEndpointLogin + " -Port 443"
        $AzureEndpointLoginURI = "https://" + $AzureEndpointLogin
        $AzureEndpointLoginURISlash = $AzureEndpointLoginURI + "/"

        # Azure notifications endpoint
        $AzureEndpointADNotificationScriptBlock = "Test-NetConnection -ComputerName " + $AzureEndpointADNotification + " -Port 443"
        $AzureEndpointADNotificationURI = "https://" + $AzureEndpointADNotification

        # Azure strong auth service endpoint
        $AzureEndpointStrongAuthServiceScriptBlock = "Test-NetConnection -ComputerName " + $AzureEndpointStrongAuthService + " -Port 443"
        $AzureEndpointStrongAuthServiceURI = "https://" + $AzureEndpointStrongAuthService

        # Azure credentials endpoint
        $AzureEndpointCredentials = "credentials.azure.com"
        $AzureEndpointCredentialsScriptBlock = "Test-NetConnection -ComputerName " + $AzureEndpointCredentials + " -Port 443"
        $AzureEndpointCredentialsURI = "https://" + $AzureEndpointCredentials


        #Muath Updates:
        ####
        function RunPSScript([String] $PSScript) {

            $GUID = [guid]::NewGuid().Guid

            $Job = Register-ScheduledJob -Name $GUID -ScheduledJobOption (New-ScheduledJobOption -RunElevated) -ScriptBlock ([ScriptBlock]::Create($PSScript)) -ArgumentList ($PSScript) -ErrorAction Stop

            $Task = Register-ScheduledTask -TaskName $GUID -Action (New-ScheduledTaskAction -Execute $Job.PSExecutionPath -Argument $Job.PSExecutionArgs) -Principal (New-ScheduledTaskPrincipal -UserID "NT AUTHORITY\SYSTEM" -LogonType ServiceAccount -RunLevel Highest) -ErrorAction Stop

            $Task | Start-ScheduledTask -AsJob -ErrorAction Stop | Wait-Job | Remove-Job -Force -Confirm:$False

            While (($Task | Get-ScheduledTaskInfo).LastTaskResult -eq 267009) { Start-Sleep -Milliseconds 150 }

            $Job1 = Get-Job -Name $GUID -ErrorAction SilentlyContinue | Wait-Job
            $Job1 | Receive-Job -Wait -AutoRemoveJob 

            Unregister-ScheduledJob -Id $Job.Id -Force -Confirm:$False

            Unregister-ScheduledTask -TaskName $GUID -Confirm:$false
        } 


        
        ###### Alternative CONNECTIVITY TEST FUNCTIONS for environments where ICMP is blocked and if Register-ScheduledTask task fails  ######

        #TCP connection based on System.Net library
        function Test-TcpConnection {
            param(
                [parameter(Mandatory = $true)]
                [string]$ComputerName,
                [int]$Port = 443,
                [int]$TimeoutMs = 5000
            )

            try {
                $tcpClient = New-Object System.Net.Sockets.TcpClient
                $connect = $tcpClient.BeginConnect($ComputerName, $Port, $null, $null)
                $wait = $connect.AsyncWaitHandle.WaitOne($TimeoutMs, $false)
        
                if ($wait) {
                    try {
                        $tcpClient.EndConnect($connect)
                        $tcpClient.Close()
                        return $true
                    }
                    catch {
                        $tcpClient.Close()
                        return $false
                    }
                }
                else {
                    $tcpClient.Close()
                    return $false
                }
            }
            catch {
                return $false
            }
        }

        #DNS resolution based on using Resolve-DNSName command
        function Test-DnsResolution {
            param(
                [parameter(Mandatory = $true)]
                [string]$Hostname
            )
    
            try {
                $dnsResult = Resolve-DnsName -Name $Hostname -ErrorAction Stop
                if ($dnsResult) {
                    return $true
                }
                return $false
            }
            catch {
                return $false
            }
        }


        # This function handles the connection invoke-request host names like https://adnotifications.windowsazure.com graciously
        function Get-WebRequestStatusCode {
            <#
            
            Performs a lightweight GET request and reports success even when the endpoint
            intentionally returns Forbidden or "Method Not Found" responses. By design, it does not return 403/404 as failure when checking StatusCode.

            Azure/MFA endpoints such as:
                - https://adnotifications.windowsazure.com
                - https://strongauthenticationservice.auth.microsoft.com
                - https://credentials.azure.com
            return security-by-design responses:
                * 403 Forbidden ("You do not have permission…")
                * 404 Not Found ("API Method Not Found.")
                StatusCode              : Forbidden
                StatusDescription       : Forbidden
                ProtocolVersion         : 1.1
                ResponseUri             : https://adnotifications.windowsazure.com/
                
                These responses indicate the endpoint **is reachable**, not broken.

            #>

            param (
                [Parameter(Mandatory)]
                [string]$Url
            )

            try {
                # If the GET succeeds unexpectedly, treat as reachable
                Invoke-WebRequest -Uri $Url -Method Get -UseBasicParsing -TimeoutSec 10 | Out-Null
                return $true
            }
            catch {
                $errorMessage = $_.ErrorDetails.Message
                $statusCode = $_.Exception.Response.StatusCode

                # Expected Azure-secured responses indicating reachable endpoint
                $expectedMessages = @(
                    "You do not have permission to view this directory or page.",
                    "API method not found.",
                    "API Method Not Found."
                )

                if ( $errorMessage -in $expectedMessages) {
                    return $true
                }

                # Any other error = unreachable or network path broken
                return $false
            }
        }


        ####

        $TestStepNumber = $TestStepNumber + 1
        Write-Message "$TestStepNumber - Checking Accessibility to $AzureEndpointLoginURI ..." -Type info

        ########################################################################
        # TCP Test
        try {
            # Primary method: RunPSScript
            $TCPLogin = (RunPSScript -PSScript $AzureEndpointLoginScriptBlock).TcpTestSucceeded
        }
        catch {
            Write-Message "Primary  TCP test failed for $AzureEndpointLogin. Using fallback test..." ForegroundColor Yellow
            try {
                $TCPLogin = Test-TcpConnection -ComputerName $AzureEndpointLogin -Port 443
            }
            catch {
                Write-Message "Fallback Test-TcpConnection failed for $AzureEndpointLogin." -Type Error
                $TCPLogin = $false
            }
        }

        # DNS Test
        try {
            # Primary method: RunPSScript
            $DNSLogin = (RunPSScript -PSScript $AzureEndpointLoginScriptBlock).NameResolutionSucceeded
        }
        catch {
            Write-Message "Primary  DNS test failed for $AzureEndpointLogin. Using fallback test..." -Type Info
            try {
                $DNSLogin = Test-DnsResolution -Hostname $AzureEndpointLogin
            }
            catch {
                Write-Message "Fallback Test-DnsResolution failed for $AzureEndpointLogin." -Type Error 
                $DNSLogin = $false
            }
        }

        # IWR Test
        $IWRLoginPage = Invoke-WebRequest -Uri $AzureEndpointLoginURI -UseBasicParsing
        $IWRLogin = if ($IWRLoginPage.StatusCode -eq 200) { $true } else { $False }

        #$IWRLogin = (RunPSScript -PSScript $IWRLoginScriptBlock)
        ########################################################################
        #Write-Host $TCPLogin " # " $DNSLogin " # " $IWRLogin
        if (($TCPLogin -and $DNSLogin) -or $IWRLogin) {
            ### write-Host "Test login.microsoftonline.com accessibility Passed" -ForegroundColor green 
            $objects += New-Object -Type PSObject -Prop @{'Test Name' = 'Checking accessibility to ' + $AzureEndpointLogin; 'Result' = 'Test Passed'; 'Recommendations' = "N/A"; 'Notes' = "N/A" }

            $loginAccessResult = "True"
        }
        Else {
            ### write-Host "Test login.microsoftonline.com accessibility Failed" -ForegroundColor red
            $loginAccessResult = "False"
            $objects += New-Object -Type PSObject -Prop @{'Test Name' = 'Checking accessiblity to ' + $AzureEndpointLogin; 'Result' = 'Test Failed'; 'Recommendations' = "Follow MS article for remediation: https://learn.microsoft.com/en-us/azure/active-directory/authentication/howto-mfa-nps-extension#network-requirements"; 'Notes' = "This will cause MFA Methods to fail" }

        }

        $TestStepNumber = $TestStepNumber + 1
        Write-Message "$TestStepNumber - Checking Accessibility to $AzureEndpointADNotificationURI  ..." -Type info
        

        ########################################################################
        # TCP Test
        #$TCPAdnotification = (RunPSScript -PSScript $AzureEndpointADNotificationScriptBlock).TcpTestSucceeded
        
        try {
            # Primary method: RunPSScript
            $TCPAdnotification = (RunPSScript -PSScript $AzureEndpointADNotificationScriptBlock).TcpTestSucceeded
        }
        catch {
            Write-Message "Primary  TCP test failed for $AzureEndpointADNotification. Using fallback test..." -Type Info
            try {
                $TCPAdnotification = Test-TcpConnection -ComputerName $AzureEndpointADNotification -Port 443
            }
            catch {
                Write-Message "Fallback Test-TcpConnection failed for $AzureEndpointADNotification." -Type Error
                $TCPAdnotification = $false
            }
        }

        # DNS Test
        #$DNSADNotification = (RunPSScript -PSScript $AzureEndpointADNotificationScriptBlock).NameResolutionSucceeded
        try {
            # Primary method: RunPSScript
            $DNSADNotification = (RunPSScript -PSScript $AzureEndpointADNotificationScriptBlock).NameResolutionSucceeded
        }
        catch {
            Write-Message "Primary  DNS test failed for $AzureEndpointADNotification. Using fallback test..." -Type Info
            try {
                $DNSADNotification = Test-DnsResolution -Hostname $AzureEndpointADNotification
            }
            catch {
                Write-Message "Fallback Test-DnsResolution failed for $AzureEndpointADNotification." -Type Error
                $DNSADNotification = $false
            }
        }

        # IWR Test
        # $DNSADNotificationPage = Invoke-WebRequest -Uri $AzureEndpointADNotificationURI
        # $IWRADNotification = if ($DNSADNotificationPage.StatusCode -eq "200") { $true } else { $False }
        $DNSADNotificationPage = Get-WebRequestStatusCode -Uri $AzureEndpointADNotificationURI
        $IWRADNotification = if ($DNSADNotificationPage) { $true } else { $False }
        

        #$IWRADNotification = (RunPSScript -PSScript $IWRADNotificationScriptBlock)
        ########################################################################
        #Write-Host $TCPAdnotification " # " $DNSADNotification " # " $IWRADNotification
        if (($TCPAdnotification -and $DNSADNotification) -or $IWRADNotification) {

            ### write-Host "Test adnotifications.windowsazure.com accessibility Passed" -ForegroundColor green
            $NotificationAccessResult = "True"

            $objects += New-Object -Type PSObject -Prop @{'Test Name' = 'Checking accessiblity to ' + $AzureEndpointADNotification; 'Result' = 'Test Passed'; 'Recommendations' = "N/A"; 'Notes' = "N/A" }
        }
        Else {
            ### write-Host "Test https://adnotifications.windowsazure.com accessibility Failed" -ForegroundColor Error

            $NotificationAccessResult = "False"

            $objects += New-Object -Type PSObject -Prop @{'Test Name' = 'Checking accessiblity to ' + $AzureEndpointADNotification; 'Result' = 'Test Failed'; 'Recommendations' = "Follow MS article for remediation: https://learn.microsoft.com/en-us/azure/active-directory/authentication/howto-mfa-nps-extension#network-requirements"; 'Notes' = "This will cause MFA Methods to fail" }


        }

        $TestStepNumber = $TestStepNumber + 1
        Write-Message "$TestStepNumber - Checking Accessibility to $AzureEndpointStrongAuthServiceURI  ..." -Type info


        ########################################################################
        #$TCPStrongAuthService = (RunPSScript -PSScript $AzureEndpointStrongAuthServiceScriptBlock).TcpTestSucceeded
        #$DNSStrongAuthService = (RunPSScript -PSScript $AzureEndpointStrongAuthServiceScriptBlock).NameResolutionSucceeded

        # TCP Test
        try {
            # Primary method: RunPSScript
            $TCPStrongAuthService = (RunPSScript -PSScript $AzureEndpointStrongAuthServiceScriptBlock).TcpTestSucceeded
        }
        catch {
            Write-Message "Primary  TCP test failed for $AzureEndpointStrongAuthService. Using fallback test..." -Type Info
            try {
                $TCPStrongAuthService = Test-TcpConnection -ComputerName $AzureEndpointStrongAuthService -Port 443
            }
            catch {
                Write-Message "Fallback Test-TcpConnection failed for $AzureEndpointStrongAuthService." -Type Error
                $TCPStrongAuthService = $false
            }
        }

        # DNS Test
        try {
            # Primary method: RunPSScript
            $DNSStrongAuthService = (RunPSScript -PSScript $AzureEndpointStrongAuthServiceScriptBlock).NameResolutionSucceeded
        }
        catch {
            Write-Message "Primary  DNS test failed for $AzureEndpointStrongAuthService. Using fallback test..." -Type Info
            try {
                $DNSStrongAuthService = Test-DnsResolution -Hostname $AzureEndpointStrongAuthService
            }
            catch {
                Write-Message "Fallback Test-DnsResolution failed for $AzureEndpointStrongAuthService." -Type Error
                $DNSStrongAuthService = $false
            }
        }

        # IWR Test
        # $IWRStrongAuthServicePage = Invoke-WebRequest -Uri $AzureEndpointStrongAuthServiceURI -UseBasicParsing
        # $IWRStrongAuthService = if ($IWRStrongAuthServicePage.StatusCode -eq 200) { $true } else { $False }
        $IWRStrongAuthServicePage = Get-WebRequestStatusCode -Uri $AzureEndpointStrongAuthServiceURI
        $IWRStrongAuthService = if ($IWRStrongAuthServicePage) { $true } else { $False }

        #$IWRStrongAuthService = (RunPSScript -PSScript $IWRStrongAuthServiceScriptBlock)
        ########################################################################
        #Write-Host $TCPStrongAuthService " # " $DNSStrongAuthService " # " $IWRStrongAuthService
        if (($TCPStrongAuthService -and $DNSStrongAuthService) -or $IWRStrongAuthService) {

            ### write-Host "Test strongauthenticationservice.auth.microsoft.com accessibility Passed" -ForegroundColor green

            $NotificationAccessResult = "True"

            $objects += New-Object -Type PSObject -Prop @{'Test Name' = 'Checking accessiblity to ' + $AzureEndpointStrongAuthService; 'Result' = 'Test Passed'; 'Recommendations' = "N/A"; 'Notes' = "N/A" }

        }
        Else {
            ### write-Host "Test https://strongauthenticationservice.auth.microsoft.com accessibility Failed" -ForegroundColor green

            $NotificationAccessResult = "False"

            $objects += New-Object -Type PSObject -Prop @{'Test Name' = 'Checking accessiblity to ' + $AzureEndpointStrongAuthService; 'Result' = 'Test Failed'; 'Recommendations' = "Follow MS article for remediation: https://learn.microsoft.com/en-us/azure/active-directory/authentication/howto-mfa-nps-extension#network-requirements"; 'Notes' = "This will cause MFA Methods to fail" }

        }

        $TestStepNumber = $TestStepNumber + 1
        Write-Message  "$TestStepNumber - Checking Accessibility to $AzureEndpointCredentialsURI  ..." -Type info
    

        ########################################################################
        #$TCPCredentials = (RunPSScript -PSScript $AzureEndpointCredentialsScriptBlock).TcpTestSucceeded
        #$DNSCredentials = (RunPSScript -PSScript $AzureEndpointCredentialsScriptBlock).NameResolutionSucceeded

        # TCP Test
        try {
            # Primary method: RunPSScript
            $TCPCredentials = (RunPSScript -PSScript $AzureEndpointCredentialsScriptBlock).TcpTestSucceeded
        }
        catch {
            Write-Message "Primary  TCP test failed for $AzureEndpointCredentials. Using fallback test..." -Type Info
            try {
                $TCPCredentials = Test-TcpConnection -ComputerName $AzureEndpointCredentials -Port 443
            }
            catch {
                Write-Message "Fallback Test-TcpConnection failed for $AzureEndpointCredentials." -Type Error
                $TCPCredentials = $false
            }
        }

        # DNS Test
        try {
            # Primary method: RunPSScript
            $DNSCredentials = (RunPSScript -PSScript $AzureEndpointCredentialsScriptBlock).NameResolutionSucceeded
        }
        catch {
            Write-Message "Primary  DNS test failed for $AzureEndpointCredentials. Using fallback test..." -Type Info
            try {
                $DNSCredentials = Test-DnsResolution -Hostname $AzureEndpointCredentials
            }
            catch {
                Write-Message "Fallback Test-DnsResolution failed for $AzureEndpointCredentials." -Type Error
                $DNSCredentials = $false
            }
        }

        # IWR Test
        # $IWRCredentialsPage = Invoke-WebRequest -Uri $AzureEndpointCredentialsURI -UseBasicParsing
        # $IWRCredentials = if ($IWRCredentialsPage.StatusCode -eq 200) { $true } else { $False }
        $IWRCredentialsPage = Get-WebRequestStatusCode -Uri $AzureEndpointCredentialsURI
        $IWRCredentials = if ($IWRCredentialsPage) { $true } else { $False }

        #$IWRCredentials = (RunPSScript -PSScript $IWRCredentialsScriptBlock)
        ########################################################################
        #Write-Host $TCPCredentials " # " $DNSCredentials " # " $IWRCredentials

        if (($TCPCredentials -and $DNSCredentials) -or $IWRCredentials) {

            ### write-Host "Test adnotifications.windowsazure.com accessibility Passed" -ForegroundColor green

            $NotificationAccessResult = "True"

            $objects += New-Object -Type PSObject -Prop @{'Test Name' = 'Checking accessiblity to ' + $AzureEndpointCredentialsURI; 'Result' = 'Test Passed'; 'Recommendations' = "N/A"; 'Notes' = "N/A" }

        }
        Else {
            ### write-Host "Test https://adnotifications.windowsazure.com accessibility Failed" -ForegroundColor red

            $NotificationAccessResult = "False"

            $objects += New-Object -Type PSObject -Prop @{'Test Name' = 'Checking accessiblity to ' + $AzureEndpointCredentialsURI; 'Result' = 'Test Failed'; 'Recommendations' = "Follow MS article for remediation: https://learn.microsoft.com/en-us/azure/active-directory/authentication/howto-mfa-nps-extension#network-requirements"; 'Notes' = "This will cause MFA Methods to fail" }
        }

        $TestStepNumber = $TestStepNumber + 1
        Write-Message "$TestStepNumber - Checking MFA version ... " -Type Info

        # Get MFA NPS installed version
        $MFAVersion = Get-WmiObject Win32_Product -Filter "Name like 'NPS Extension For Azure MFA'" | Select-Object -ExpandProperty Version

        # Get the latest version of MFA NPS Extension
        $MFADownloadPage = Invoke-WebRequest -Uri 'https://www.microsoft.com/en-us/download/details.aspx?id=54688'
        $MFADownloadPageHTML = $MFADownloadPage.RawContent
        $MFADownloadPageHTMLSplit = ($MFADownloadPageHTML -split '"version":"', 2)[1]
        $latestMFAVersion = ($MFADownloadPageHTMLSplit -split '","datePublished":', 2)[0]

        # Evaluate Download Page content
        # write-Host $MFADownloadPage
        # write-Host " # # # # # "
        # write-Host $MFADownloadPageHTML
        # write-Host
        # write-Host " # # # # # "
        # write-Host $MFADownloadPageHTMLSplit
        # write-Host
        # Write-Host $MFAVersion " # " $latestMFAVersion

        # Compare if the current version match the latest version
        if ($latestMFAVersion -le $MFAVersion) {

            # Display the Current MFA NPS version and mention it's latest one
            $MFATestVersion = "True"

            $objects += New-Object -Type PSObject -Prop @{'Test Name' = 'Checking if the current installed MFA NPS Extension Version is the latest'; 'Result' = 'Test Passed'; 'Recommendations' = "N/A"; 'Notes' = "The current installed version is the latest which is: " + $latestMFAVersion }

        }

        Else {

            # Display the Current MFA NPS version and mention it's Not the latest one, Advise to upgrade
            $MFATestVersion = "False"

            $objects += New-Object -Type PSObject -Prop @{'Test Name' = 'Checking if the current installed MFA NPS Extension Version is the latest'; 'Result' = 'Test Failed'; 'Recommendations' = "Make sure to upgrade to the latest version: " + $latestMFAVersion ; 'Notes' = "Current installed MFA Version is: " + $MFAVersion }

        }


        # Check if the NPS Service is Running or not

        $TestStepNumber = $TestStepNumber + 1
        Write-Message "$TestStepNumber - Checking if the NPS Service is Running ..." -Type Info

        if (((Get-Service -Name ias).status -eq "Running")) {

            $NPSServiceStatus = "True"

            ### write-Host "Passed" -ForegroundColor green
            $objects += New-Object -Type PSObject -Prop @{'Test Name' = 'Checking if NPS Service is Running'; 'Result' = 'Test Passed'; 'Recommendations' = "N/A"; 'Notes' = "N/A" }

        }

        Else {
            ### write-Host "Failed" -ForegroundColor Red
            $NPSServiceStatus = "False"

            $objects += New-Object -Type PSObject -Prop @{'Test Name' = 'Checking if NPS Service is Running'; 'Result' = 'Test Failed'; 'Recommendations' = "Troubleshoot NPS service, using MS article: https://learn.microsoft.com/en-us/troubleshoot/windows-server/networking/troubleshoot-network-policy-server"; 'Notes' = "N/A" }

        }

        # It will check the MS SPN in Cloud is Exist and Enabled
        $TestStepNumber = $TestStepNumber + 1
        Write-Message "$TestStepNumber - Checking if the SPN for Azure MFA Exists and is Enabled ..." -Type Info

        #Get All Registered SPNs in the tenant, save it in $AllSPNs variable

        $AllSPNs = ''
        $AllSPNs = Get-MgServicePrincipal -All | Select-Object AppId

        #if the MFA NPS is exist in $AllSPNs then it will check its status if it's enabled or not, if it doesn't exist the test will fail directly

        if ($AllSPNs -match "981f26a1-7f43-403b-a875-f8b09b8cd720") {
            $SPNExist = "True"
            $objects += New-Object -Type PSObject -Property @{'Test Name' = 'Checking if Azure MFA SPN Exists in the tenant'; 'Result' = 'Test Passed'; 'Recommendations' = "N/A"; 'Notes' = "N/A" }

            # Test if the SPN is enabled or Disabled
            if (((Get-MgServicePrincipal -Filter "appid eq '981f26a1-7f43-403b-a875-f8b09b8cd720'").AccountEnabled -eq $true)) {
                $SPNEnabled = "True"
                $objects += New-Object -Type PSObject -Property @{'Test Name' = 'Checking if Azure MFA SPN is Enabled in the tenant'; 'Result' = 'Test Passed'; 'Recommendations' = "N/A"; 'Notes' = "N/A" }
            }

            Else {

            
                $objects += New-Object -Type PSObject -Prop @{'Test Name' = 'Checking if Azure MFA SPN is Enabled in the tenant'; 'Result' = 'Test Failed'; 'Recommendations' = "Check if you have a valid MFA License and it's active for Azure MFA NPS. Follow MS article: https://learn.microsoft.com/en-us/azure/active-directory/authentication/howto-mfa-nps-extension#licenses"; 'Notes' = "If there is a valid non expired license, then consult MS Support" }

                ###write-Host "The SPN is Exist but not enabled, make sure that the SPN is enabled, Check your MFA license if it's valid - Test Failed" -ForegroundColor red
                $SPNEnabled = "False"
            }

        }

        Else {
            ###write-Host "The SPN Not Exist at all in your tenant, please check your MFA license if it's valid - Test Failed" -ForegroundColor red
            $SPNExist = "False"
            $SPNEnabled = "False"

            $objects += New-Object -Type PSObject -Prop @{'Test Name' = 'Checking if Azure MFA SPN Exists in the tenant'; 'Result' = 'Test Failed'; 'Recommendations' = "Check if you have a valid MFA License for Azure MFA NPSS. Follow MS article: https://learn.microsoft.com/en-us/azure/active-directory/authentication/howto-mfa-nps-extension#licenses"; 'Notes' = "If there is a valid non expired license, then consult MS Support" }

        }


        #check all registry keys for MFA NPS Extension

        # 1- It will check if the MFA NPS reg have the correct values.
        $TestStepNumber = $TestStepNumber + 1
        Write-Message "$TestStepNumber - Checking if Authorization and Extension Registry keys have the right values ... " -Type Info

        $AuthorizationDLLs = (Get-ItemProperty -path HKLM:\SYSTEM\CurrentControlSet\Services\AuthSrv\Parameters -name "AuthorizationDLLs").AuthorizationDLLs

        $ExtensionDLLs = (Get-ItemProperty -path HKLM:\SYSTEM\CurrentControlSet\Services\AuthSrv\Parameters -name "ExtensionDLLs").ExtensionDLLs

        if ($AuthorizationDLLs -eq "C:\Program Files\Microsoft\AzureMfa\Extensions\MfaNpsAuthzExt.dll" -and $ExtensionDLLs -eq "C:\Program Files\Microsoft\AzureMfa\Extensions\MfaNpsAuthnExt.dll") {

            ###Write-Host "MFA NPS AuthorizationDLLs and ExtensionDLLs Registries have the currect values - Test Passed" -ForegroundColor green

            $FirstSetofReg = "True"

            $objects += New-Object -Type PSObject -Prop @{'Test Name' = 'Checking if Authorization \ Extension Registry keys have the correct values'; 'Result' = 'Test Passed'; 'Recommendations' = "N/A"; 'Notes' = "N/A" }

        }

        Else {

            ### Write-Host "MFA NPS AuthorizationDLLs and/Or ExtensionDLLs Registries may have incorrect values - Test Failed" -ForegroundColor red

            $FirstSetofReg = "False"

            $objects += New-Object -Type PSObject -Prop @{'Test Name' = 'Checking if Authorization \ Extension Registry keys have the correct values'; 'Result' = 'Test Failed'; 'Recommendations' = "Follow MS article: https://learn.microsoft.com/en-us/azure/active-directory/authentication/howto-mfa-nps-extension-errors#troubleshooting-steps-for-common-errors"; 'Notes' = "As a quick solution, you can re-register MFA NPS extension again, by running its PowerShell script" }

        }

        # Check for other registry keys
        $TestStepNumber = $TestStepNumber + 1
        Write-Message "$TestStepNumber - Checking other Azure MFA related Registry keys have the right values ... " -Type Info


        $AZURE_MFA_HOSTNAME = (Get-ItemProperty -path HKLM:\SOFTWARE\Microsoft\AzureMfa -name "AZURE_MFA_HOSTNAME").AZURE_MFA_HOSTNAME

        $AZURE_MFA_RESOURCE_HOSTNAME = (Get-ItemProperty -path HKLM:\SOFTWARE\Microsoft\AzureMfa -name "AZURE_MFA_RESOURCE_HOSTNAME").AZURE_MFA_RESOURCE_HOSTNAME

        $AZURE_MFA_TARGET_PATH = (Get-ItemProperty -path HKLM:\SOFTWARE\Microsoft\AzureMfa -name "AZURE_MFA_TARGET_PATH").AZURE_MFA_TARGET_PATH

        $CLIENT_ID = (Get-ItemProperty -path HKLM:\SOFTWARE\Microsoft\AzureMfa -name "CLIENT_ID").CLIENT_ID

        $STS_URL = (Get-ItemProperty -path HKLM:\SOFTWARE\Microsoft\AzureMfa -name "STS_URL").STS_URL

        #if ($AZURE_MFA_HOSTNAME -eq "strongauthenticationservice.auth.microsoft.com" -and $AZURE_MFA_RESOURCE_HOSTNAME -eq "adnotifications.windowsazure.com" -and $AZURE_MFA_TARGET_PATH -eq "StrongAuthenticationService.svc/Connector" -and $CLIENT_ID -eq "981f26a1-7f43-403b-a875-f8b09b8cd720" -and $STS_URL -eq "https://login.microsoftonline.com/")
        if ($AZURE_MFA_HOSTNAME -eq $AzureEndpointStrongAuthService -and $AZURE_MFA_RESOURCE_HOSTNAME -eq $AzureEndpointADNotification -and $AZURE_MFA_TARGET_PATH -eq "StrongAuthenticationService.svc/Connector" -and $CLIENT_ID -eq "981f26a1-7f43-403b-a875-f8b09b8cd720" -and $STS_URL -eq $AzureEndpointLoginURISlash ) {

            ###Write-Host "MFA NPS other Registry keys have the currect values - Test Passed" -ForegroundColor green

            $SecondSetofReg = "True"

            $objects += New-Object -Type PSObject -Prop @{'Test Name' = 'Checking Other MFA Registry keys status'; 'Result' = 'Test Passed'; 'Recommendations' = "N/A"; 'Notes' = "N/A" }


        }

        Else {

            ###Write-Host "One or more registry key has incorrect value - Test Failed" -ForegroundColor green

            $SecondSetofReg = "False"

            $objects += New-Object -Type PSObject -Prop @{'Test Name' = 'Checking Other MFA Registry keys status'; 'Result' = 'Test Failed'; 'Recommendations' = "Re-register the MFA NPS extension or follow MS documentation"; 'Notes' = "If using Azure Government or Azure operated by 21Vianet clouds, follow MS article: https://learn.microsoft.com/en-us/azure/active-directory/authentication/howto-mfa-nps-extension#microsoft-azure-government-or-microsoft-azure-operated-by-21vianet-additional-steps" }


        }

        # below section is to check the current cert in Azure and current Cert in local NPS Server
        $TestStepNumber = $TestStepNumber + 1
        Write-Message "$TestStepNumber - Checking if there is a valid certificated matched with the Certificates stored in Entra ID ..." -Type Info

        # Count the number of certificate in the cloud for MFA NPS Extension
        $NumberofCert = (Get-MgServicePrincipal -Filter "appid eq '981f26a1-7f43-403b-a875-f8b09b8cd720'" -Property "KeyCredentials").KeyCredentials.Count

        # Store all the certificate in this variable; since customer may have more than one certificate and we need to check all of them, then we are storing the values of certs into array.
        $NPSCertValue = (Get-MgServicePrincipal -Filter "appid eq '981f26a1-7f43-403b-a875-f8b09b8cd720'" -Property "KeyCredentials").KeyCredentials

        # Get local Cert thumbprint from local NPS Server. 
        #$localCert =  (Get-ChildItem((Set-Location cert:\localmachine\my))).Thumbprint
        $localCert = (Get-ChildItem((Push-Location cert:\localmachine\my))).Thumbprint
        Pop-Location


        # $Tp will be used to store the Thumbprint for the cloud certs
        $TP = New-Object System.Collections.ArrayList

        # will be used to store the validity period of the Certs
        $Validity = New-Object System.Collections.ArrayList

        # Get the thumbprint for all Certificates in the cloud.
        for ($i = 0; $i -lt $NumberofCert; $i++) {

            $Cert = New-object System.Security.Cryptography.X509Certificates.X509Certificate2

            $Cert.Import([System.Text.Encoding]::UTF8.GetBytes([System.Convert]::ToBase64String($NPSCertValue[$i].Key)))
            $TP.Add($Cert.Thumbprint) | Out-Null
            $Validity.Add($cert.NotAfter) | Out-Null
        }


        # It will compare the thumbprint with the one's on the server, it will stop if one of the certificates were matched and still in it's validity period. All matched 
        #$result =Compare-Object -ReferenceObject ($localCert | Sort-Object) -DifferenceObject ($TP | Sort-Object)

        #if(!$result){echo "Matched"}

        # matched Cert from items in $localcert an $TP 

        $MatchedCert = @($TP | Where { $localCert -Contains $_ })

        if ($MatchedCert.count -gt 0) {

            $ValidCertThumbprint = @()
            $ValidCertThumbprintExpireSoon = @()

            # List All Matched Cetificate and still not expired, show Progress if the certificate will expire withen less than 30 days

            for ($x = 0; $x -lt $MatchedCert.Count ; $x++) {
   
                $CertTimeDate = $Validity[$TP.IndexOf($MatchedCert[$x])]

                $Diff = ((Get-Date) - $CertTimeDate).duration()
                   
                # If time difference less than 0, it means certificate has expired
                if ($Diff -lt 0) { 
                   
                    $certificateResult = "False"
                    $ValidCertThumbprint = "False"
                    $objects += New-Object -Type PSObject -Prop @{'Test Name' = 'Checking if there is a matched certificate with Azure MFA'; 'Result' = 'Test Failed'; 'Recommendations' = "Re-register the MFA NPS Extension again to generate new certificate, because current has expired"; 'Notes' = "More info: https://learn.microsoft.com/en-us/azure/active-directory/authentication/howto-mfa-nps-extension#how-do-i-verify-that-the-client-cert-is-installed-as-expected" }

                }
                   
                # If time difference is greater than 0 (still valid) and less than 30, it means certificate is valid but will expire soon
                Elseif ($Diff -gt 0 -and $Diff -lt 30 ) {
                   
                    $certificateResult = "True" 
                    $ValidCertThumbprint = $TP[$x]
                    $objects += New-Object -Type PSObject -Prop @{'Test Name' = 'Checking if there is a matched certificate with Azure MFA'; 'Result' = 'Test Passed'; 'Recommendations' = "Current certificate is valid for " + $Diff.Days + " days and will expire soon."; 'Notes' = "The matched Certificate(s) have these thumbprints: " + $ValidCertThumbprint + ". Follow MS article: https://learn.microsoft.com/en-us/azure/active-directory/authentication/howto-mfa-nps-extension#certificate-rollover" }
                   
                }
                # If time difference is greater than 30, it means certificate is valid for more than 1 month and less than 2 years
                Elseif ($Diff -gt 30 ) {

                    $certificateResult = "SuperTrue"
                    $ValidCertThumbprint = $TP[$x]
                    $objects += New-Object -Type PSObject -Prop @{'Test Name' = 'Checking if there is a matched certificate with Azure MFA'; 'Result' = 'Test Passed'; 'Recommendations' = "Current certificate is valid for " + $Diff.Days + " days"; 'Notes' = "The matched Certificate(s) have these thumbprints: " + $ValidCertThumbprint }
                   
                }
                   
            }
    
        }

        else {

            $certificateResult = "False"
            $objects += New-Object -Type PSObject -Prop @{'Test Name' = 'Checking if there is a matched certificate with Azure MFA'; 'Result' = 'Test Failed'; 'Recommendations' = "Re-register the MFA NPS Extension again to generate new certificate"; 'Notes' = "More info: https://learn.microsoft.com/en-us/azure/active-directory/authentication/howto-mfa-nps-extension#how-do-i-verify-that-the-client-cert-is-installed-as-expected" }

        }



        #list all missing Updates on the server

        #write-host "11- Checking all Missing Updates on the server ..." -ForegroundColor Yellow
        #write-host
        #
        #
        #$UpdateSession = New-Object -ComObject Microsoft.Update.Session
        #$UpdateSearcher = $UpdateSession.CreateupdateSearcher()
        #$Updates = @($UpdateSearcher.Search("IsHidden=0 and IsInstalled=0").Updates)
        #
        #if ($Updates -ne $null)
        #
        #{

        ###write-Host "List of missing updates on the server" -ForegroundColor Yellow


        #$ListofMissingUpdates = $Updates
        #
        #$updateResult = "False"
        #
        #     
        #   $objects += New-Object -Type PSObject -Prop @{'Test Name'='Checking missing Updates on the server';'Result'='Test Failed';'Recommendations' ="Usually we recommend to install all missing updates, please make a good plan before you proceed with the installtion";'Notes' = "Current missing updates is: " + $ListofMissingUpdates.title}
        #
        #
        #}
        #Else
        ##{
        #
        #### write-Host "The server is up to date" -ForegroundColor green
        #$updateResult = "True"
        #
        #$objects += New-Object -Type PSObject -Prop @{'Test Name'='Checking missing Updates on the server';'Result'='Test Passed';'Recommendations' ="N/A";'Notes' = "N/A"}
        #}
        #
        #
    }
    else {
        Write-Message "Connection to Entra Failed - Skipped all tests, please make sure to connect to your tenant first with global Admin role ..." -Type Error -BackgroundColor White
        Break
    }

    # Check if tests were done or not
    if ($null -ne $objects) {
        $Header = @"
<head>
<title>Azure MFA NPS Extension HealchCheck Report</title>
</head>
<body>
<p align ="Center"><font size="12" color="blue">Azure MFA NPS Extension Health Check Results</font></p>
</body>
<style>
table {
    font-family: arial, sans-serif;
    border-collapse: collapse;
    width: 100%;
    
}
td, th {
    border: 1px solid #dddddd;
    text-align: left;
    padding: 8px;
}
tr:nth-child(even) {
    background-color: #dddddd;
}
</style>
"@

        #$objects | ConvertTo-HTML -As Table -Fragment | Out-File c:\test1.html

        #cd c:\
        Push-Location "C:\"

        # Check if output directory C:\AzureReport is created. If not, create a new C:\AzureReport folder
        $DirectoryToCreate = "c:\AzureReport"
        if (-not (Test-Path -LiteralPath $DirectoryToCreate)) {
    
            try {
                New-Item -Path $DirectoryToCreate -ItemType Directory -ErrorAction Stop | Out-Null #-Force
            }
            catch {
                Write-Message -Message "Unable to create directory '$DirectoryToCreate'. Error was: $_" -ErrorAction Stop
            }
            Write-Message "Successfully created directory '$DirectoryToCreate'." -Type Success

        }
        else {
            Write-Message "Directory '$DirectoryToCreate' already existed" -Type Success
        }
        Remove-Item "c:\AzureReport\*.html"

        $objects | ConvertTo-Html -Head $Header | Out-File c:\AzureReport\AzureMFAReport.html

        Write-Message "The Report saved on this Path: C:\AzureReport\AzureMFAReport.html" -Type Success
        Pop-Location

    }

    Disconnect-MgGraph

}

##### This Function will be run against one affected user ######
##### Microsoft 2018 @Ahmad Yasin ##########

Function User_Test_Module {

    param(
        [string]$Cloud_Choice_Number
    )

    $Global:DialInStatus = 'N/A' # Define a non Null value to avoid conflict with the value restured from local AD when the user has no assigned policy under Dial-in tab in local AD

    $ErrorActionPreference = 'silentlycontinue'

    $Global:UPN = ''

    while ( $Global:UPN -eq '') {

        $Global:UPN = Read-Host -Prompt "`nEnter the UPN for the affected user in the format of User@MyDomain.com " 

    }

    $Global:UPN = $Global:UPN.Trim()

    Function Install_AD_Module {

        # Checking Active Directory Module
        Write-Message "Checking Active Directory Module..." -Type Info
        if (Get-Module -ListAvailable -Name ActiveDirectory) {
            #Importing Active Directory Module
            Import-Module ActiveDirectory
            Write-Message "Active Directory Module has imported." -Type info -BackgroundColor Black
        }
        else {
            Write-Message  "Active Directory Module is not installed." -Type Error -BackgroundColor Black
    
            #Installing Active Directory Module
            Write-Message  "Installing Active Directory Module..." -Type Info
            Add-WindowsFeature RSAT-AD-PowerShell
            
            Write-Message  "Active Directory Module has installed." -Type Error -BackgroundColor Black
            #Importing Active Directory Module
            Import-Module ActiveDirectory
            Write-Message "Active Directory Module has imported." -Type info -BackgroundColor Black
        }

    }


    Function Check_User {

        [String] $Global:UPN


        # This is replaced with 
        # Required script modules
        # Manage_Script_Libraries


        # if ($Cloud_Choice_Number -eq 'C') { 
    
        #     Connect-MgGraph -Scopes Domain.Read.All, User.Read.All, UserAuthenticationMethod.Read.All -NoWelcome -Environment Global

        # }

        # if ($Cloud_Choice_Number -eq 'G') { 

        #     Connect-MgGraph -Scopes Domain.Read.All, User.Read.All, UserAuthenticationMethod.Read.All -NoWelcome -Environment USGov

        # }

        # if ($Cloud_Choice_Number -eq 'V') { 

        #     Connect-MgGraph -Scopes Domain.Read.All, User.Read.All, UserAuthenticationMethod.Read.All -NoWelcome -Environment China

        # }

        
        $Global:verifyConnection = Get-MgDomain -ErrorAction SilentlyContinue # This will check if the connection succeeded or not

        $Global:DialInStatus = "N/A" # Initial value not null as option 3 in AD will be null value, to avoid conflict

        if ($null -ne $Global:verifyConnection) {
            Write-Message "Connection established Successfully - Starting the User Health Check Process ..." -Type Success
            Install_AD_Module

            
            # $Global:Result = (Get-MgUser -Filter "UserPrincipalName eq '$Global:upn'").UserPrincipalName  # Will check if the user exists in Entra ID based on the Provided UPN
            # $Global:IsSynced = (Get-MgUser -Filter "UserPrincipalName eq '$Global:upn'" -Property "OnPremisesImmutableId").OnPremisesImmutableId 
            # $Global:UserSignInStatus = (Get-MgUser -Filter "UserPrincipalName eq '$Global:upn'" -Property "AccountEnabled").AccountEnabled  # Check if the user is blocked to sign-in in Entra ID
            # $Global:SAMAccountName = (Get-ADUser -Filter "UserPrincipalName -eq '$Global:UPN'").SamAccountName 
            # $Global:DialInStatus = Get-ADUser $Global:SAMAccountName -Properties * | select -ExpandProperty msNPAllowDialin 
            # $Global:UserSyncErrorCount = (Get-MgUser -Filter "UserPrincipalName eq '$Global:upn'" -Property "OnPremisesProvisioningErrors").OnPremisesProvisioningErrors.Count  # Check if the user is healthy in Entra ID
            # $Global:UserLastSync = (Get-MgUser -Filter "UserPrincipalName eq '$Global:upn'" -Property "OnPremisesLastSyncDateTime").OnPremisesLastSyncDateTime # Check the last sync time for the user in Entra ID



            # Fetch user details from Entra ID (Microsoft Graph) and AD (Active Directory) in one call
            $Global:EntraUserInfo = Get-MgUser -Filter "UserPrincipalName eq '$($Global:UPN)'" -Property UserPrincipalName, OnPremisesImmutableId, AccountEnabled, OnPremisesProvisioningErrors, OnPremisesLastSyncDateTime

            # Assign key attributes from Entra ID object
            $Global:Result = $Global:EntraUserInfo.UserPrincipalName                   # Will check if the user exists in Entra ID based on the Provided UPN
            $Global:IsSynced = $Global:EntraUserInfo.OnPremisesImmutableId               # WIll check if the user is synced to Entra ID
            $Global:UserSignInStatus = $Global:EntraUserInfo.AccountEnabled                      # Check if the user is blocked to sign-in in Entra ID
            $Global:UserSyncErrorCount = $Global:EntraUserInfo.OnPremisesProvisioningErrors.Count  # Check if the user has any sync errors in Entra ID  
            $Global:UserLastSync = $Global:EntraUserInfo.OnPremisesLastSyncDateTime          # Check the last sync time for the user in Entra ID

            # Fetch on-prem AD attributes
            $Global:ADUserInfo = Get-ADUser -Filter "UserPrincipalName -eq '$($Global:UPN)'" -Properties SamAccountName, msNPAllowDialin
            $Global:SAMAccountName = $Global:ADUserInfo.SamAccountName      # Check user logon name
            $Global:DialInStatus = $Global:ADUserInfo.msNPAllowDialin     # check user NPS configuration 


            # If user doesn't exist on Entra ID, it's not able to get its MFA methods neither its license, returning error. If it does, let's return its MFA and licenses assigned
            if ($Global:Result -eq $Global:UPN) {
    
                $Global:StrongAuthMethods = Get-MgUserAuthenticationMethod -UserId $Global:upn  # To retrieve the current Strong Auth Methods configured
                $Global:UserAssignedLicense = (Get-MgUserLicenseDetail -UserId $Global:upn).SkuPartNumber #Check User Assigned license
                $Global:UserAssignedLicense = ($Global:UserAssignedLicense -replace ':', ' ')
                $Global:UserAssignedLicense = -split $Global:UserAssignedLicense
    
            }

            # Variable filled in from doc https://learn.microsoft.com/en-us/entra/identity/users/licensing-service-plan-reference
            $Global:UserPlans = "AAD_PREMIUM" , "AAD_PREMIUM_FACULTY" , "AAD_PREMIUM_USGOV_GCCHIGH" , "AAD_PREMIUM_P2" , "EMS_EDU_FACULTY" , "EMS", "EMSPREMIUM" , "EMSPREMIUM_USGOV_GCCHIGH" , "EMS_GOV" , "EMSPREMIUM_GOV" , "MFA_STANDALONE" , "M365EDU_A3_FACULTY" , "M365EDU_A3_STUDENT" , "M365EDU_A3_STUUSEBNFT" , "M365EDU_A3_STUUSEBNFT_RPA1" , "M365EDU_A5_FACULTY" , "M365EDU_A5_STUDENT" , "M365EDU_A5_STUUSEBNFT" , "M365EDU_A5_NOPSTNCONF_STUUSEBNFT" , "SPB" , "SPE_E3" , "SPE_E3_RPA1" , "Microsoft_365_E3" , "SPE_E3_USGOV_DOD" , "SPE_E3_USGOV_GCCHIGH" , "SPE_E5" , "Microsoft_365_E5" , "DEVELOPERPACK_E5" , "SPE_E5_CALLINGMINUTES" , "SPE_E5_NOPSTNCONF" , "Microsoft_365_E5_without_Audio_Conferencing" , "M365_F1" , "SPE_F1" , "M365_F1_COMM" , "SPE_E5_USGOV_GCCHIGH" , "M365_F1_GOV" , "M365_G3_GOV" , "M365_G5_GCC" , "MFA_PREMIUM"

            # VALUES OF USER ACCOUNT
            # Write-Host "UserPrincipalName: " $Global:Result
            # Write-Host "Is Synched: " $Global:IsSynced
            # Write-Host "MFA methods: " $Global:StrongAuthMethods | ConvertTo-Json
            # Write-Host "Sign-In status:" $Global:UserSignInStatus
            # Write-Host "SAMAccountName: " $Global:SAMAccountName
            # Write-Host "DialIn status:" $Global:DialInStatus
            # Write-Host "User Sync Error Count: " $Global:UserSyncErrorCount
            # Write-Host "Last Sync Date: " $Global:UserLastSync
            # Write-Host "License SKU: " $Global:UserAssignedLicense
            # Write-Host "Plans: " $Global:UserPlans

            # Write-Message "If no additional tests needed, Type Y and press Enter, This will remove the AD module which was installed at the beginning of this test. Removing the module requires a machine restart.`
            # If you don't want to remove it OR you need to perform the test again, press Enter directly." -BackgroundColor Red
            # $Global:Finishing_Test = Read-Host -Prompt 

            # if ($Global:Finishing_Test -eq "Y") {
            #     Write-Message "Thanks for Using MS Products, Removing AD module now ..." -Type Success
            #     Remove_AD_Module
            # }

        }

        else {
            write-Message "Connection to Entra Failed - Skipped all tests, please make sure to connect to your tenant first with global Admin role ..." -type Error -BackgroundColor White
            Break
        }

    }


    Function Remove_AD_Module {

        # Checking Active Directory Module
        Write-Message "Checking Active Directory Module..." -Type Info
        if (Get-Module -ListAvailable -Name ActiveDirectory) {
            Remove-WindowsFeature RSAT-AD-PowerShell
        } 
        
        else {
            Write-Message "Active Directory Module is not installed." -Type Error -BackgroundColor Black
        }

    }


    Function Test_Results {

        #Check if the user exists in AD, if not the test will be terminated

        $TestResultObjects = @()

        Write-Message "start Running the tests..." -Type Default

        Write-Message "Checking if $Global:UPN  EXISTS in Entra ID ... " -Type Info

        if ($Global:Result -eq $Global:UPN) {
            Write-Message " User $Global:UPN  EXISTS in Entra ID... TEST PASSED" -Type Success
        }
        else {

            Write-Message " User $Global:UPN  NOT EXISTS in Entra ID... TEST FAILED" -Type Error
            Write-Message " Test was terminated, Please make sure that the user EXISTS on Entra ID" -Type Error -BackgroundColor White
            
            Break
        }


        #Check if the user Synced to Entra ID, if Not the test will be terminated

        Write-Message "Checking if $Global:UPN  is SYNCHED to Entra ID from On-premises AD ... " -Type info

        if ($null -ne $Global:IsSynced -and $null -ne $Global:UserLastSync) {
            Write-Message " User $Global:UPN   is SYNCHED to Entra ID ... Test PASSED" -Type Success
        }
        else {
    
            Write-Message " User $Global:UPN  is NOT SYNCHED to Entra ID ... Test FAILED" -Type Error
            
            Write-Message " Test was terminated, Please make sure that the user is SYNCHED to Entra ID" -Type Error -BackgroundColor White

            Break
        }

        #Check if the user not blocked from Azure portal to sign in, even the test failed other tests will be performed
        Write-Message "Checking if $Global:UPN  is BLOCKED to sign in to Entra ID or Not ... " -Type info

        if ($Global:UserSignInStatus -eq $true) {

            Write-Message " User $Global:UPN  is NOT BLOCKED to sign in to Entra ID ... Test PASSED" -Type Success
    
        }
        else {

            Write-Message " User $Global:UPN  is BLOCKED to sign in to Entra ID ... Test FAILED" -Type Error
            Write-Message " Refer to: https://learn.microsoft.com/en-us/entra/fundamentals/how-to-manage-user-profile-info#add-or-change-profile-information for more info about this .... "  -Type Error -BackgroundColor White
            Write-Message " Test will continue to detect additional issue(s), Please make sure that the user is allowed to sign in to Entra ID" -Type Error -BackgroundColor White
        
        }


        #Check if the user is in healthy status in Entra ID, even the test failed other tests will be performed.
        Write-Message  "Checking if $Global:UPN  is HEALTHY in Entra ID or Not ..." -Type info

        if ($Global:UserSyncErrorCount -eq 0) {

            Write-Message " User $Global:UPN  status is HEALTHY in Entra ID ... Test PASSED" -Type Success
        }
        else {

            Write-Message " User $Global:UPN  is NOT HEALTHY in Entra ID ... Test FAILED" -Type Error
            Write-Message " Test will continue to detect additional issue(s), Please make sure that the user status is HEALTHY in Entra ID" -Type Error -BackgroundColor White
            
        }

        #Check if the user have MFA method(s) and there is one default MFA method.

        Write-Message "Checking if $Global:UPN already completed MFA Proofup in Entra ID or Not ... " -Type info

        $Global:HasMfaMethod = $false

        foreach ($method in $Global:StrongAuthMethods) {
            if ($method.AdditionalProperties["@odata.type"].Contains("phoneAuthenticationMethod") -or $method.AdditionalProperties["@odata.type"].Contains("microsoftAuthenticatorAuthenticationMethod")) {
                $Global:HasMfaMethod = $true
            }
        }

        if ($Global:HasMfaMethod -eq $false) {

            Write-Message " User $Global:UPN did NOT Complete the MFA Proofup at all or Admin require the user to provide MFA method again ... Test FAILED" -BackgroundColor Yellow
            Write-Message " Please refer to https://learn.microsoft.com/en-us/entra/identity/authentication/howto-mfa-getstarted#plan-user-registration for more info ... Test will continue to detect additional issue(s), Please make sure that the user has completed MFA Proofup in Entra ID" -BackgroundColor Yellow

        }
        else {

            Write-Message " User $Global:UPN  Completed MFA Proofup in Entra ID with $Global:DefaultMFAMethod as a Default MFA Method ... Test PASSED" -Type Success
            
        }

        #Check the user assigned licenses, usually even the user don't have direct assigned license the MFA will not fail, so only Progress we will throw here if the user have no license assigned
        # refer to this for the plans: https://learn.microsoft.com/en-us/azure/active-directory/users-groups-roles/licensing-service-plan-reference

        Write-Message  "Checking if $Global:UPN  has a valid license for MFA ... " -Type info

        # Check assigned licenses on valid licensing plans
        $IsMFALicenseValid = $false
        $MFALicense = $Global:UserAssignedLicense[0]

        # If there is no License assigned to user, make it noticed
        if ($MFALicense.Length -eq 0) {
            $MFALicense = "No License Assigned"
        }

        For ($i = 0; $i -lt $Global:UserAssignedLicense.Count; $i++) {
            For ($k = 0; $k -lt $Global:UserPlans.Count; $k++) {
                # Write-Host $Global:userAssignedLicense[$i] "#" $Global:UserPlans[$k]
                if ($Global:UserAssignedLicense[$i] -eq $Global:UserPlans[$k]) {
                    $MFALicense = $Global:UserAssignedLicense[$i]
                    $IsMFALicenseValid = $true
                }
            }
        }


        if ($IsMFALicenseValid) {

            Write-Message " User $Global:UPN  has a valid assigned license ( $MFALicense ) ... Test PASSED" -Type Success
        }
        else {

            Write-Message  " User $Global:UPN  has not a valid license for MFA ( $MFALicense ). It's a Progress message to be legal from licensing side... Test FAILED" -Type Error
            Write-Message  " Please, refer to https://learn.microsoft.com/en-us/azure/active-directory/users-groups-roles/licensing-service-plan-reference for more info ... " -Type Error -BackgroundColor White
            Write-Message  " Test will continue to detect additional issue(s), Please make sure to assign a valid MFA License for the user (AD Premium, EMS or MFA standalone license)" -Type Error -BackgroundColor White
        }

        #checking Network Access Permission under Dial-In Tab in AD, for more info refer to https://docs.microsoft.com/en-us/windows-server/networking/technologies/nps/nps-np-access

        Write-Message "Checking the Dial-In status for $Global:UPN in local AD" -Type info

        if ($null -ne $Global:SAMAccountName) {

            if ($Global:DialInStatus -eq $true) {

                Write-Message  " User $Global:UPN  allowed for Network Access Permission in local AD ... Test PASSED" -Type Success
                
            }
            elseif ($Global:DialInStatus -eq $false) {
                Write-Message " User $Global:UPN is Denied for Network Access Permission in local AD ... Test Failed" -Type Error
                Write-Message " Refer to https://learn.microsoft.com/en-us/windows-server/networking/technologies/nps/nps-np-access for more infor about this option" -Type Error -BackgroundColor White

            }
            elseif ($null -eq $Global:DialInStatus) {
                Write-Message " User $Global:UPN local AD Dial-In property is unchecked. This may result in denial, depending on the NPS policy. You need to check the NPS policy if the user is allowed or not." -Type Error
                Write-Message " Refer to https://learn.microsoft.com/en-us/windows-server/networking/technologies/nps/nps-np-access for more infor about this option " -Type Error -BackgroundColor White
            }

        }
        Else {
            Write-Message " For some reason, we are not able to get the SAMACCOUNTNAME for $Global:UPN From Local AD ... Hence we consider test was failed ..." -Type Error
        }

        #All Tests finished
        Write-Message " Check Completed. Please fix any issue identified and run the test again. If you required further troubleshooting, please contact MS support" -Type Success
    }

    Check_User

    Test_Results

    Disconnect-MgGraph

}

Function Collect_logs {

    $ErrorActionPreference = 'silentlycontinue'

    #start collecting logs

    Write-Message "Starting the log collection process ..." -Type Progress

    Set-Itemproperty -path 'HKLM:\SOFTWARE\Microsoft\AzureMfa' -Name 'VERBOSE_LOG' -value 'True'

    # Check if output directory C:\NPS is created. If not, create a new C:\NPS folder
    $DirectoryToCreate = "C:\NPS"
    if (-not (Test-Path -LiteralPath $DirectoryToCreate)) {
    
        try {
            New-Item -Path $DirectoryToCreate -ItemType Directory -ErrorAction Stop | Out-Null #-Force
        }
        catch {
            Write-Error -Message "Unable to create directory '$DirectoryToCreate'. Error was: $_" -ErrorAction Stop
        }
        Write-Message "Successfully created directory '$DirectoryToCreate'." -Type Success

    }
    else {
        Write-Message "Directory '$DirectoryToCreate' already existed. Removing existing files." -Type Info
    }
    Remove-Item "c:\nps\*.txt", "c:\nps\*.evtx", "c:\nps\*.etl", "c:\nps\*.log", "c:\nps\*.cab", "c:\nps\*.zip", "c:\nps\*.reg"

    netsh trace start capture=yes overwrite=yes  tracefile=C:\NPS\nettrace.etl
    REG QUERY "HKLM\SOFTWARE\Microsoft\AzureMfa" > C:\NPS\BeforeRegAdd_AzureMFA.txt
    REG QUERY "HKLM\SYSTEM\CurrentControlSet\Services\AuthSrv\Parameters" > C:\NPS\BeforeRegAdd_AuthSrv.txt
    REG ADD HKLM\SOFTWARE\Microsoft\AzureMfa /v VERBOSE_LOG /d TRUE /f
    net stop ias
    net start ias

    $npsext = "NPSExtension"
    $logmancmd = "logman create trace '$npsext' -ow -o C:\NPS\NPSExtension.etl -p {7237ED00-E119-430B-AB0F-C63360C8EE81} 0xffffffffffffffff 0xff -nb 16 16 -bs 1024 -mode Circular -f bincirc -max 4096 -ets"
    $logmancmdupdate = "logman update trace '$nps' -p {EC2E6D3A-C958-4C76-8EA4-0262520886FF} 0xffffffffffffffff 0xff -ets"
    cmd /c $logmancmd

    Write-Message "If you see 'Error: Data Collector Set was not found.' after this, that is " -Type info  -NoNewline
    Write-Message  "GOOD," -Type Success -NoNewline 
    Write-Message " if not then it means the files already existed in C:\NPS."  -NoNewLine

    cmd /c $logmancmdupdate

    Write-Message "Please Reproduce the issue quickly, once you finish please Press the Enter key to finish and gather logs." -BackgroundColor 
    Read-Host

    # Stop and Collect the logs
    $logmanstop = "logman stop '$npsext' -ets"
    cmd /c $logmanstop
    netsh trace stop
    REG QUERY "HKLM\SOFTWARE\Microsoft\AzureMfa" > C:\NPS\AfterRegAdd_AzureMFA.txt
    REG QUERY "HKLM\SYSTEM\CurrentControlSet\Services\AuthSrv\Parameters" > C:\NPS\AfterRegAdd_AuthSrv.txt
    REG ADD HKLM\SOFTWARE\Microsoft\AzureMfa /v VERBOSE_LOG /d FALSE /f
    Set-Itemproperty -path 'HKLM:\SOFTWARE\Microsoft\AzureMfa' -Name 'VERBOSE_LOG' -value 'False'
    wevtutil epl AuthNOptCh C:\NPS\%computername%_AuthNOptCh.evtx /ow:True
    wevtutil epl AuthZOptCh C:\NPS\%computername%_AuthZOptCh.evtx
    wevtutil epl AuthZAdminCh C:\NPS\%computername%_AuthZAdminCh.evtx
    wevtutil qe Security "/q:*[System [(EventID=6272) or (EventID=6273) or (EventID=6274)]]" /f:text | out-file c:\nps\NPS_EventLog.log

    $Compress = @{
        Path             = "c:\nps\*.txt", "c:\nps\*.evtx", "c:\nps\*.etl", "c:\nps\*.log", "c:\nps\*.cab"
        CompressionLevel = "Fastest"
        DestinationPath  = "c:\nps\" + $Timestamp + "_NpsLogging.zip"
    }  
    Write-Message  "Compressing log files." -Type info
    Compress-Archive @Compress

    Write-Message  "Data collection has completed.  Please upload the most recent Zip file to MS support. " -Type Success

    Invoke-Item c:\nps
    Break

}

Function MFAorNPS {

    # This test will remove the MFA registry key and restart NPS, so that you can determine if the issue related to MFA or NPS.

    $AuthorizationDLLs_Backup = ''
    $ExtensionDLLs_Backup = ''

    $AuthorizationDLLs_Backup = (Get-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\AuthSrv\Parameters -Name AuthorizationDLLs).AuthorizationDLLs
    $ExtensionDLLs_Backup = (Get-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\AuthSrv\Parameters -Name ExtensionDLLs).ExtensionDLLs

    Write-Message  "In this test we will remove some registry keys that will bypass the MFA module to determine if the issue is related to the MFA extension or the NPS role.  `nAfter the test finishes the regkeys will be restored." -Type info
    Write-Message   "Press ENTER to continue, otherwise please close the PowerShell window or hit CTRL+C to exit script." -BackgroundColor Red
    Read-Host

    # Check if output directory C:\NPS is created. If not, create a new C:\NPS folder
    $DirectoryToCreate = "C:\NPS"
    if (-not (Test-Path -LiteralPath $DirectoryToCreate)) {
    
        try {
            New-Item -Path $DirectoryToCreate -ItemType Directory -ErrorAction Stop | Out-Null #-Force
        }
        catch {
            Write-Error -Message "Unable to create directory '$DirectoryToCreate'. Error was: $_" -ErrorAction Stop
        }
        Write-Message "Successfully created directory '$DirectoryToCreate'." -type success

    }
    else {
        Write-Message  "Directory '$DirectoryToCreate' already existed" -Type Info
    }
    Remove-Item "c:\nps\*.txt", "c:\nps\*.evtx", "c:\nps\*.etl", "c:\nps\*.log", "c:\nps\*.cab", "c:\nps\*.zip", "c:\nps\*.reg"

    # Export NPS MFA registry keys
    Write-Message  "Exporting the NPS MFA registry keys." -Type info

    reg export hklm\system\currentcontrolset\services\authsrv c:\nps\AuthSrv.reg /y 

    Set-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\AuthSrv\Parameters -Name AuthorizationDLLs -Value ''
    Set-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\AuthSrv\Parameters -Name ExtensionDLLs -Value ''

    Write-Message   "Restarting NPS" -Type Progress
    Stop-Service -Name "IAS" -Force
    Start-Service -Name "IAS"
    Write-Message  "NPS has been restarted.  MFA is not being used at this time." -Type Success 
    
    Write-Message  "Try to repro the issue now.  If the user is now able to connect successfully without MFA then the issue is related more to the MFA module.  `nAfter you finish this test press Enter to restore the MFA functionality."  -BackgroundColor Red
    Read-Host

    Set-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\AuthSrv\Parameters -Name AuthorizationDLLs -Value $AuthorizationDLLs_Backup
    Set-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\AuthSrv\Parameters -Name ExtensionDLLs -Value $ExtensionDLLs_Backup

    $AuthorizationDLLs_Backup = (Get-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\AuthSrv\Parameters -Name AuthorizationDLLs).AuthorizationDLLs
    $ExtensionDLLs_Backup = (Get-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\AuthSrv\Parameters -Name ExtensionDLLs).ExtensionDLLs

    if ($null -ne $AuthorizationDLLs_Backup -and $null -ne $ExtensionDLLs_Backup) {

        Write-Message   "Registry Keys were restored, restarting NPS." -Type  Progress

        Stop-Service -Name "IAS" -Force
        Start-Service -Name "IAS"
        Write-Message "NPS has been restarted.  MFA has been reenabled." -Type Success

    }
    Else {

        Write-Message "Something went wrong while restoring the Registries, please import them manually from C:\NPS\AuthSrv.reg and restart the NPS Service. Hit Enter now to open Services and C:\NPS " -Type Error
        Read-Host
        services.msc
        Invoke-Item c:\nps

    }

    Break

}

##### This function evaluates the need to install MS Graph libraries #####
##### Microsoft 2024 @Miguel Ferreira #####

# Function Manage_Script_Libraries {
#     # Install required MG Graph modules
    
#     Write-Host "Ensure Microsoft.Graph module is installed ..." -ForegroundColor Yellow
    

#     # Required MG Graph modules
#     $moduleName = "Microsoft.Graph.Authentication"
#     if (Get-Module -ListAvailable -Name $moduleName) {
#         Write-Message  "$moduleName module available" -ForegroundColor Yellow
#     }
#     else {
#         Write-Message  "Installing" $moduleName "module ..." -ForegroundColor Yellows
#         Install-Module -Name $moduleName -ErrorAction Stop
#     }

#     $moduleName = "Microsoft.Graph.Applications"
#     if (Get-Module -ListAvailable -Name $moduleName) {
#         Write-Message  $moduleName "module available" -ForegroundColor Cyan
#         Import-Module -Name $moduleName
#     }
#     else {
#         Write-Message  "Installing" $moduleName "module ..." -ForegroundColor Yellow
#         Install-Module -Name $moduleName -ErrorAction Stop
#     }

#     $moduleName = "Microsoft.Graph.Users"
#     if (Get-Module -ListAvailable -Name $moduleName) {
#         Write-Message  $moduleName "module available" -ForegroundColor Cyan
#         Import-Module -Name $moduleName
#     }
#     else {
#         Write-Message  "Installing" $moduleName "module ..." -ForegroundColor Yellow
#         Install-Module -Name $moduleName -ErrorAction Stop
#     }

#     $moduleName = "Microsoft.Graph.Identity.DirectoryManagement"
#     if (Get-Module -ListAvailable -Name $moduleName) {
#         Write-Message  $moduleName "module available" -ForegroundColor Cyan
#         Import-Module -Name $moduleName
#     }
#     else {
#         Write-Message "Installing" $moduleName "module ..." ForegroundColor Yellow
#         Install-Module -Name $moduleName -ErrorAction Stop
#     }

#     $moduleName = "Microsoft.Graph.Identity.SignIns"
#     if (Get-Module -ListAvailable -Name $moduleName) {
#         Write-Message  $moduleName "module available" -ForegroundColor Cyan
#         Import-Module -Name $moduleName
#     }
#     else {
#         Write-Message  "Installing" $moduleName "module ..." ForegroundColor Yellow
#         Install-Module -Name $moduleName -ErrorAction Stop
#     }

# }


function Uninstall_Script_Libraries {
    # Remove modules from memory
    $ModulesMemory = Get-Module Microsoft.Graph* -ListAvailable | Select-Object Name -Unique

    Foreach ($Module in $ModulesMemory) {
        $ModuleName = $Module.Name
        Write-Message "Remove-Module" $ModuleName -Type info
        Remove-Module -Name $ModuleName
    }

    #Uninstall Microsoft.Graph modules except Microsoft.Graph.Authentication
    $Modules = Get-Module Microsoft.Graph* -ListAvailable | 
    Where-Object { $_.Name -ne "Microsoft.Graph.Authentication" } | Select-Object Name -Unique

    Foreach ($Module in $Modules) {
        $ModuleName = $Module.Name
        $Versions = Get-Module $ModuleName -ListAvailable
        Foreach ($Version in $Versions) {
            $ModuleVersion = $Version.Version
            Write-Message "Uninstall-Module $ModuleName $ModuleVersion" -Type info
            Uninstall-Module $ModuleName -RequiredVersion $ModuleVersion -ErrorAction SilentlyContinue
        }
    }

    #Uninstall the modules cannot be removed from first part.
    $InstalledModules = Get-InstalledModule Microsoft.Graph* | 
    Where-Object { $_.Name -ne "Microsoft.Graph.Authentication" } | Select-Object Name -Unique

    Foreach ($InstalledModule in $InstalledModules) {
        $InstalledModuleName = $InstalledModule.Name
        $InstalledVersions = Get-Module $InstalledModuleName -ListAvailable
        Foreach ($InstalledVersion in $InstalledVersions) {
            $InstalledModuleVersion = $InstalledVersion.Version
            Write-Message "Uninstall-Module $InstalledModuleName $InstalledModuleVersion" -Type info
            Uninstall-Module $InstalledModuleName -RequiredVersion $InstalledModuleVersion -ErrorAction SilentlyContinue
        }
    }

    #Uninstall Microsoft.Graph.Authentication
    $ModuleName = "Microsoft.Graph.Authentication"
    $Versions = Get-Module $ModuleName -ListAvailable

    Foreach ($Version in $Versions) {
        $ModuleVersion = $Version.Version
        Write-Message "Uninstall-Module $ModuleName $ModuleVersion" -Type info
        Remove-Module -Name $ModuleName -Force
        Uninstall-Module $ModuleName -RequiredVersion $ModuleVersion
    }
}



if ($Choice_Number -eq 'E') { Break }
if ($Choice_Number -eq '1') { MFAorNPS }
if ($Choice_Number -eq '2') { Check_Nps_Server_Module -Cloud_Choice_Number $Cloud_Choice_Number }
if ($Choice_Number -eq '3') { User_Test_Module -Cloud_Choice_Number $Cloud_Choice_Number }
if ($Choice_Number -eq '4') { collect_logs }