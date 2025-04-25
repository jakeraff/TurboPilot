param(
    [string]$TenantId,
    [string]$ClientId,
    [string]$ClientSecret,
    [string]$DeploymentProfile,
    [string]$AssignedUser,
    [string]$ComputerName
)

function Write-HostCenter { param($Message) Write-Host; Write-Host ("{0}{1}" -f (' ' * (([Math]::Max(0, $Host.UI.RawUI.BufferSize.Width / 2) - [Math]::Floor($Message.Length / 2)))), $Message) -f Cyan; Write-Host }

function ShowInfo {
    Write-HostCenter "TurboPilot - Windows Enrollment Script"
    Write-Host "Author: Jacob Raffoul"
    Write-Host "Version: 1.1.0"
    Write-Host "License: GPL-3.0" 
}

function LoadModules{
    param (
        [string]$TempFolder = 'C:\Windows\Temp\TurboPilot',
        [switch]$SetupComplete
    )

    Write-HostCenter 'Loading Modules from Files'

    $Modules = @(
        @{ Name = 'Microsoft.Graph.Authentication'; Version = '2.26.1' },
        @{ Name = 'Microsoft.Graph.Users'; Version = '2.26.1' },
        @{ Name = 'Microsoft.Graph.Groups'; Version = '2.26.1' },
        @{ Name = 'Microsoft.Graph.Intune'; Version = '6.1907.1' },
        @{ Name = 'WindowsAutopilotIntune'; Version = '5.7.0' }
    )
    
    if ($SetupComplete) {
        Write-HostCenter 'Loading Modules from PSGallery'
    
        if (Get-PackageProvider -ListAvailable -Name NuGet -ErrorAction SilentlyContinue) { Write-Host "NuGet is already installed!" }
        else { Install-PackageProvider NuGet -Confirm:$false -Force }
    
        foreach ($Module in $Modules) {
            if ((Get-Module -ListAvailable -Name $Module.Name).Version -like $Module.Version) {
                Write-Host "$($Module.Name) is already installed. Importing..."
                Import-Module $Module.Name
            }
            else {
                Write-Host "$($Module.Name) not found! Installing and importing..."
                Install-Module $Module.Name -RequiredVersion $Module.Version -Confirm:$false -Force | Import-Module
            }
        }
    }
    else {
        ## Download and extract required modules
        $Links = $Modules | ForEach-Object {
            "https://psg-prod-eastus.azureedge.net/packages/$($_.Name.ToLower()).$($_.Version).nupkg"
        }
        New-Item -ItemType Directory $TempFolder -Force | Out-Null
    
        $ModulesRequired = $Links | ForEach-Object { $_.split('/').trim('.nupkg') | Select-Object -Last 1 }
        $ModulesPresent = Get-ChildItem $TempFolder | Where-Object { Get-ChildItem $_.FullName -Filter '*.psd1' } | Select-Object -ExpandProperty Name
    
        ## Check if modules are already present by comparing the arrays
        if (!($ModulesPresent)) { $ModulesPresent = @() }
        if ($ModulesRequired -notlike $ModulesPresent) {
            Compare-Object $ModulesPresent $ModulesRequired -IncludeEqual | Sort-Object -Property InputObject | ForEach-Object {
                if ($_.SideIndicator -eq '<=') { continue }
                $Module = $_.InputObject
                $Path = "$TempFolder\$Module.zip"
                $Folder = $Path.Replace('.zip','')
                Write-Host -NoNewLine "Loading module: $Module"
                if ($_.SideIndicator -eq '=>') {
                    $Link = $Links | Where-Object { $_ -like "*$Module*" }
                    Invoke-WebRequest -Uri $Link -OutFile $Path
                    Write-Host ' | ' -F White -NoNewline
                    Write-Host 'Downloaded' -NoNewline -F White
                    New-Item -ItemType Directory $Folder -Force | Out-Null
                    Expand-Archive -Path $Path -DestinationPath $Folder -Force
                    Remove-Item -Path $Path -Force
                    Write-Host ' | ' -F White -NoNewline
                    Write-Host 'Extracted' -NoNewline -F Yellow
                }
                Import-Module (Get-ChildItem $Folder -Filter '*.psd1' | Select-Object -ExpandProperty FullName)
                Write-Host ' | ' -F White -NoNewline
                Write-Host 'Imported' -F Green
            }
        }
    }
}

function ConnectGraph {
    param(
        [String]$TenantId,
        [String]$ClientId,
        [String]$ClientSecret,
        [switch]$SetupComplete
    )

    Write-HostCenter "Connecting to Microsoft Graph"

    $Scopes = @(
        'Group.ReadWrite.All',
        'Device.ReadWrite.All',
        'DeviceManagementManagedDevices.ReadWrite.All',
        'DeviceManagementServiceConfig.ReadWrite.All',
        'GroupMember.ReadWrite.All',
        'User.Read.All'
    )

    if ($TenantId -and $ClientId -and $ClientSecret) {
        ## Connect silently
        $ClientSecretPass = ConvertTo-SecureString -String $ClientSecret -AsPlainText -Force
        $ClientSecretCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ClientId, $ClientSecretPass
        Connect-MgGraph -TenantId $tenantId -ClientSecretCredential $ClientSecretCredential
    } else {
        ## Connect interactively
        if ($SetupComplete) { Connect-MgGraph -Scopes $Scopes }
        else { Connect-MgGraph -Scopes $Scopes -UseDeviceCode }
    }
}

function OptionsCLI {
    param(
        [string]$DeploymentProfile,
        [string]$AssignedUser,
        [string]$ComputerName
    )

    Write-HostCenter 'Collating Autopilot Parameters'

    $Options = @{}

    # Validate that the deployment profile exists
    if ($DeploymentProfile) {
        $Profile = Get-AutopilotProfile | Where-Object displayName -eq $DeploymentProfile | Select-Object -First 1
        if (-not $Profile) {
            Write-Error "Deployment profile '$DeploymentProfile' not found."
            Exit 1
        }
        # Only allow profiles with a group tag
        $Assignment = Get-AutopilotProfileAssignments -Id $Profile.Id | Select-Object -First 1
        $GroupID = $Assignment.Id
        if (-not $GroupID) {
            Write-Error "Deployment profile '$DeploymentProfile' has no assignments."
            Exit 1
        }
        $Rule = (Get-MgGroup -GroupId $GroupID).membershipRule
        $GroupTag = $Rule.Split(':')[1].Trim(')','"')
        if (-not $GroupTag) {
            Write-Error "Deployment profile '$DeploymentProfile' has no grouptag."
            Exit 1
        }
        $Options['Profile'] = $DeploymentProfile
    }

    # Validate that the user exists
    if ($AssignedUser) {
        if (!(Get-MgUser -Filter "UserPrincipalName eq '$AssignedUser'" -All | Select-Object -First 1)) {
            Write-Error "User '$AssignedUser' not found."
            Exit 1
        }
        else { $Options['User'] = $AssignedUser }
    }

    # Validate computer name
    if ($ComputerName) {
        if ($ComputerName.Length -gt 15 -or $ComputerName -notmatch '^[a-zA-Z0-9\-]+$') {
            Write-Error "Computer name '$ComputerName' is invalid. Must be 1-15 characters, letters, numbers, hyphens only."
            Exit 1
        }
        else { $Options['Name'] = $ComputerName }
    }

    Write-Host "Selection: $(ConvertTo-Json $Options -Compress)"
    return $Options
}

function OptionsGUI {
    ## Build provisioning GUI
    Write-HostCenter 'Building Provisioning GUI'

    [void] [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")

    Write-Host "Loaded WinForms."

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'TurboPilot'
    $form.Size = New-Object System.Drawing.Size(200,240)
    $form.StartPosition = 'CenterScreen'
    $Form.FormBorderStyle = 'Fixed3D'
    $Form.ControlBox = $false
    $Form.MaximizeBox = $false
    $form.TopMost = $true

    Write-Host "Built frame."

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Size = New-Object System.Drawing.Size(50,25)
    $okButton.Location = New-Object System.Drawing.Point(20, ($form.ClientSize.Height - 40))
    $okButton.Anchor = 'Bottom'
    $okButton.Text = 'OK'
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $okButton
    $form.Controls.Add($okButton)

    Write-Host "Added OK button."

    $ProfileLabel = New-Object System.Windows.Forms.Label
    $ProfileLabel.Location = New-Object System.Drawing.Point(22,0)
    $ProfileLabel.Size = New-Object System.Drawing.Size(($form.ClientSize.Width-$ProfileLabel.Location.X * 2),20)
    $ProfileLabel.Text = "Deployment profile:"
    $ProfileLabel.Anchor = 'top'
    $form.Controls.Add($ProfileLabel)

    Write-Host "Added profile list label."

    $ProfileList = New-Object system.Windows.Forms.ComboBox
    $ProfileList.text = ""
    $ProfileList.Location = New-Object System.Drawing.Point(20,($ProfileLabel.Location.Y+20))
    $ProfileList.width = $form.ClientSize.Width-$ProfileList.Location.X*2
    $ProfileList.Anchor = 'top'

    ## Gets list of all AP profiles and adds them to the drop down
    $Profiles = Get-AutopilotProfile | Where-Object {
        # Only show profiles with a group tag
        $Assignment = Get-AutopilotProfileAssignments -Id $_.Id | Select-Object -First 1
        $GroupID = $Assignment.Id
        if (-not $GroupID) { return $false }
        $Rule = (Get-MgGroup -GroupId $GroupID).membershipRule
        $GroupTag = $Rule.Split(':')[1].Trim(')','"')
        return -not [string]::IsNullOrEmpty($GroupTag)
    }
    [array]::Reverse($Profiles)
    foreach ($Profile in $Profiles) { [void] $ProfileList.Items.Add("$($Profile.displayName)") }
    Write-Host "Downloaded Autopilot profiles."

    ## Gets the current profile assigned to the device (if any) and makes it the default selection
    $serial = (Get-ComputerInfo -Property BiosSeralNumber | Select-Object -ExpandProperty BiosSeralNumber)
    $APID = (Get-AutopilotDevice -serial $serial).id
    if ($Null -ne $APID) { $CurrentProfile = (Get-AutopilotDevice -id $APID -Expand).deploymentProfile.DisplayName }
    if ($CurrentProfile) { $ProfileList.SelectedItem = $CurrentProfile }
    else { $ProfileList.SelectedIndex = 0 }
    $form.Controls.Add($ProfileList)
    Write-Host "Added deployment profile field."

    $UserLabel = New-Object System.Windows.Forms.Label
    $UserLabel.Location = New-Object System.Drawing.Point(22,($ProfileList.Location.Y+30))
    $UserLabel.Size = New-Object System.Drawing.Size(($form.ClientSize.Width-$UserLabel.Location.X*2),20)
    $UserLabel.Text = "Assigned User:"
    $UserLabel.Anchor = 'top'
    $form.Controls.Add($UserLabel)

    Write-Host "Added assigned user label."

    $UserClear = New-Object System.Windows.Forms.Button
    $UserClear.Size = New-Object System.Drawing.Size(16,16)
    $UserClear.Location = New-Object System.Drawing.Point(($form.ClientSize.Width-$UserLabel.Location.X*2+4), ($UserLabel.Location.Y+23))
    $UserClear.Text = 'X'
    $UserClear.Add_Click({ $UserInput.Clear() })
    $form.Controls.Add($UserClear)

    $UserInput = New-Object System.Windows.Forms.TextBox
    $UserInput.text = ""
    $UserInput.Location = New-Object System.Drawing.Point(20,($UserLabel.Location.Y+20))
    $UserInput.Width = $form.ClientSize.Width-$UserInput.Location.X*2
    $UserInput.Anchor = 'top'

    # Get Users
    $Users = Get-MgUser -Filter "accountEnabled eq true" -Property UserPrincipalName -All | Select-Object UserPrincipalName
    $UserInput.AutoCompleteSource = 'CustomSource'
    $UserInput.AutoCompleteMode = 'Suggest'
    $AutoComplete = New-Object System.Windows.Forms.AutoCompleteStringCollection
    $UserInput.AutoCompleteCustomSource = $AutoComplete
    $AutoComplete.AddRange($Users.UserPrincipalName)

    # Pre-fill the input box with the current assigned user if available
    $serial = (Get-ComputerInfo -Property BiosSeralNumber | Select-Object -ExpandProperty BiosSeralNumber)
    $CurrentUser = (Get-AutopilotDevice -serial $serial).userPrincipalName
    if ($CurrentUser) { $UserInput.Text = $CurrentUser }

    # Clear the input box if the user inputs an invalid UPN
    $UserInput.Add_LostFocus({
        if ($Users.UserPrincipalName -notcontains $UserInput.Text) {
            $UserInput.Clear()
        }
        else {
            $UserInput.SelectionStart = 0
            $UserInput.SelectionLength = 0
            $UserInput.ScrollToCaret()
        }
    })

    $form.Controls.Add($UserInput)

    Write-Host "Added assigned user field."

    $NameLabel = New-Object System.Windows.Forms.Label
    $NameLabel.Location = New-Object System.Drawing.Point(22,($UserInput.Location.Y+30))
    $NameLabel.Size = New-Object System.Drawing.Size(($form.ClientSize.Width-$NameLabel.Location.X*2),20)
    $NameLabel.Text = "Computer Name:"
    $NameLabel.Anchor = 'top'
    $form.Controls.Add($NameLabel)

    Write-Host "Added computer name label."

    $NameClear = New-Object System.Windows.Forms.Button
    $NameClear.Size = New-Object System.Drawing.Size(16,16)
    $NameClear.Location = New-Object System.Drawing.Point(($form.ClientSize.Width-$NameLabel.Location.X*2+4), ($NameLabel.Location.Y+23))
    $NameClear.Text = 'X'
    $NameClear.Add_Click({ $NameInput.Clear() })
    $form.Controls.Add($NameClear)

    $NameInput = New-Object System.Windows.Forms.TextBox
    $NameInput.text = (hostname)
    $NameInput.Location = New-Object System.Drawing.Point(20,($NameLabel.Location.Y+20))
    $NameInput.width = $form.ClientSize.Width-$NameInput.Location.X*2
    $NameInput.Anchor = 'top'
    $NameInput.MaxLength = 15

    # Prevent illegal chars from being typed (allow only A–Z, a–z, 0–9, hyphen, backspace)
    $NameInput.Add_KeyPress({
        if ($_.KeyChar -notmatch '^[A-Za-z0-9-]$' -and $_.KeyChar -ne [char]8) {
            $_.Handled = $true
        }
    })
    
    $form.Controls.Add($NameInput)

    $result = $form.ShowDialog()
    
    if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
        Write-Host "Script exited!" -B Red -F White
        timeout /t -1
        Exit 0
    }

    $Options = @{ Profile = $ProfileList.SelectedItem }
    if (!([String]::IsNullOrWhiteSpace($UserInput.Text))) { $Options['User'] = $UserInput.Text }
    if (!([String]::IsNullOrWhiteSpace($NameInput.Text))) { $Options['Name'] = $NameInput.Text }

    Write-Host "Selection: $(ConvertTo-Json $Options -Compress)"
    return $Options
}

function ImportDevice {
    param (
        $Options
    )

    Write-HostCenter 'Processing Autopilot Request'

    ## Find GroupTag
    $serial = (Get-ComputerInfo -Property BiosSeralNumber | Select-Object -ExpandProperty BiosSeralNumber)
    $deviceId = (Get-AutopilotDevice -serial $serial).id
    if ($deviceId) { $Device = Get-AutopilotDevice -id $deviceId -expand }

    $OldOptions = @{}
    if ($Options.Profile) {
        $SelectedProfile = Get-AutopilotProfile | Where-Object displayName -like $($Options.Profile) | Select-Object -First 1
        $SelectedGroup = Get-AutopilotProfileAssignments -id $SelectedProfile.id | Select-Object -First 1
        $GroupTag = (Get-MgGroup -groupId $SelectedGroup.Id | Select-Object -ExpandProperty membershipRule).split(':')[1].trim(')','"')
        $OldOptions['Profile'] = $Device.deploymentProfile.DisplayName 
    }
    if ($Options.User) { $OldOptions['User'] = $Device.userPrincipalName }
    if ($Options.Name) { $OldOptions['Name'] = $Device.displayName }

    ## Import into Autopilot
    if (!($Device)) {
        # Flow for new devices
        Write-Host "Importing device with serial '$serial' with GroupTag '$GroupTag'..."
        $hash = Get-CimInstance -Namespace root/cimv2/mdm/dmmap -Class MDM_DevDetail_Ext01 -Filter "InstanceID='Ext' AND ParentID='./DevDetail'"

        $ImportArgs = @{
            serialNumber = $serial
            hardwareIdentifier = $hash.DeviceHardwareData
        }
        if ($GroupTag) { $ImportArgs['groupTag'] = $GroupTag }
        if ($Options.User) { $ImportArgs['assignedUser'] = $Options.User }

        Add-AutopilotImportedDevice @ImportArgs

        ## Wait for import to complete
        do {
            for ($i = 1; $i -le 20; $i++ ) { 
                Write-Progress -Activity "Waiting for import..." -Status "Serial: $serial" -SecondsRemaining (20-$i) 
                Start-Sleep 1
            }
        } until (Get-AutopilotDevice -serial $serial)

        if ($Options.Name) { Set-AutopilotDevice -id (Get-AutopilotDevice -serial $serial).id -computerName $Options.Name }
        Write-Host "Device with serial '$serial' imported successfully! Waiting for details to update..."
    } elseif ((ConvertTo-Json $Options -Depth 3) -ne (ConvertTo-Json $OldOptions -Depth 3)) {
        # Flow for existing devices with new details
        Write-Host "Device with serial '$serial' already imported! Updating attributes instead..."
        $apArgs = @{
            id = $Device.id
        }
        if ($Options.Profile -notlike $OldOptions.Profile) { $apArgs['groupTag'] = $GroupTag }
        if ($Options.User -notlike $OldOptions.User) { $apArgs['userPrincipalName'] = $Options.User }
        if ($Options.Name -notlike $OldOptions.Name) { $apArgs['displayName'] = $Options.Name }

        Set-AutopilotDevice @apArgs
    } else { 
        # Flow for existing devices with same details
        Write-Host "Device with serial '$serial' already imported with specified attributes!" -F Green
        return
    }

    ## Try getting ID again
    do { $Device = Get-AutopilotDevice -serial $serial } until ($Device)
    
    ## Wait for attributes to finish updating
    while ($true) {
        $Device = Get-AutopilotDevice -id (Get-AutopilotDevice -serial $serial).id -expand

        $CurrentOptions = @{}
        if ($Options.Profile) { $CurrentOptions['Profile'] = $Device.deploymentProfile.DisplayName }
        if ($Options.User) { $CurrentOptions['User'] = $Device.userPrincipalName }
        if ($Options.Name) { $CurrentOptions['Name'] = $Device.displayName }

        if ((ConvertTo-Json $Options -Depth 3) -eq (ConvertTo-Json $CurrentOptions -Depth 3)) {
            Write-Host "Device with serial '$serial' successfully imported!" -F Green
            break
        }

        $ProfileStatus = $Device.deploymentProfileAssignmentStatus
        if ($ProfileStatus -eq 'Failed') { Write-Error $Device.deploymentProfileAssignmentDetailedStatus }

        for ($i = 1; $i -le 20; $i++ ) { 
            Write-Progress -Activity "Assigning attributes: $(ConvertTo-Json $Options -Compress)" -Status "Status: $ProfileStatus" -SecondsRemaining (20-$i) 
            Start-Sleep 1
        }
    }
}

## MAIN PROCESS ##

# If you're calling the script through iex(irm), set the graph connection parameters here:
# $TenantId = 'your-tenant-id'
# $ClientId = 'your-client-id'
# $ClientSecret = 'your-client-secret'

ShowInfo
$SetupComplete = ((Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Setup\State' -Name 'ImageState').ImageState -eq 'IMAGE_STATE_COMPLETE')
if ($SetupComplete) { $ErrorActionPreference = 'Stop' } else { $ErrorActionPreference = 'Inquire' }
if ($SetupComplete) { LoadModules } else { LoadModules -TempFolder 'C:\Windows\Temp\TurboPilot' } 
ConnectGraph -SetupComplete $SetupComplete -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret
if ($DeploymentProfile -or $AssignedUser -or $ComputerName) { $Options = OptionsCLI -DeploymentProfile $DeploymentProfile -AssignedUser $AssignedUser -ComputerName $ComputerName }
else { $Options = OptionsGUI }
ImportDevice -Options $Options
Disconnect-MgGraph | Out-Null
if (!($SetupComplete)) { Start-Sleep -Seconds 5 } # Don't restart immediately