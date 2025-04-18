function Write-HostCenter { param($Message) Write-Host; Write-Host ("{0}{1}" -f (' ' * (([Math]::Max(0, $Host.UI.RawUI.BufferSize.Width / 2) - [Math]::Floor($Message.Length / 2)))), $Message) -f Cyan; Write-Host }

function ShowInfo {
    Write-HostCenter "TurboPilot - Windows Enrollment Script"
    Write-Host "Author: Jacob Raffoul"
    Write-Host "Version: 1.0.0"
    Write-Host "License: GPL-3.0" 
}

function LoadModules {
    Write-HostCenter 'Loading Modules'

    New-Item -ItemType Directory $TempFolder -Force | Out-Null
    Remove-Item -Path "$TempFolder/*" -Recurse -Force | Out-Null

    $Links = @(
        "https://psg-prod-eastus.azureedge.net/packages/microsoft.graph.authentication.2.26.1.nupkg"
        "https://psg-prod-eastus.azureedge.net/packages/microsoft.graph.users.2.26.1.nupkg"
        "https://psg-prod-eastus.azureedge.net/packages/microsoft.graph.groups.2.26.1.nupkg"
        "https://psg-prod-eastus.azureedge.net/packages/microsoft.graph.intune.6.1907.1.nupkg"
        "https://psg-prod-eastus.azureedge.net/packages/windowsautopilotintune.5.7.0.nupkg"
    )

    ## Extract and load downloaded modules
    foreach ($Link in $Links) {
        try {
            $Module = $Link.split('/').trim('.nupkg') | Select-Object -Last 1
            $Path = "$TempFolder\$Module.zip"
            Write-Host -NoNewLine "Loading module: $Module"
            Invoke-WebRequest -Uri $Link -OutFile $Path
            Write-Host ' | ' -F White -NoNewline; Write-Host 'Downloaded!' -NoNewline -F White
            $Folder = New-Item -ItemType Directory $Path.Replace('.zip','')
            Expand-Archive -Path $Path -DestinationPath $Folder -Force
            Write-Host ' | ' -F White -NoNewline; Write-Host 'Extracted!' -NoNewline -F Yellow
            Start-Sleep 2
            Import-Module (Get-ChildItem $Folder -Filter '*.psd1' | Select-Object -ExpandProperty FullName)
            Write-Host ' | ' -F White -NoNewline; Write-Host 'Imported!' -F Green
            Start-Sleep 2
        } catch { Write-Error "Failed to install module: $Module" }
    }
}

function ConnectGraph {
    param(
        [String]$TenantId,
        [String]$ClientId,
        [String]$ClientSecret
    )

    $Scopes = @(
        'Group.ReadWrite.All',
        'Device.ReadWrite.All',
        'DeviceManagementManagedDevices.ReadWrite.All',
        'DeviceManagementServiceConfig.ReadWrite.All',
        'GroupMember.ReadWrite.All',
        'User.Read.All'
    )
    
    ## Connect Graph services
    Write-HostCenter "Connecting to Microsoft Graph"

    if ($TenantId -and $ClientId -and $ClientSecret) {
        ## Connect to Graph API silently
        $ClientSecretPass = ConvertTo-SecureString -String $ClientSecret -AsPlainText -Force
        $ClientSecretCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ClientId, $ClientSecretPass
        Connect-MgGraph -TenantId $tenantId -ClientSecretCredential $ClientSecretCredential
    } else {
        ## Connect interactively
        Connect-MgGraph -Scopes $Scopes
    }
}

function ProfileGUI {
    ## Build provisioning GUI
    Write-HostCenter 'Building Provisioning GUI'

    [void] [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")

    Write-Host "Loaded WinForms."

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'TurboPilot - Autopilot V1'
    $form.Size = New-Object System.Drawing.Size(200,250)
    $form.StartPosition = 'CenterScreen'
    $Form.FormBorderStyle = 'Fixed3D'
    $Form.ControlBox = $false
    $Form.MaximizeBox = $false
    $form.TopMost = $true

    Write-Host "Built frame."

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Size = New-Object System.Drawing.Size(50,25)
    $okButton.Location = New-Object System.Drawing.Point(20, ($form.ClientSize.Height - 50)) # Position near the bottom
    $okButton.Anchor = 'Bottom'
    $okButton.Text = 'OK'
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $okButton
    $form.Controls.Add($okButton)

    Write-Host "Added OK button."

    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(22,0)
    $label.Size = New-Object System.Drawing.Size(280,20)
    $label.Text = "Deployment profile:"
    $label.Anchor = 'top'
    $form.Controls.Add($label)

    Write-Host "Added profile list label."

    $ProfileList = New-Object system.Windows.Forms.ComboBox
    $ProfileList.text = ""
    $ProfileList.Location = New-Object System.Drawing.Point(20,20)
    $ProfileList.width = 142
    $ProfileList.Anchor = 'top'

    ## Gets list of all AP profiles and adds them to the drop down
    $Profiles = Get-AutopilotProfile
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
    $UserLabel.Location = New-Object System.Drawing.Point(22,50)
    $UserLabel.Size = New-Object System.Drawing.Size(280,20)
    $UserLabel.Text = "Assigned User:"
    $UserLabel.Anchor = 'top'
    $form.Controls.Add($UserLabel)

    Write-Host "Added assigned user label."

    $UserInput = New-Object System.Windows.Forms.TextBox
    $UserInput.text = ""
    $UserInput.Location = New-Object System.Drawing.Point(20,70)
    $UserInput.Width = 142
    $UserInput.Anchor = 'top'

    # Get Users
    $Users = Get-MgUser -Filter "accountEnabled eq true" -Property UserPrincipalName -All | Select-Object UserPrincipalName
    $UserInput.AutoCompleteSource = 'CustomSource' 
    $UserInput.AutoCompleteMode = 'SuggestAppend'
    $AutoComplete = New-Object System.Windows.Forms.AutoCompleteStringCollection
    $UserInput.AutoCompleteCustomSource = $AutoComplete
    $AutoComplete.AddRange($Users.UserPrincipalName)

    # Pre-fill the input box with the current assigned user if available
    $serial = (Get-ComputerInfo -Property BiosSeralNumber | Select-Object -ExpandProperty BiosSeralNumber)
    $CurrentUser = (Get-AutopilotDevice -serial $serial).userPrincipalName
    if ($CurrentUser) {
        $UserInput.Text = $CurrentUser
        Write-Host "Pre-filled assigned user: $CurrentUser"
    }

    $form.Controls.Add($UserInput)

    Write-Host "Added assigned user field."

    $NameLabel = New-Object System.Windows.Forms.Label
    $NameLabel.Location = New-Object System.Drawing.Point(22,100)
    $NameLabel.Size = New-Object System.Drawing.Size(280,20)
    $NameLabel.Text = "Computer Name:"
    $NameLabel.Anchor = 'top'
    $form.Controls.Add($NameLabel)

    Write-Host "Added computer name label."

    $NameInput = New-Object System.Windows.Forms.TextBox
    $NameInput.text = (hostname)
    $NameInput.Location = New-Object System.Drawing.Point(20,120)
    $NameInput.width = 142
    $NameInput.Anchor = 'top'
    $form.Controls.Add($NameInput)

    $result = $form.ShowDialog()
    
    if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
        Write-Host "Script exited!" -B Red -F White
        timeout /t -1
        Exit 0
    }

    return [PSCustomObject]@{
        Profile = $ProfileList.SelectedItem
        User = $UserInput.Text
        Name = $NameInput.Text
    }
}

function ImportDevice {
    param (
        [PSCustomObject]$Options
    )

    Write-HostCenter 'Processing AutoPilot Request'

    ## Find GroupTag
    $serial = (Get-ComputerInfo -Property BiosSeralNumber | Select-Object -ExpandProperty BiosSeralNumber)
    $Device = Get-AutopilotDevice -id (Get-AutopilotDevice -serial $serial).id -expand

    $OldOptions = [PSCustomObject]@{
        Profile = $Device.deploymentProfile.DisplayName
        User = $Device.userPrincipalName
        Name = $Device.displayName
    }

    $SelectedProfile = Get-AutopilotProfile | Where-Object displayName -like $($Options.Profile)
    $SelectedGroup = Get-AutopilotProfileAssignments -id $SelectedProfile.id
    $GroupTag = (Get-MgGroup -groupId $SelectedGroup.Id | Select-Object -ExpandProperty membershipRule).split(':')[1].trim(')','"')

    while (!($GroupTag)) {
        $GroupTag = [Microsoft.VisualBasic.Interaction]::InputBox(
            "No GroupTag found! Please manually input one", #Description
            "TurboPilot" #Title
        )
    }

    ## Import into Autopilot
    if (!($Device)) {
        Write-Host "Importing device with serial '$serial' with GroupTag '$GroupTag'..."
        $hash = Get-CimInstance -Namespace root/cimv2/mdm/dmmap -Class MDM_DevDetail_Ext01 -Filter "InstanceID='Ext' AND ParentID='./DevDetail'"

        $ImportArgs = @{
            serialNumber = $serial
            hardwareIdentifier = $hash.DeviceHardwareData
            groupTag = $GroupTag
        }
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
    } elseif ($Options -notlike $OldOptions) {
        Write-Host "Device with serial '$serial' already imported! Updating attributes instead..."
        $apArgs = @{
            id = $Device.id
        }
        if ($Options.Profile -notlike $OldOptions.Profile) { $apArgs['groupTag'] = $GroupTag }
        if ($Options.User -notlike $OldOptions.User) { $apArgs['userPrincipalName'] = $Options.User }
        if ($Options.Name -notlike $OldOptions.Name) { $apArgs['displayName'] = $Options.Name }

        Set-AutopilotDevice @apArgs
    } else { 
        Write-Host "Device with serial '$serial' already imported with specified attributes!" -F Green
        return
    }

    ## Try getting ID again
    do { $Device = Get-AutopilotDevice -serial $serial } until ($Device)
    
    ## Wait for attributes to finish updating
    while ($true) {
        $Device = Get-AutopilotDevice -id (Get-AutopilotDevice -serial $serial).id -expand

        $CurrentOptions = [PSCustomObject]@{
            Profile = $Device.deploymentProfile.DisplayName
            User = $Device.userPrincipalName
            Name = $Device.displayName
        }

        if ($CurrentOptions -like $Options) {
            Write-Host "Device with serial '$serial' successfully imported!" -F Green
            break
        }

        for ($i = 1; $i -le 20; $i++ ) { 
            Write-Progress -Activity "Assigning attributes: $Options" -Status "Current attributes: $CurrentOptions" -SecondsRemaining (20-$i) 
            Start-Sleep 1
        }
    }
}

## MAIN PROCESS ##

$ErrorActionPreference = 'Inquire'
$Global:TempFolder = 'C:\Windows\Temp\TurboPilot'

ShowInfo
LoadModules
ConnectGraph
ImportDevice -Options (ProfileGUI)
Disconnect-MgGraph -ErrorAction Ignore | Out-Null
Remove-Item -Path $TempFolder -Force -Recurse | Out-Null