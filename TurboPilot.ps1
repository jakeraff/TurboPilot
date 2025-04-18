function Write-HostCenter { param($Message) Write-Host; Write-Host ("{0}{1}" -f (' ' * (([Math]::Max(0, $Host.UI.RawUI.BufferSize.Width / 2) - [Math]::Floor($Message.Length / 2)))), $Message) -f Cyan; Write-Host }

function ShowInfo {
    Write-HostCenter "TurboPilot - Windows Enrollment Script"
    Write-Host "Author: Jacob Raffoul"
    Write-Host "Version: 5.0"
}


function StartGUI {
    Write-HostCenter 'Building main GUI'

    [void] [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")

    Write-Host "Loaded WinForms."

    # Create a new Form
    $form = New-Object Windows.Forms.Form
    $form.Text = 'Confirm options'
    $form.Size = New-Object Drawing.Size(300, 460)
    $form.StartPosition = 'CenterScreen'
    $Form.FormBorderStyle = 'Fixed3D'
    $Form.ControlBox = $false
    $Form.MaximizeBox = $false
    $form.TopMost = $true

    Write-Host "Built frame."

    $checkapv2 = New-Object Windows.Forms.CheckBox
    $checkapv2.Text = "Join Intune - Autopilot V2"
    $checkapv2.ForeColor = [System.Drawing.Color]::DarkGreen
    $checkapv2.Checked = $false
    $checkapv2.Location = New-Object Drawing.Point(20, 70)
    $checkapv2.AutoSize = $true
    $checkapv2.Add_CheckedChanged({
        if ($checkapv2.Checked) {
            $checkremoveapv1.Checked = $true
            $checkapv1.Checked = $false
            $checkreset.Checked = $false
        }
    })

    $checkapv1 = New-Object Windows.Forms.CheckBox
    $checkapv1.Text = "Join Intune - Autopilot V1"
    $checkapv1.ForeColor = [System.Drawing.Color]::DarkGreen
    $checkapv1.Checked = $false
    $checkapv1.Location = New-Object Drawing.Point(20, 100)
    $checkapv1.AutoSize = $true
    $checkapv1.Add_CheckedChanged({
        if ($checkapv1.Checked) {
            $checkapv2.Checked = $false
            $checkunenrolsemm.Checked = $false
            $checkintune.Checked = $false
            $checkremoveapv1.Checked = $false
            $checkreset.Checked = $false
        }
    })

    $checkattributes = New-Object Windows.Forms.CheckBox
    $checkattributes.Text = "Manually upload attributes"
    $checkattributes.ForeColor = [System.Drawing.Color]::DarkGreen
    $checkattributes.Checked = $false
    $checkattributes.Location = New-Object Drawing.Point(20, 130)
    $checkattributes.AutoSize = $true
    $checkattributes.Add_CheckedChanged({
        if ($checkattributes.Checked) {
            $checkunenrolsemm.Checked = $false
            $checkintune.Checked = $false
            $checkremoveapv1.Checked = $false
        }
    })

    $checkenrolsemm = New-Object Windows.Forms.CheckBox
    $checkenrolsemm.Text = "Manually enroll SEMM"
    $checkenrolsemm.ForeColor = [System.Drawing.Color]::DarkGreen
    $checkenrolsemm.Checked = $false
    $checkenrolsemm.Location = New-Object Drawing.Point(20, 160)
    $checkenrolsemm.AutoSize = $true
    $checkenrolsemm.Add_CheckedChanged({
        if ($checkenrolsemm.Checked) {
            $checkunenrolsemm.Checked = $false
        }
    })

    $checkunenrolsemm = New-Object Windows.Forms.CheckBox
    $checkunenrolsemm.Text = "Unenrol from SEMM"
    $checkunenrolsemm.ForeColor = [System.Drawing.Color]::DarkRed
    $checkunenrolsemm.Checked = $false
    $checkunenrolsemm.Location = New-Object Drawing.Point(20, 190)
    $checkunenrolsemm.AutoSize = $true
    $checkunenrolsemm.Add_CheckedChanged({
        if ($checkunenrolsemm.Checked) {
            $checkenrolsemm.Checked = $false
            $checkattributes.Checked = $false
        }
    })

    $checksophos = New-Object Windows.Forms.CheckBox
    $checksophos.Text = "Remove from Sophos"
    $checksophos.ForeColor = [System.Drawing.Color]::DarkRed
    $checksophos.Checked = $false
    $checksophos.Location = New-Object Drawing.Point(20, 220)
    $checksophos.AutoSize = $true

    $checkintune = New-Object Windows.Forms.CheckBox
    $checkintune.Text = "Remove from Intune"
    $checkintune.ForeColor = [System.Drawing.Color]::DarkRed
    $checkintune.Checked = $false
    $checkintune.Location = New-Object Drawing.Point(20, 250)
    $checkintune.AutoSize = $true
    $checkintune.Add_CheckedChanged({
        if ($checkintune.Checked) {
            $checkapv1.Checked = $false
            $checkattributes.Checked = $false
        }
    })

    $checkremoveapv1 = New-Object Windows.Forms.CheckBox
    $checkremoveapv1.Text = "Remove from Autopilot V1"
    $checkremoveapv1.ForeColor = [System.Drawing.Color]::DarkRed
    $checkremoveapv1.Checked = $false
    $checkremoveapv1.Location = New-Object Drawing.Point(20, 280)
    $checkremoveapv1.AutoSize = $true
    $checkremoveapv1.Add_CheckedChanged({
        if ($checkremoveapv1.Checked) {
            $checkapv1.Checked = $false
            $checkattributes.Checked = $false
        }
    })

    $checkreset = New-Object Windows.Forms.CheckBox
    $checkreset.Text = "Start Windows factory reset"
    $checkreset.ForeColor = [System.Drawing.Color]::DarkGoldenrod
    $checkreset.Checked = $false
    $checkreset.Location = New-Object Drawing.Point(20, 310)
    $checkreset.AutoSize = $true
    $checkreset.Add_CheckedChanged({
        if ($checkreset.Checked) {
            $checkapv1.Checked = $false
            $checkapv2.Checked = $false
        }
    })

    $checkrestart = New-Object Windows.Forms.CheckBox
    $checkrestart.Text = "Restart Windows"
    $checkrestart.ForeColor = [System.Drawing.Color]::DarkGoldenrod
    $checkrestart.Checked = $false
    $checkrestart.Location = New-Object Drawing.Point(20, 340)
    $checkrestart.AutoSize = $true

    Write-Host "Added checkboxes."

    $apv2button = New-Object Windows.Forms.Button
    $apv2button.Text = "Autopilot V2"
    $apv2button.Location = New-Object Drawing.Point(10, 20)
    $apv2button.Add_Click({
        $checkapv2.Checked = $true
        $checkapv1.Checked = $false
        $checkattributes.Checked = $false
        $checkenrolsemm.Checked = $false
        $checkunenrolsemm.Checked = $false
        $checksophos.Checked = $false
        $checkintune.Checked = $false
        $checkremoveapv1.Checked = $true
        $checkreset.Checked = $false
    })

    $apv1button = New-Object Windows.Forms.Button
    $apv1button.Text = "Autopilot V1"
    $apv1button.Location = New-Object Drawing.Point(100, 20)
    $apv1button.Add_Click({
        $checkapv2.Checked = $false
        $checkapv1.Checked = $true
        $checkattributes.Checked = $false
        $checkenrolsemm.Checked = $false
        $checkunenrolsemm.Checked = $false
        $checksophos.Checked = $false
        $checkintune.Checked = $false
        $checkremoveapv1.Checked = $false
        $checkreset.Checked = $false
    })

    $deprovbutton = New-Object Windows.Forms.Button
    $deprovbutton.Text = "Deprovision"
    $deprovbutton.Location = New-Object Drawing.Point(190, 20)
    $deprovbutton.Add_Click({
        $checkapv2.Checked = $false
        $checkapv1.Checked = $false
        $checkattributes.Checked = $false
        $checkenrolsemm.Checked = $false
        $checkunenrolsemm.Checked = $true
        $checksophos.Checked = $true
        $checkintune.Checked = $true
        $checkremoveapv1.Checked = $true
        $checkreset.Checked = $true
    })

    Write-Host "Built presets."

    # Create OK Button
    $okButton = New-Object Windows.Forms.Button
    $okButton.Text = "OK"
    $okButton.Location = New-Object Drawing.Point(50, 370)
    $okButton.Add_Click({
        $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.Close()
    })

    # Create Cancel Button
    $cancelButton = New-Object Windows.Forms.Button
    $cancelButton.Text = "Cancel"
    $cancelButton.Location = New-Object Drawing.Point(150, 370)
    $cancelButton.Add_Click({
        $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $form.Close()
    })

    Write-Host "Added buttons."

    $presetlabel = New-Object System.Windows.Forms.Label
    $presetlabel.Location = New-Object System.Drawing.Point(105,0)
    $presetlabel.Size = New-Object System.Drawing.Size(280,20)
    $presetlabel.Text = "Select a preset:"
    $presetlabel.Anchor = 'top'

    $settingslabel = New-Object System.Windows.Forms.Label
    $settingslabel.Location = New-Object System.Drawing.Point(75,55)
    $settingslabel.Size = New-Object System.Drawing.Size(280,20)
    $settingslabel.Text = "Toggle individual sequences:"
    $settingslabel.Anchor = 'top'

    Write-Host "Added labels."

    # Add controls to the form
    $form.Controls.Add($checkapv2)
    $form.Controls.Add($checkapv1)
    $form.Controls.Add($checkattributes)
    $form.Controls.Add($checkenrolsemm)
    $form.Controls.Add($checkunenrolsemm)
    $form.Controls.Add($checksophos)
    $form.Controls.Add($checkintune)
    $form.Controls.Add($checkremoveapv1)
    $form.Controls.Add($checkreset)
    $form.Controls.Add($checkrestart)
    $form.Controls.Add($deprovbutton)
    $form.Controls.Add($apv2button)
    $form.Controls.Add($apv1button)
    $form.Controls.Add($okButton)
    $form.Controls.Add($cancelButton)
    $form.Controls.Add($presetlabel)
    $form.Controls.Add($settingslabel)

    # Show the form as a modal dialog
    $result = $form.ShowDialog()

    if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
        Write-Host "Script exited!" -F Yellow
        Exit 0
    }

    return [PSCustomObject]@{
        JoinAPv2 = $checkapv2.Checked
        JoinAPv1 = $checkapv1.Checked
        GetAttr = $checkattributes.Checked
        EnrolSEMM = $checkenrolsemm.Checked
        UnenrolSEMM = $checkunenrolsemm.Checked
        RemoveSophos = $checksophos.Checked
        RemoveIntune = $checkintune.Checked
        RemoveAPv1 = $checkremoveapv1.Checked
        FactoryReset = $checkreset.Checked
        Restart = $checkrestart.Checked
    }
}

function LoadModules {
    Write-HostCenter 'Loading Modules'

    ## Download modules as .nupkg files and convert to .zip
    $url = ''
    $WebResponse = Invoke-WebRequest -UseBasicParsing -Uri $url
    Write-Host "Found $(($WebResponse.Links.Count) - 1) modules! Downloading them now..."; Write-Host
    $WebResponse.Links | Select-Object -ExpandProperty href -Skip 1 | ForEach-Object {
        Write-Host "Downloading file '$_'"
        $ChildPath = $_ -replace "/Modules/",""
        $filePath = Join-Path -Path $TempFolder -ChildPath $ChildPath
        $Link  = "$url/$ChildPath"
        Invoke-WebRequest -Uri $Link -OutFile $filePath.Replace('.nupkg','.zip')
    }
    Write-Host

    ## Extract and load downloaded modules
    foreach ($item in (Get-ChildItem -Path "$TempFolder\*.zip")) {
        $Module = ($item.BaseName -Replace '\d+','').TrimEnd('.')
        $Path = "$TempFolder\$Module"
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
        try {
            Write-Host -NoNewLine "Loading module: $Module"
            Expand-Archive -Path $item.FullName -DestinationPath $Path -Force
            Write-Host ' | ' -F White -NoNewline; Write-Host 'Extracted!' -NoNewline -F Yellow
            Start-Sleep 2
            Import-Module "$Path\$Module.psd1"
            Write-Host ' | ' -F White -NoNewline; Write-Host 'Imported!' -F Green
            Start-Sleep 2
        } catch { Write-Error "Failed to install module: $Module" }
    }
}

function ConnectGraph {
    ## Connect Graph services
    Write-HostCenter "Connecting to Microsoft Graph"

    ## Connect to Graph API
    $ClientId = ""
    $TenantId = ""
    $ClientSecret = ""

    $ClientSecretPass = ConvertTo-SecureString -String $ClientSecret -AsPlainText -Force
    $ClientSecretCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ClientId, $ClientSecretPass
    Connect-MgGraph -TenantId $tenantId -ClientSecretCredential $ClientSecretCredential
}

function APv1GUI {
    ## Build provisioning GUI
    Write-HostCenter 'Building Provisioning GUI'

    [void] [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")

    Write-Host "Loaded WinForms."

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'TurboPilot - Autopilot V1'
    $form.Size = New-Object System.Drawing.Size(200,150)
    $form.StartPosition = 'CenterScreen'
    $Form.FormBorderStyle = 'Fixed3D'
    $Form.ControlBox = $false
    $Form.MaximizeBox = $false
    $form.TopMost = $true

    Write-Host "Built frame."

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(20,70)
    $okButton.Size = New-Object System.Drawing.Size(50,25)
    $okButton.Anchor = 'left'
    $okButton.Text = 'OK'
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $okButton
    $form.Controls.Add($okButton)

    Write-Host "Added OK button."

    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(22,5)
    $label.Size = New-Object System.Drawing.Size(280,20)
    $label.Text = "Deployment profile:"
    $label.Anchor = 'top'
    $form.Controls.Add($label)

    Write-Host "Added list label."

    $List = New-Object system.Windows.Forms.ComboBox
    $List.text = ""
    $List.Location = New-Object System.Drawing.Point(20,30)
    $List.width = 142
    $List.Anchor = 'top'

    ## Gets list of all AP profiles and adds them to the drop down
    $Profiles = Get-AutopilotProfile -Verbose
    [array]::Reverse($Profiles)
    foreach ($Profile in $Profiles) { [void] $List.Items.Add("$($Profile.displayName)") }
    Write-Host "Downloaded Autopilot profiles."

    ## Gets the current profile assigned to the device (if any) and makes it the default selection
    $serial = (Get-ComputerInfo -Property BiosSeralNumber -Verbose | Select-Object -ExpandProperty BiosSeralNumber)
    $APID = (Get-AutopilotDevice -serial $serial -Verbose).id
    Write-Host "Found Autopilot ID."
    if ($Null -ne $APID) { $CurrentProfile = (Get-AutopilotDevice -id $APID -Expand -Verbose).deploymentProfile.DisplayName }
    Write-Host "Found current profile for device."
    if ($CurrentProfile) { $List.SelectedItem = $CurrentProfile }
    else { $List.SelectedIndex = 0 }

    $form.Controls.Add($List)

    $result = $form.ShowDialog()
    
    if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
        Write-Host "Script exited!" -B Red -F White
        timeout /t -1
        Exit 0
    }

    return $List.SelectedItem
}

function ImportDevice {
    Param (
        $SelectedItem
    )

    Write-HostCenter 'Processing AutoPilot Request'

    ## Find GroupTag
    $serial = (Get-ComputerInfo -Property BiosSeralNumber -Verbose | Select-Object -ExpandProperty BiosSeralNumber)
    $APID = (Get-AutopilotDevice -serial $serial).id
    Write-Host "Found Autopilot ID."
    if ($APID) { $CurrentProfile = (Get-AutopilotDevice -id $APID -Expand -Verbose).deploymentProfile.DisplayName }
    Write-Host "Found current profile for device."
    $SelectedProfile = Get-AutopilotProfile -Verbose | Where-Object displayName -like $SelectedItem
    $SelectedGroup = Get-AutopilotProfileAssignments -id $SelectedProfile.id -Verbose
    $GroupTag = (Get-MgGroup -groupId $SelectedGroup.Id -Verbose | Select-Object -ExpandProperty membershipRule).split(':')[1].trim(')','"')

    while (!($GroupTag)) {
        $GroupTag = [Microsoft.VisualBasic.Interaction]::InputBox(
            "No GroupTag found! Please manually input one", #Description
            "TurboPilot" #Title
        )
    }

    ## Import into Autopilot
    if (!($APID)) {
        Write-Host "Importing device with GroupTag '$GroupTag'..."
        $hash = Get-CimInstance -Namespace root/cimv2/mdm/dmmap -Class MDM_DevDetail_Ext01 -Filter "InstanceID='Ext' AND ParentID='./DevDetail'"
        Add-AutopilotImportedDevice -serialNumber $serial -hardwareIdentifier $hash.DeviceHardwareData -groupTag $GroupTag    
        
        ## Wait for import to complete
        do {
            for ($i = 1; $i -le 20; $i++ ) { 
                Write-Progress -Activity "Waiting for import..." -Status "Serial: $serial" -SecondsRemaining (20-$i) 
                Start-Sleep 1
            }
        } until (Get-AutopilotDevice -serial $serial)
    } elseif (($NULL -ne $CurrentProfile) -and ("$SelectedItem" -ne "$CurrentProfile")) {
        ## Sets the groupTag
        Write-Host "Device already imported. Setting groupTag instead."
        Set-AutopilotDevice -id $APID -groupTag $GroupTag -Verbose
    } else { Write-Host "Device already imported with correct groupTag!" }

    ## Try getting ID again
    do { $APID = (Get-AutopilotDevice -serial $serial).id }
    until ($APID)
    
    ## Wait for profile assignment to finish
    while ("$AssignedProfile" -notlike "$SelectedItem") {
        $AssignedProfile = (Get-AutopilotDevice -id $APID -Expand).deploymentProfile.DisplayName
        for ($i = 1; $i -le 20; $i++ ) { 
            Write-Progress -Activity "Assigning profile: $SelectedItem" -Status "Current profile: $AssignedProfile" -SecondsRemaining (20-$i) 
            Start-Sleep 1
        }
    }

    Write-Host
    Write-Host "imported successfully!" -F Green
}

function GetAttributes {
    Write-HostCenter "Collecting Azure Extension Attributes"

    [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null

    ## Prompt for asset and contract until at least asset has been provided
    do { 
        if (!(Get-MgContext)) { ConnectGraph }
        
        $Asset = [int] [Microsoft.VisualBasic.Interaction]::InputBox(
            "Please input the 'Asset Number' (e.g. 123456):", #Description
            "TurboPilot" #Title
        )

        $Contract = [Microsoft.VisualBasic.Interaction]::InputBox(
            "Please input the 'Contract Number' (e.g. 12). Leave blank if not applicable:", #Description
            "TurboPilot" #Title
        )
    } until ($Asset)
    if (!($Contract)) { $Contract = 'NA' }

    Write-Host "Uploading provided attributes..."

    $serial = (Get-ComputerInfo -Property BiosSeralNumber -Verbose | Select-Object -ExpandProperty BiosSeralNumber)
    $AzureID = (Get-AutopilotDevice -serial $serial).azureAdDeviceId
    $IDs = Get-MgDevice -Filter "DeviceID eq '$AzureID'" | Select-Object -ExpandProperty Id
    
    ## Build Graph extension attribute request for each device found
    foreach ($Id in $IDs) {
        $uri = $null
        $uri = "https://graph.microsoft.com/beta/devices/" + "$Id"

        $json = @{
        "extensionAttributes" = @{
            "extensionAttribute1" = "$($Details.Asset)"
            "extensionAttribute2" = "$($Details.Contract)"
        }
        } | ConvertTo-Json

        try {
            Invoke-MgGraphRequest -Uri $uri -Body $json -Method PATCH -ContentType "application/json" -Verbose
            Write-Host "Successfully set attributes asset $($Details.Asset) and contract $($Details.Contract)." -ForegroundColor Green
        }
        catch { 
            Write-Error `
                -Exception "Failed to upload attributes!" `
                -RecommendedAction "Please use the Update-ExtensionAttributes.ps1 script in the SharePoint File Library."
        }
    }
}

function FindIDs {
    $AzureDeviceId = (("$((dsregcmd /status) -match 'DeviceID')" -Split ":" ).trim())[1]
    #$AzureObjectId = Get-MgDevice -Filter "DeviceId eq '$AzureDeviceId'" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id
    $AutopilotId = Get-MgDeviceManagementWindowsAutopilotDeviceIdentity -Filter "AzureActiveDirectoryDeviceId eq '$AzureDeviceId'" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id
    $IntuneDeviceId = Get-MgDeviceManagementWindowsAutopilotDeviceIdentity -Filter "AzureActiveDirectoryDeviceId eq '$AzureDeviceId'" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty ManagedDeviceId    

    if ([String]::IsNullOrWhiteSpace($AzureDeviceId)) { throw "Azure Device ID missing from DSREGCMD!"}
    #if ([String]::IsNullOrWhiteSpace($AzureObjectId)) { throw "Can't find Azure Object ID from DSREGCMD Device ID!"}

    $IDCollection = [PSCustomObject]@{
        AzureDeviceId = $AzureDeviceId
        #AzureObjectId = $AzureObjectId
        AutopilotId = $AutopilotId
        IntuneDeviceId = $IntuneDeviceId
    }

    Write-Host "Successfully found IDs!" -F Green
    Write-Host $IDCollection

    return $IDCollection
}

function RemoveIntune {

    Write-HostCenter "Removing from Intune..."

    $AzureDeviceId = (("$((dsregcmd /status) -match 'DeviceID')" -Split ":" ).trim())[1]
    if ([String]::IsNullOrWhiteSpace($AzureDeviceId)) { throw "Azure Device ID missing from DSREGCMD!"}
    $IntuneDeviceId = Get-MgDeviceManagementWindowsAutopilotDeviceIdentity -Filter "AzureActiveDirectoryDeviceId eq '$AzureDeviceId'" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty ManagedDeviceId    
    if ([String]::IsNullOrWhiteSpace($IntuneDeviceId)) { throw "Can't find Intune Device ID from DSREGCMD Device ID!"}

    try {
        Remove-MgDeviceManagementManagedDevice -ManagedDeviceId $IntuneDeviceId -ErrorAction Stop
    }
    catch [System.AggregateException] {
        Write-Host "Caught an aggregate exception:"
    
        # This will enumerate over any inner exceptions and print their details
        foreach($innerEx in $_.Exception.InnerExceptions) {
            Write-Host "=== Inner Exception ===" -F Red
            Write-Host $innerEx.Message -F Red
            Write-Host $innerEx.StackTrace -F Red
        }
    }
    do {
        for ($i = 1; $i -le 20; $i++ ) { 
            Write-Progress -Status "Device still in Intune!" -Activity "Waiting for deletion..." -SecondsRemaining (20-$i) 
            Start-Sleep 1
        }
    } while (Get-MgDeviceManagementManagedDevice -ManagedDeviceId $IntuneDeviceId -ErrorAction SilentlyContinue)
    Write-Host "Successfully removed from Intune!" -F Green
}

function RemoveAutopilot {

    Write-HostCenter "Removing from Autopilot (and Azure)..."

    $serial = Get-ComputerInfo -Property BiosSeralNumber | Select-Object -ExpandProperty BiosSeralNumber
    $AutopilotId = (Get-AutopilotDevice -serial $serial).id
    
    if ([String]::IsNullOrWhiteSpace($AutopilotId)) { throw "Can't find Autopilot ID!"}

    Write-Host "Found Autopilot ID: $AutopilotId"

    Remove-AutopilotDevice -id $AutopilotId -ErrorAction Stop

    do {
        for ($i = 1; $i -le 20; $i++ ) { 
            Write-Progress -Status "Device still in autopilot!" -Activity "Waiting for deletion..." -SecondsRemaining (20-$i) 
            Start-Sleep 1
        }
    } while ((Get-AutopilotDevice -serial $serial).id)
    Write-Host "Successfully removed from Autopilot!" -F Green
}

function Connect-Sophos {
    $client_id = ""
    $client_secret = ""

    Write-HostCenter "Connecting to Sophos..."

    # authenticate with sophos (returns time/scope limited java web token)
    $token_resp = Invoke-RestMethod -Method Post -ContentType "application/x-www-form-urlencoded" -Body "grant_type=client_credentials&client_id=$client_id&client_secret=$client_secret&scope=token" -Uri https://id.sophos.com/api/v2/oauth2/token -Verbose
    Write-Host "Established session!"
    $whoami_resp = Invoke-RestMethod -Method Get -Headers @{Authorization="Bearer $($token_resp.access_token)"} https://api.central.sophos.com/whoami/v1
    Write-Host "Retrieved service principal."
    $dataRegionApiUri = $whoami_resp.apiHosts.dataRegion
    Write-Host "Found URI with region."
    # Get all endpoints within the tenant
    $endpoints_resp = Invoke-RestMethod -Method Get -Headers @{Authorization="Bearer $($token_resp.access_token)"; "X-Tenant-ID"=$whoami_resp.id} ($($whoami_resp.apiHosts.dataRegion)+"/endpoint/v1/endpoints")
    return [PSCustomObject]@{
        token_resp = $token_resp
        whoami_resp = $whoami_resp
        endpoints_resp = $endpoints_resp
        sophosApiResponse = $dataRegionApiUri
    }
}

function Get-SophosEndpoints {
    param (
        [Parameter(Mandatory=$true)]
        $sophosApiResponse
    )

    $endpoint_key = $sophosApiResponse.endpoints_resp.pages.nextKey
    
    $sophosEndpoints = @()

    Write-HostCenter "Finding Endpoint ID of current host..."
    $Page = 0
    Do {
        $endpoints_resp = Invoke-RestMethod -Method Get -Headers @{Authorization="Bearer $($sophosApiResponse.token_resp.access_token)"; "X-Tenant-ID"=$sophosApiResponse.whoami_resp.id} ($($sophosApiResponse.sophosApiResponse)+"/endpoint/v1/endpoints?pageSize=500&pageTotal=true&pageFromKey=$($endpoint_key)") -ErrorAction Stop -Verbose

        # enumerate results and append to csv
        $sophosEndpoints += @($endpoints_resp.items | 
            Select-Object -Property id,type,hostname,health,os,
            @{name="ipv4Addresses"; expression={$_.ipv4Addresses | Select-Object -First 1}},
            @{name="ipv6Addresses"; expression={$_.ipv6Addresses | Select-Object -First 1}},
            @{name="macAddresses"; expression={$_.macAddresses | Select-Object -First 1}},
            associatedPerson,tamperProtectionEnabled,
            @{name="endpointProtection"; expression={$_.assignedProducts[0] | Where-Object -Property code -eq -Value "endpointProtection"}},
            @{name="interceptX"; expression={$_.assignedProducts[1] | Where-Object -Property code -eq -Value "interceptX"}},
            @{name="coreAgent"; expression={$_.assignedProducts[2] | Where-Object -Property code -eq -Value "coreAgent"}},
            lastSeenAt)
            
        Start-Sleep -Seconds 2
        $endpoint_key = $endpoints_resp.pages.nextKey

        $Page ++
        Write-Progress -Activity "Collating paginated list of endpoints" -Status "Added page $Page!"
    } 
    While ($endpoints_resp.pages.fromKey -ne "")
    Write-Progress -Activity "Sophos API" -Completed
    Write-Host "Got list of all endpoints in the tenant!"

    return ($sophosEndpoints | Sort-Object "hostname")
}

### function ###
function Get-SophosEndpointId {
    param (
        [Parameter(Mandatory=$true)]
        [String] 
        $computerName
        ,
        [Parameter(Mandatory=$true)]
        $sophosApiResponse
    )
    
    $sophosEndpoints = Get-SophosEndpoints -sophosApiResponse $sophosApiResponse

    $EndpointIDs = @()

    # loop through all devices on sophos to find matching id for current device
    foreach ($endpoint in $sophosEndpoints) {
        if($computerName -eq $endpoint.hostname) {
            $EndpointIDs += $endpoint.id
        }
    }
    if ($EndpointIDs.count -ne 0) {
        Write-Host "Found $($EndpointIDs.count) endpoints with matching IDs!"
        return ($EndpointIDs | Select-Object -Unique)
    }
    else {
        Write-Host "No endpoints found with name $computerName." -F Red
    }
}

function Remove-SophosEndpoint {
    param (
        [Parameter(Mandatory=$true)]
        $EndpointIDs
        ,
        [Parameter(Mandatory=$true)]
        $sophosApiResponse
    )
    foreach ($EndpointID in $EndpointIDs) {
        Write-Host "Removing endpoint with ID $EndpointID"
        Invoke-RestMethod -Method Delete -Headers @{Authorization="Bearer $($sophosApiResponse.token_resp.access_token)"; "X-Tenant-ID"=$sophosApiResponse.whoami_resp.id} ($($sophosApiResponse.sophosApiResponse)+"/endpoint/v1/endpoints/$EndpointID") -ErrorAction Stop -Verbose
    }
}

function LoadSEMM {
    Write-HostCenter "Enabling SEMM Enrollment"

    if ((Get-ComputerInfo -Property CsModel | Select-Object -ExpandProperty CsModel) -notlike "*Surface*") { return }
    Write-Host "Device is eligible for SEMM enrollment. Initializing now..."

    ## Load dependencies and enroll into SEMM
    Invoke-WebRequest '' -OutFile "$TempFolder\UEFI.pfx" -Verbose
    Invoke-WebRequest '' -OutFile "$TempFolder\SEMM.msi" -Verbose
    Start-Process msiexec.exe -ArgumentList "/i $TempFolder\SEMM.msi /passive /norestart" -Wait
    while (((get-process) -like "*msiexec*").count -ge 2) { start-sleep 3 }

    Invoke-WebRequest '/semm/ConfigureSEMM.ps1' -OutFile "$TempFolder\ConfigureSEMM.ps1" -Verbose
    powershell.exe -ep bypass -File "$TempFolder\ConfigureSEMM.ps1"
}

function ResetSEMM {
    Write-HostCenter 'Downloading Surface UEFI Manager...'

    if ((Get-ComputerInfo -Property CsModel | Select-Object -ExpandProperty CsModel) -notlike "*Surface*") { 
        Write-Host "This is not a Surface device, can't install UEFI manager!" -F Red
        return 
    }

    Invoke-WebRequest '/semm/SurfaceUEFI_Manager_v2.198.139.0_x64.msi' -OutFile 'C:\Windows\Temp\SEMM.msi' -Verbose
    Start-Process msiexec.exe -ArgumentList "/i C:\Windows\Temp\SEMM.msi /passive /norestart" -Wait
    while (((get-process) -like "*msiexec*").count -ge 2) { start-sleep 3 }

    Write-HostCenter "Downloading SEMM configuration..."
    Invoke-WebRequest '/semm/rgUEFI.pfx' -OutFile 'C:\Windows\Temp\rgUEFI.pfx' -Verbose
    Invoke-WebRequest '/semm/ResetSemm.ps1' -OutFile 'C:\Windows\Temp\ResetSemm.ps1' -Verbose
    powershell.exe -ep bypass -File 'C:\Windows\Temp\ResetSemm.ps1'

    Remove-Item 'C:\Windows\Temp\rgUEFI.pfx' -Force -Verbose
    Remove-Item 'C:\Windows\Temp\ResetSemm.ps1' -Force -Verbose
    Remove-Item 'C:\Windows\Temp\SEMM.msi' -Force -Verbose
}

## MAIN PROCESS ##

ShowInfo

#$ErrorActionPreference = 'Inquire'
$Global:TempFolder = 'C:\Windows\Temp\'
New-Item -ItemType Directory $TempFolder -Force | Out-Null

$Options = (StartGUI)
Write-HostCenter "Selected Options"
($Options | Format-List | Out-String).Trim()

if ($Options.JoinAPv1 -or $Options.RemoveAPv1 -or $Options.RemoveIntune) { LoadModules; ConnectGraph}

if ($Options.JoinAPv1) { ImportDevice -SelectedItem (APv1GUI) }

if ($Options.GetAttr) { GetAttributes }

if ($Options.EnrolSEMM) { LoadSEMM }

if ($Options.UnenrolSEMM) { ResetSEMM }

if ($Options.RemoveSophos) {
    $SophosSession = (Connect-Sophos)
    $SophosEndpointIDs = Get-SophosEndpointId -computerName (hostname) -sophosApiResponse $SophosSession
    if ($SophosEndpointIDs) { Remove-SophosEndpoint -EndpointIDs $SophosEndpointIDs -sophosApiResponse $SophosSession }
}

if ($Options.RemoveIntune) { RemoveIntune }

if ($Options.RemoveAPv1) { RemoveAutopilot }

if ($Options.JoinAPv1 -or $Options.RemoveAPv1 -or $Options.RemoveIntune) { Disconnect-MgGraph -ErrorAction Ignore | Out-Null }

Remove-Item -Path $TempFolder -Force -Recurse -ErrorAction 'SilentlyContinue' | Out-Null

if ($Options.FactoryReset) { Write-HostCenter "Initiating Factory Reset..."; systemreset --factoryreset }

if ($Options.Restart) { shutdown /r /t 0 }