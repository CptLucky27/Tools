# Load Windows Forms and Exchange Online module
Add-Type -AssemblyName System.Windows.Forms
Import-Module ExchangeOnlineManagement

# Connect to Exchange Online
Connect-ExchangeOnline

# Fetch the list of public folders
$publicFolders = Get-PublicFolder -Recurse -ResultSize Unlimited

# Main Form Configuration
$form = New-Object System.Windows.Forms.Form
$form.Text = "Public Folder Permissions Management"
$form.Size = New-Object System.Drawing.Size(800, 600)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Calibri", 10)

# TableLayoutPanel Configuration
$tableLayoutPanel = New-Object System.Windows.Forms.TableLayoutPanel
$tableLayoutPanel.RowCount = 3
$tableLayoutPanel.ColumnCount = 1
$tableLayoutPanel.Dock = 'Fill'
$form.Controls.Add($tableLayoutPanel)

# Set fixed heights for the rows
$tableLayoutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$tableLayoutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$tableLayoutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))

# Create the Search Label
$searchLabel = New-Object System.Windows.Forms.Label
$searchLabel.Text = "Search:"
$searchLabel.Dock = 'Fill'
$searchLabel.Margin = New-Object System.Windows.Forms.Padding(5, 5, 0, 0) # Adjusted top and left padding
$tableLayoutPanel.Controls.Add($searchLabel, 0, 0)

# Create the Search Box
$searchBox = New-Object System.Windows.Forms.TextBox
$searchBox.Dock = 'Fill'
$searchBox.Margin = New-Object System.Windows.Forms.Padding(5, 5, 5, 5) # Adjusted all margins
$tableLayoutPanel.Controls.Add($searchBox, 0, 1)

# Listbox for Public Folders
$listBox = New-Object System.Windows.Forms.ListBox
$listBox.Dock = 'Fill'
$listBox.IntegralHeight = $false
$listBox.Margin = New-Object System.Windows.Forms.Padding(5)
$tableLayoutPanel.Controls.Add($listBox, 0, 2)

# Initially populate the list box
$publicFolders | ForEach-Object {
    $listBox.Items.Add($_.Identity)
}

# Search functionality
$searchBox.Add_TextChanged({
    $listBox.BeginUpdate()
    $listBox.Items.Clear()
    $publicFolders | Where-Object { $_.Identity -like "*$($searchBox.Text)*" } | ForEach-Object {
        $listBox.Items.Add($_.Identity)
    }
    $listBox.EndUpdate()
})

# Checkbox for inheritance
$checkBox = New-Object System.Windows.Forms.CheckBox
$checkBox.Text = "Include inheritance"
$checkBox.Dock = 'Bottom'
$form.Controls.Add($checkBox)

# Panel for User/Group Email Input and Permission Level Dropdown
$panel = New-Object System.Windows.Forms.FlowLayoutPanel
$panel.Dock = 'Bottom'
$panel.AutoSize = $true
$form.Controls.Add($panel)

# Label for User/Group Email Input
$userGroupLabel = New-Object System.Windows.Forms.Label
$userGroupLabel.Text = "Enter user or group email:"
$userGroupLabel.AutoSize = $true # Auto-size label
$panel.Controls.Add($userGroupLabel)

# Textbox for User/Group Email input (used for both applying and removing permissions)
$userGroupTextBox = New-Object System.Windows.Forms.TextBox
$userGroupTextBox.Width = 200
$panel.Controls.Add($userGroupTextBox)

# Dropdown for Permission Level
$permissionDropdown = New-Object System.Windows.Forms.ComboBox
$permissionDropdown.Width = 200
$permissionDropdown.DropDownStyle = 'DropDownList'
$permissionDropdown.Items.AddRange(@("Owner", "PublishingEditor", "Editor", "PublishingAuthor", "Author", "NonEditingAuthor", "Reviewer", "Contributor", "None"))
$permissionDropdown.SelectedIndex = 0
$panel.Controls.Add($permissionDropdown)

# Function to add permissions if not already present
function Add-PermissionIfNotExists {
    param($folderIdentity, $user, $rights)
    $existingPermissions = Get-PublicFolderClientPermission -Identity $folderIdentity
    $userPermission = $existingPermissions | Where-Object {
        $_.User -and 
        $_.User.UserId -and 
        $_.User.UserId.ToString() -eq $user
    }
    if ($userPermission -and $userPermission.AccessRights -contains $rights) {
        return # Permissions already exist, no need to add
    }
    Add-PublicFolderClientPermission -Identity $folderIdentity -User $user -AccessRights $rights
}

# Button for Applying Permissions
$applyButton = New-Object System.Windows.Forms.Button
$applyButton.Text = "Apply"
$applyButton.Dock = 'Bottom'
$applyButton.Add_Click({
    try {
        $selectedFolder = $listBox.SelectedItem
        if (-not $selectedFolder) {
            [System.Windows.Forms.MessageBox]::Show("Please select a public folder from the list.", "Input Error")
            return
        }
        $userOrGroup = $userGroupTextBox.Text
        $permissionLevel = $permissionDropdown.SelectedItem.ToString()
        if ([string]::IsNullOrWhiteSpace($userOrGroup)) {
            [System.Windows.Forms.MessageBox]::Show("Please enter a user or group email.", "Input Error")
            return
        }
        Add-PermissionIfNotExists $selectedFolder $userOrGroup $permissionLevel
        if ($checkBox.Checked) {
            $childFolders = Get-PublicFolder -Recurse -Identity $selectedFolder
            foreach ($childFolder in $childFolders) {
                if ($childFolder.Identity -ne $selectedFolder) { 
                    Add-PermissionIfNotExists $childFolder.Identity $userOrGroup $permissionLevel
                }
            }
        }
        [System.Windows.Forms.MessageBox]::Show("Permissions applied successfully.", "Success")
    } catch {
        [System.Windows.Forms.MessageBox]::Show("An error occurred: " + $_.Exception.Message, "Error")
    }
})

# Button for Removing User Permissions
$removeButton = New-Object System.Windows.Forms.Button
$removeButton.Text = "Remove"
$removeButton.Dock = 'Bottom'
$removeButton.Add_Click({
    try {
        $selectedFolder = $listBox.SelectedItem
        if (-not $selectedFolder) {
            [System.Windows.Forms.MessageBox]::Show("Please select a public folder from the list.", "Input Error")
            return
        }
        $userToRemove = $userGroupTextBox.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($userToRemove)) {
            [System.Windows.Forms.MessageBox]::Show("Please enter the user email to remove.", "Input Error")
            return
        }
        Remove-PublicFolderClientPermission -Identity $selectedFolder -User $userToRemove -Confirm:$false
        if ($checkBox.Checked) {
            $childFolders = Get-PublicFolder -Recurse -Identity $selectedFolder
            foreach ($childFolder in $childFolders) {
                if ($childFolder.Identity -ne $selectedFolder) {
                    Remove-PublicFolderClientPermission -Identity $childFolder.Identity -User $userToRemove -Confirm:$false
                }
            }
        }
        [System.Windows.Forms.MessageBox]::Show("User permissions removed successfully.", "Success")
    } catch {
        [System.Windows.Forms.MessageBox]::Show("An error occurred: " + $_.Exception.Message, "Error")
    }
})

# Add Apply and Remove Buttons to the panel instead of directly to the form
$panel.Controls.Add($applyButton)
$panel.Controls.Add($removeButton)


# Load the form content and show it
$form.Add_Shown({
    $form.Activate()
})
$form.ShowDialog()

# Disconnect from Exchange Online
Disconnect-ExchangeOnline -Confirm:$false