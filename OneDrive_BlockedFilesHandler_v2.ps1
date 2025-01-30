# Load necessary .NET assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName Microsoft.VisualBasic
Add-Type -AssemblyName System.IO

# Define blocked extensions and the target directory
[array]$blockedExtensions = ".iqy", ".ace", ".ani", ".app", ".chm", ".com", ".csh", ".exe", ".fxp", ".hpj", ".gadget", ".ins", ".isp", ".its", ".jar", ".ksh", ".ocx", ".pl", ".reg", ".scf", ".scr", ".sct", ".shs", ".vbe", ".vbs", ".z"
$destinationPath = "C:\Users\$env:USERNAME\OneDrive - OECD"

# Initialize the Windows Form
$form = New-Object System.Windows.Forms.Form
$form.Text = 'OneDrive Blocked Files Handler'
$form.Size = New-Object System.Drawing.Size(550, 550)
$form.StartPosition = 'CenterScreen'

# Load the custom icon
#$iconPath = ".\Content\icon.ico"  # Specify the path to your icon file
#$form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($iconPath)

# Initialize and configure the TreeView
$treeView = New-Object System.Windows.Forms.TreeView
$treeView.Dock = [System.Windows.Forms.DockStyle]::Fill
$treeView.CheckBoxes = $true
$treeView.Font = New-Object System.Drawing.Font("Segoe UI", 11) # Reduced font size for better visibility

# Function to check if files are selected
function Check-FilesSelected {
    $selectedNodes = $treeView.Nodes[0].Nodes | Where-Object { $_.Checked }
    if ($selectedNodes.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('No files selected.', 'Warning', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return $false
    }
    return $true
}

# Function to populate the TreeView with blocked files
function Populate-TreeView {
    param ([string]$path)
    $treeView.Nodes.Clear()
    $rootNode = New-Object System.Windows.Forms.TreeNode("Select All")
    $rootNode.NodeFont = New-Object System.Drawing.Font("Segoe UI", 11) # Same font size as other nodes
    $treeView.Nodes.Add($rootNode)

    Get-ChildItem -Path $path -Recurse -Force -File | Where-Object {
        $blockedExtensions -contains $_.Extension -and $_.DirectoryName.StartsWith($destinationPath)
    } | ForEach-Object {
        $node = New-Object System.Windows.Forms.TreeNode($_.FullName)
        $node.Tag = $_.FullName
        $rootNode.Nodes.Add($node)
    }

    $rootNode.Expand()
}

# Add the TreeView to the form
$form.Controls.Add($treeView)

$treeView.Add_AfterCheck({
    param($sender, $e)
    # Check if the action is performed on the "Blocked Files" node
    if ($e.Node.Text -eq "Select All") {
        # Set the Checked property of all child nodes to match the "Blocked Files" node
        $e.Node.Nodes | ForEach-Object { $_.Checked = $e.Node.Checked }
    }
})

# Function to handle renaming of blocked files
function Rename-File {
    param ([string]$filePath)
    # Adjust to append ".renamed" to the full filename
    $newFilePath = "$filePath.renamed"
    Rename-Item -Path $filePath -NewName $newFilePath
}

# Function to move selected blocked files to a different location
function Move-FilesToDifferentLocation {
    param ([string[]]$filePaths, [string]$destination)
    $movedCount = 0
    foreach ($filePath in $filePaths) {
        Move-Item -Path $filePath -Destination $destination
        $movedCount++
    }
    [System.Windows.Forms.MessageBox]::Show("$movedCount file(s) moved to the chosen location", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
}

# Function to open the location of a single selected file
function Open-FileLocation {
    param ([string[]]$filePaths)
    if ($filePaths.Count -eq 1) {
        explorer.exe "/select,`"$($filePaths[0])`""
    } elseif ($filePaths.Count -gt 1) {
        [System.Windows.Forms.MessageBox]::Show("Please select a single file to open its location.", "Selection Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    }
}

# Function to show information about blocked file extensions
function Show-ExtensionInfo {
    $infoMessage = @"
Files with certain extensions are not authorized in OneDrive synchronization. These file types include:

- .iqy
- .ace
- .ani
- .app
- .chm
- .com
- .csh
- .exe
- .fxp
- .hpj
- .gadget
- .ins
- .isp
- .its
- .jar
- .ksh
- .ocx
- .pl
- .reg
- .scf
- .scr
- .sct
- .shs
- .vbe
- .vbs
- .z

It is important to avoid syncing these file types to ensure the security and stability of your OneDrive storage.
"@

    [System.Windows.Forms.MessageBox]::Show($infoMessage, "Blocked File Extensions Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
}

# Adding buttons for actions
$buttonHeight = 40
$font = New-Object System.Drawing.Font("Segoe UI", 11)

# Rename button
$btnRename = New-Object System.Windows.Forms.Button
$btnRename.Text = 'Rename Selected Files and Keep Them on OneDrive'
$btnRename.Font = $font
$btnRename.Height = $buttonHeight

$btnRename.Add_Click({
    if (-not (Check-FilesSelected)) { return }
    $renamedCount = 0
    $treeView.Nodes[0].Nodes | Where-Object { $_.Checked } | ForEach-Object {
        Rename-File -filePath $_.Tag
        $renamedCount++
    }
    if ($renamedCount -gt 0) {
        [System.Windows.Forms.MessageBox]::Show("$renamedCount file(s) renamed successfully", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        Populate-TreeView -path $destinationPath
    }
})

# Move button
$btnMoveDifferentLocation = New-Object System.Windows.Forms.Button
$btnMoveDifferentLocation.Text = 'Move the Files to a Different Location'
$btnMoveDifferentLocation.Font = $font
$btnMoveDifferentLocation.Height = $buttonHeight
$btnMoveDifferentLocation.Add_Click({
    $selectedNodes = $treeView.Nodes[0].Nodes | Where-Object { $_.Checked }
    if ($selectedNodes.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('No files selected.', 'Warning', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    $folderBrowserDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($folderBrowserDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $destination = $folderBrowserDialog.SelectedPath
        $selectedFilePaths = $selectedNodes | ForEach-Object { $_.Tag }
        Move-FilesToDifferentLocation -filePaths $selectedFilePaths -destination $destination
        Populate-TreeView -path $destinationPath
    }
})

# Open location button
$btnOpenLocation = New-Object System.Windows.Forms.Button
$btnOpenLocation.Text = 'Open File Location'
$btnOpenLocation.Font = $font
$btnOpenLocation.Height = $buttonHeight
$btnOpenLocation.Add_Click({
    if (-not (Check-FilesSelected)) { return }
    $selectedNodes = $treeView.Nodes[0].Nodes | Where-Object { $_.Checked }
    $selectedFilePaths = $selectedNodes | ForEach-Object { $_.Tag }
    Open-FileLocation -filePaths $selectedFilePaths
})

# More info button
$btnMoreInfo = New-Object System.Windows.Forms.Button
$btnMoreInfo.Text = 'More Info'
$btnMoreInfo.Font = $font
$btnMoreInfo.Height = $buttonHeight
$btnMoreInfo.Add_Click({
    Show-ExtensionInfo
})

# Arrange buttons
$buttons = @($btnRename, $btnMoveDifferentLocation, $btnOpenLocation, $btnMoreInfo)
$buttons | ForEach-Object {
    $_.Dock = [System.Windows.Forms.DockStyle]::Bottom
    $form.Controls.Add($_)
}

# Initial population of the TreeView
Populate-TreeView -path $destinationPath

# Show the form
$form.ShowDialog() | Out-Null
