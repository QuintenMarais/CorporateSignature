# Define paths
$scriptName = "Corporate Signature 2024Q1.ps1"
$destinationFolder = "C:\Signatures\2024Q1"
$desktopPath = [Environment]::GetFolderPath("Desktop")
$currentLocation = Get-Location
$iconName = "icon.ico"
$shortcutName = "Corporate Signature 2024Q1.lnk"

# Step 1: Create the destination folder if it doesn't exist
if (-not (Test-Path -Path $destinationFolder -PathType Container)) {
    New-Item -Path $destinationFolder -ItemType Directory -Force
}

# Step 2: Copy the script file and the icon to the destination folder
Copy-Item -Path (Join-Path -Path $currentLocation -ChildPath $scriptName) -Destination $destinationFolder -Force
Copy-Item -Path (Join-Path -Path $currentLocation -ChildPath $iconName) -Destination $destinationFolder -Force

# Step 3: Create a shortcut on the desktop
$shortcutPath = Join-Path -Path $desktopPath -ChildPath $shortcutName
$shortcutTarget = Join-Path -Path $destinationFolder -ChildPath $scriptName
$iconPath = Join-Path -Path $destinationFolder -ChildPath $iconName

$WshShell = New-Object -ComObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut($shortcutPath)
$Shortcut.TargetPath = "powershell.exe"
$Shortcut.Arguments = "-ExecutionPolicy Bypass -File `"$shortcutTarget`""
$Shortcut.WorkingDirectory = $destinationFolder
$Shortcut.IconLocation = $iconPath
$Shortcut.Save()
