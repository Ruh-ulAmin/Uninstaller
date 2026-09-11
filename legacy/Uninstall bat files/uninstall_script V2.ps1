# Define the list of software to remove
$softwareList = @(
    "Software1",
    "Software2"
)

# Output file path
$outputFile = "$env:USERPROFILE\Desktop\Uninstall_Log.txt"

# Function to uninstall software using wmic
function Uninstall-Software {
    param (
        [string]$ProductName
    )

    Write-Output "Uninstalling $ProductName..."
    Write-Output "--------------------------------------"

    try {
        $uninstallCommand = "wmic product where `"`"name='$ProductName'`"`" call uninstall /nointeractive"
        $output = Invoke-Expression -Command $uninstallCommand -ErrorAction Stop
        Write-Output "Uninstallation of $ProductName completed successfully."
        Write-Output "--------------------------------------"
    }
    catch {
        Write-Output "Failed to uninstall $ProductName. Error: $_"
        Write-Output "--------------------------------------"
    }
}

# Function to remove Office registry keys
function Remove-OfficeKeys {
    param (
        [string]$RegistryPath
    )

    Write-Output "Removing Office keys: $RegistryPath"
    Write-Output "--------------------------------------"

    try {
        Remove-Item -Path $RegistryPath -Recurse -Force -ErrorAction Stop
        Write-Output "Office keys removal completed successfully."
        Write-Output "--------------------------------------"
    }
    catch {
        Write-Output "Failed to remove Office keys: $RegistryPath. Error: $_"
        Write-Output "--------------------------------------"
    }
}

# Discover Office installation paths
$officePaths = Get-ChildItem -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall" | Get-ItemProperty | Where-Object { $_.DisplayName -like "Microsoft Office*" } | Select-Object -ExpandProperty UninstallString

# Remove software
foreach ($software in $softwareList) {
    Uninstall-Software -ProductName $software
}

# Remove Office keys for each installation path
foreach ($officePath in $officePaths) {
    $registryPath = $officePath -replace "msiexec.exe /x", "" -replace "/.*$"
    Remove-OfficeKeys -RegistryPath $registryPath
}

# Save output to a text file
$output = "Script execution completed." + [Environment]::NewLine
$output += "--------------------------------------" + [Environment]::NewLine
$output += "Please check the log file: $outputFile" + [Environment]::NewLine
$output += "--------------------------------------" + [Environment]::NewLine
$output += "Log:" + [Environment]::NewLine
$output += "--------------------------------------" + [Environment]::NewLine
$output += Get-Content -Path $MyInvocation.MyCommand.Path | Out-String

$output | Out-File -FilePath $outputFile -Encoding UTF8

Write-Output $output
