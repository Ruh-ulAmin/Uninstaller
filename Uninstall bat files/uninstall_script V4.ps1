$scriptName = "uninstall_script V4.ps1"
$scriptPath = Get-ChildItem -Path $HOME -Filter $scriptName -Recurse -File | Select-Object -First 1 -ExpandProperty FullName
$powershellPath = (Get-Command -Name "powershell.exe").Source

if ($scriptPath) {
    Start-Process -FilePath $powershellPath -ArgumentList "-File '$scriptPath'" -Verb RunAs
}
else {
    Write-Host "Script '$scriptName' not found."
}



# Function to write output to file
function Write-OutputToFile {
    param (
        [string]$Output,
        [string]$FilePath
    )

    Add-Content -Path $FilePath -Value $Output -Force
}

# Get the current directory
$currentDirectory = Get-Location

# Create the "Log" folder if it doesn't exist
$logFolder = Join-Path -Path $currentDirectory -ChildPath "Log"
if (-not (Test-Path -Path $logFolder -PathType Container)) {
    New-Item -Path $logFolder -ItemType Directory | Out-Null
}

# Generate the log file name with the current date and time
$logFileName = "$(Get-Date -Format 'yyyy-MM-dd_HH-mm-ss')_Uninstalled_log.txt"

# Set the log file path inside the "Log" folder
$outputFile = Join-Path -Path $logFolder -ChildPath $logFileName

# Function to prompt for the first letter of software
function Prompt-FirstLetter {
    $firstLetter = Read-Host "Enter the first letter of the software you want to uninstall (or 'Quit' to quit this check)"
    if ($firstLetter -eq 'Quit') {
        return $null
    }
    elseif ($firstLetter -match '^[a-zA-Z]$') {
        return $firstLetter.ToUpper()
    }
    else {
        $output = "Invalid input. Please enter a single alphabetic character or 'Quit' to quit."
        Write-Host $output
        Write-OutputToFile -Output $output -FilePath $outputFile
        return Prompt-FirstLetter
    }
}

function Prompt-SoftwareChoice {
    param (
        [string]$PromptMessage,
        [array]$SoftwareList
    )

    Write-Host $PromptMessage
    for ($i = 0; $i -lt $SoftwareList.Count; $i++) {
        Write-Host "$($i + 1). $($SoftwareList[$i].Name)"
    }
    Write-Host "Q. Quit"

    $choice = Read-Host "Enter the number corresponding to the software you want to uninstall (or 'Q' to quit)"

    if ($choice -ge 1 -and $choice -le $SoftwareList.Count) {
        return $SoftwareList[$choice - 1].Name
    } elseif ($choice -eq 'Q' -or $choice -eq 'q') {
        return "Quit"
    } else {
        Write-Host "Invalid choice. Please try again."
        return (Prompt-SoftwareChoice -PromptMessage $PromptMessage -SoftwareList $SoftwareList)
    }
}

function Prompt-Password {
    param (
        [string]$PromptMessage
    )

    $password = Read-Host -Prompt $PromptMessage -AsSecureString
    return [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($password))
}

function Uninstall-Software {
    param (
        [string]$ProductName
    )

    $uninstallSuccess = $false
    $outputFile = "uninstall_log.txt"

    # Approach 1: UninstallString from Get-Package
    $package = Get-Package | Where-Object { $_.Name -eq $ProductName }
    if ($package) {
        try {
            $uninstallArgs = if ($package.PackageFamilyName) { "-PackageFamilyName $($package.PackageFamilyName)" } else { "/x `"$($package.PackageId)`" /qn" }
            $processInfo = Start-Process -FilePath "msiexec.exe" -ArgumentList $uninstallArgs -PassThru -Wait
            if ($processInfo.ExitCode -eq 0) {
                $output = "Successfully uninstalled software: $ProductName"
                Write-Host $output
                Write-OutputToFile -Output $output -FilePath $outputFile
                $uninstallSuccess = $true
            }
        } catch {
            Write-Host "An error occurred while uninstalling $ProductName :`n$($_.Exception.Message)"
        }
    }

    # Approach 2: WMIC Command
    if (-not $uninstallSuccess) {
        $uninstallCommand = "wmic product where name='$ProductName' call uninstall"
        try {
            $output = Invoke-Expression -Command $uninstallCommand
            if ($output -notlike "*Invalid query*") {
                $output = "Successfully uninstalled software using 'wmic' command: $ProductName"
                Write-Host $output
                Write-OutputToFile -Output $output -FilePath $outputFile
                $uninstallSuccess = $true
            }
        } catch {
            Write-Host "An error occurred while uninstalling $ProductName using 'wmic' command:`n$($_.Exception.Message)"
        }
    }

    # Approach 3: msiexec.exe Uninstall
    if (-not $uninstallSuccess) {
        $uninstallArgs = "/x `"$($ProductName.Replace(" ", "{SPACE}"))`" /qn"
        try {
            $processInfo = Start-Process -FilePath "msiexec.exe" -ArgumentList $uninstallArgs -PassThru -Wait
            if ($processInfo.ExitCode -eq 0) {
                $output = "Successfully uninstalled software using 'msiexec.exe': $ProductName"
                Write-Host $output
                Write-OutputToFile -Output $output -FilePath $outputFile
                $uninstallSuccess = $true
            }
        } catch {
            Write-Host "An error occurred while uninstalling $ProductName using 'msiexec.exe':`n$($_.Exception.Message)"
        }
    }

    # Approach 4: Uninstall registry keys
    if (-not $uninstallSuccess) {
        $uninstallKey = "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall", "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
        try {
            $registryKey = Get-ChildItem -Path $uninstallKey -ErrorAction Stop | Get-ItemProperty | Where-Object { $_.DisplayName -eq $ProductName }
            if ($registryKey) {
                $uninstallString = $registryKey.UninstallString
                $processInfo = Start-Process -FilePath $uninstallString -ArgumentList "/S" -PassThru -Wait
                if ($processInfo.ExitCode -eq 0) {
                    $output = "Successfully uninstalled software using registry keys: $ProductName"
                    Write-Host $output
                    Write-OutputToFile -Output $output -FilePath $outputFile
                    $uninstallSuccess = $true
                }
            }
        } catch {
            Write-Host "An error occurred while uninstalling $ProductName using registry keys:`n$($_.Exception.Message)"
        }
    }

    # Approach 5: Win32_Product WMI Class
    if (-not $uninstallSuccess) {
        try {
            $wmiQuery = "SELECT * FROM Win32_Product WHERE Name='$ProductName'"
            $software = Get-WmiObject -Query $wmiQuery
            if ($software) {
                $uninstallResult = $software.Uninstall()
                if ($uninstallResult.ReturnValue -eq 0) {
                    $output = "Successfully uninstalled software using Win32_Product: $ProductName"
                    Write-Host $output
                    Write-OutputToFile -Output $output -FilePath $outputFile
                    $uninstallSuccess = $true
                } else {
                    Write-Host "Failed to uninstall $ProductName using Win32_Product. Return code: $($uninstallResult.ReturnValue)"
                }
            }
        } catch {
            Write-Host "An error occurred while uninstalling $ProductName using Win32_Product:`n$($_.Exception.Message)"
        }
    }

    if (-not $uninstallSuccess) {
        $output = "Software not found or unable to uninstall: $ProductName"
        Write-Host $output
        Write-OutputToFile -Output $output -FilePath $outputFile
    }
}

function Write-OutputToFile {
    param (
        [string]$Output,
        [string]$FilePath
    )

    try {
        $Output | Out-File -FilePath $FilePath -Append
    } catch {
        Write-Host "An error occurred while writing to the output file:`n$($_.Exception.Message)"
    }
}

# Define and populate the $softwareList variable with installed software
$softwareList = @(Get-Package -ProviderName Programs -IncludeWindowsInstaller) |
               Where-Object { $_.Name -ne $null } | Select-Object -Property Name

# Remove Software
do {
    $firstLetter = Prompt-FirstLetter
    if ($firstLetter -ne $null) {
        $softwareSubset = $softwareList | Where-Object { $_.Name -like "$firstLetter*" }
        if ($softwareSubset.Count -eq 0) {
            $output = "No software found starting with '$firstLetter'."
            Write-Host $output
            Write-OutputToFile -Output $output -FilePath $outputFile
        } else {
            $uninstallChoice = Prompt-SoftwareChoice -PromptMessage "Select the software you want to uninstall:" -SoftwareList $softwareSubset

            if ($uninstallChoice -eq "Quit") {
                $output = "Stopping / Exiting the uninstallation process."
                Write-Host $output
                Write-OutputToFile -Output $output -FilePath $outputFile
            } else {
                $softwarePasswords = @{
                    # Add more software and passwords as needed
                }

                if ($softwarePasswords.ContainsKey($uninstallChoice)) {
                    $passwordRequired = $true
                    while ($passwordRequired) {
                        $password = Prompt-Password -PromptMessage "Enter the password to uninstall $uninstallChoice (or click Cancel to skip)"
                        if ($password -eq $softwarePasswords[$uninstallChoice]) {
                            $passwordRequired = $false
                            Uninstall-Software -ProductName $uninstallChoice
                            $output = "Uninstalled software: $uninstallChoice"
                            Write-Host $output
                            Write-OutputToFile -Output $output -FilePath $outputFile
                        } elseif ($password -eq $null) {
                            $passwordRequired = $false
                            $output = "Uninstallation of $uninstallChoice skipped."
                            Write-Host $output
                            Write-OutputToFile -Output $output -FilePath $outputFile
                        } else {
                            $output = "Incorrect password. Please try again or click Cancel to skip uninstallation."
                            Write-Host $output
                            Write-OutputToFile -Output $output -FilePath $outputFile
                        }
                    }
                } else {
                    Uninstall-Software -ProductName $uninstallChoice
                }
            }
        }
    }
} until ($firstLetter -eq $null)


# Function to prompt for a yes or no choice
function Prompt-YesNoChoice {
    param (
        [string]$PromptMessage
    )

    $validInput = $false
    $choice = ""

    while (-not $validInput) {
        $choice = Read-Host "$PromptMessage (Y/N)"

        if ($choice -eq "Y" -or $choice -eq "y") {
            $validInput = $true
        }
        elseif ($choice -eq "N" -or $choice -eq "n") {
            $validInput = $true
        }
        else {
            $output = "Invalid choice. Please enter 'Y' for Yes or 'N' for No."
            Write-Host $output
            Write-OutputToFile -Output $output -FilePath $outputFile
        }
    }

    return $choice -eq "Y" -or $choice -eq "y"
}

# Function to remove Office keys
function Remove-OfficeKeys {
    param (
        [string]$RegistryPath
    )

    # Your code to remove Office keys goes here
    $output = "Removing Office keys for registry path: $RegistryPath"
    Write-Host $output
    Write-OutputToFile -Output $output -FilePath $outputFile
}

# Discover Office installation paths
$officePaths = Get-ChildItem -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall" |
    Get-ItemProperty |
    Where-Object { $_.DisplayName -like "Microsoft Office*" } |
    Select-Object -ExpandProperty UninstallString

# Prompt to check office license
$checkOfficeLicense = Prompt-YesNoChoice -PromptMessage "Do you want to check the Office license?"


if ($checkOfficeLicense) {
    if ($officePaths) {
        $output = "Active Office keys found on the system:"
        Write-Host $output
        Write-OutputToFile -Output $output -FilePath $outputFile

        $officePaths | ForEach-Object {
            $output = $_
            Write-Host $output
            Write-OutputToFile -Output $output -FilePath $outputFile
        }

        # Prompt to remove Office keys
        $removeOfficeKeys = Prompt-YesNoChoice -PromptMessage "Do you want to remove all Office license keys?"

        if ($removeOfficeKeys) {
            # Remove Office keys for each installation path
            foreach ($officePath in $officePaths) {
                $registryPath = $officePath -replace "msiexec.exe /x", "" -replace "/.*$"
                Remove-OfficeKeys -RegistryPath $registryPath
            }
        }
        else {
            $output = "Skipping removal of Office license keys."
            Write-Host $output
            Write-OutputToFile -Output $output -FilePath $outputFile
        }
    }
    else {
        $output = "No active Office keys found on the system."
        Write-Host $output
        Write-OutputToFile -Output $output -FilePath $outputFile
    }
}
else {
    $output = "Skipping Office license check."
    Write-Host $output
    Write-OutputToFile -Output $output -FilePath $outputFile
}