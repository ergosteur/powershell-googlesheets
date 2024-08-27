# Function to check if the script is running as administrator
function Test-Administrator {
    $isAdmin = [bool](([System.Security.Principal.WindowsPrincipal] [System.Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator))
    if (-not $isAdmin) {
        Write-Error "This script must be run as an administrator. Please re-run it with elevated privileges."
        exit 1
    }
}

# Function to check if Get-WindowsAutopilotInfo.ps1 is available
function Test-WindowsAutopilotInfo {
    $scriptPath = Get-Command Get-WindowsAutopilotInfo.ps1 -ErrorAction SilentlyContinue

    if (-not $scriptPath) {
        $install = Read-Host "Get-WindowsAutopilotInfo.ps1 is not available. Do you want to install it? (y/n)"
        if ($install -eq 'y') {
            try {
                Install-Script -Name Get-WindowsAutopilotInfo -Force -ErrorAction Stop
                Write-Output "Get-WindowsAutopilotInfo.ps1 has been installed successfully."
            } catch {
                Write-Error "Failed to install Get-WindowsAutopilotInfo.ps1. Please ensure you have internet access and try again."
                exit 1
            }
        } else {
            Write-Error "Get-WindowsAutopilotInfo.ps1 is required to run this script. Exiting."
            exit 1
        }
    }
}

# Function to get the Hardware Hash using Get-WindowsAutopilotInfo.ps1
function Get-HardwareHash {
    try {
        $output = $(Get-WindowsAutopilotInfo.ps1).'Hardware Hash'
        Write-Output "Hardware Hash: $output"
    } catch {
        Write-Error "Failed to retrieve the Hardware Hash using Get-WindowsAutopilotInfo.ps1."
        exit 1
    }
}

# Main script execution
Test-Administrator
Test-WindowsAutopilotInfo
Get-HardwareHash
