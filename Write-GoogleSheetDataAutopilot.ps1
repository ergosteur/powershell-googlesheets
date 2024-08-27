param (
    [switch]$Debug,
    [switch]$RefreshToken,
    [string]$SpreadsheetId = "1E5rUayzVrFi_rjUoW0PKKEDsljlDFvvYQ-r9ztolOyA"  # Fallback to hardcoded value
)

# Path to your service account JSON key file and access token file
$serviceAccountKeyFile = "userdata\service-account-key.json"
$accessTokenFile = "userdata\accesstoken.txt"

# Function to generate a new access token
function Get-NewAccessToken {
    if (-not (Test-Path $serviceAccountKeyFile)) {
        Write-Error "Service account key file not found: $serviceAccountKeyFile. Please provide a valid file."
        exit 1
    }

    try {
        # Load the JSON key file
        $json = Get-Content $serviceAccountKeyFile | ConvertFrom-Json
    } catch {
        Write-Error "Failed to load or parse the service account key file: $serviceAccountKeyFile. Please ensure the file is valid."
        exit 1
    }

    # Create a JWT Header
    $header = @{
        alg = "RS256"
        typ = "JWT"
    } | ConvertTo-Json -Compress

    # Create a JWT Claim Set
    $now = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
    $expires = $now + 3600
    $scope = "https://www.googleapis.com/auth/spreadsheets"

    $claimSet = @{
        iss = $json.client_email
        scope = $scope
        aud = "https://oauth2.googleapis.com/token"
        exp = $expires
        iat = $now
    } | ConvertTo-Json -Compress

    # Base64 URL Encode the Header and Claim Set
    $base64Header = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($header)).TrimEnd('=').Replace('+', '-').Replace('/', '_')
    $base64ClaimSet = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($claimSet)).TrimEnd('=').Replace('+', '-').Replace('/', '_')

    # Create the unsigned JWT
    $unsignedJWT = "$base64Header.$base64ClaimSet"

    # Process and sign the private key correctly
    $privateKey = $json.private_key -replace "\\n", "`n"

    $rsa = New-Object System.Security.Cryptography.RSACryptoServiceProvider
    $rsa.ImportFromPem($privateKey)

    $signatureBytes = $rsa.SignData([System.Text.Encoding]::UTF8.GetBytes($unsignedJWT), [System.Security.Cryptography.HashAlgorithmName]::SHA256, [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)
    $signature = [Convert]::ToBase64String($signatureBytes).TrimEnd('=').Replace('+', '-').Replace('/', '_')

    # Create the signed JWT
    $signedJWT = "$unsignedJWT.$signature"

    # Request an access token
    try {
        $response = Invoke-RestMethod -Uri "https://oauth2.googleapis.com/token" -Method Post -ContentType "application/x-www-form-urlencoded" -Body @{
            grant_type = "urn:ietf:params:oauth:grant-type:jwt-bearer"
            assertion = $signedJWT
        }
    } catch {
        Write-Error "Failed to retrieve access token. Please check your service account credentials and network connection."
        exit 1
    }

    # Extract the access token and save it to the file
    $accessToken = $response.access_token
    Set-Content -Path $accessTokenFile -Value $accessToken
    Write-Output "Access token refreshed and saved to $accessTokenFile."
    Write-Output "Please share your Google Sheet with the Service Account Email $($json.client_email)"
}

# Main script execution

# Check if the RefreshToken flag is used
if ($RefreshToken) {
    Get-NewAccessToken
    exit 0
}

# Ensure the access token file exists and is not empty
if (-not (Test-Path $accessTokenFile) -or -not (Get-Content $accessTokenFile)) {
    Write-Error "Access token file is missing or empty. Please run the script with the -RefreshToken flag to generate a new access token."
    exit 1
}

# Load the access token
$accessToken = Get-Content $accessTokenFile

try {
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
            Write-Output $output
        } catch {
            Write-Error "Failed to retrieve the Hardware Hash using Get-WindowsAutopilotInfo.ps1."
            exit 1
        }
    }

    # Main script execution
    #Test-Administrator
    Test-WindowsAutopilotInfo
    $output3 = Get-HardwareHash

    # Step 1: Retrieve the sheetId based on the sheet name
    $sheetInfoEndpoint = "https://sheets.googleapis.com/v4/spreadsheets/${SpreadsheetId}?fields=sheets.properties"

    $response = Invoke-RestMethod -Uri $sheetInfoEndpoint `
        -Method Get `
        -Headers @{ "Authorization" = "Bearer ${accessToken}"; "Content-Type" = "application/json" }

    # Find the sheetId for the given sheet name
    $sheetId = $response.sheets | Where-Object { $_.properties.title -eq $sheetName } | Select-Object -ExpandProperty properties | Select-Object -ExpandProperty sheetId

    # Define the commands whose output you want to send to the Google Sheet
    $output1 = $(Get-WmiObject win32_bios).serialnumber
    $output2 = $(Get-WmiObject win32_bios).Manufacturer

    # Get the current date and time in the desired format
    $currentDateTime = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")

    # Prepare the data to be inserted
    $data = "$currentDateTime, $output1, $output2, $output3"

    # Step 2: Insert a new row at the top of the sheet
    $insertRowRequest = @{
        requests = @(
            @{
                insertRange = @{
                    range = @{
                        sheetId = $sheetId
                        startRowIndex = 0
                        endRowIndex = 1
                    }
                    shiftDimension = "ROWS"
                }
            }
        )
    }

    # Convert the insertRowRequest to JSON
    $insertRowRequestJson = $insertRowRequest | ConvertTo-Json -Depth 5

    # Use curly braces to correctly interpret the variable in the string
    $insertRowEndpoint = "https://sheets.googleapis.com/v4/spreadsheets/${SpreadsheetId}:batchUpdate"

    $response = Invoke-RestMethod -Uri $insertRowEndpoint `
        -Method Post `
        -Headers @{ "Authorization" = "Bearer ${accessToken}"; "Content-Type" = "application/json" } `
        -Body $insertRowRequestJson

    # Check if the response is successful
    if ($response) {
        # Step 3: Paste the data into the newly inserted row
        $pasteDataRequest = @{
            requests = @(
                @{
                    pasteData = @{
                        data = $data
                        type = "PASTE_NORMAL"
                        delimiter = ","
                        coordinate = @{
                            sheetId = $sheetId
                            rowIndex = 0
                        }
                    }
                }
            )
        }

        # Convert the pasteDataRequest to JSON
        $pasteDataRequestJson = $pasteDataRequest | ConvertTo-Json -Depth 5

        $pasteDataEndpoint = "https://sheets.googleapis.com/v4/spreadsheets/${SpreadsheetId}:batchUpdate"

        $response = Invoke-RestMethod -Uri $pasteDataEndpoint `
            -Method Post `
            -Headers @{ "Authorization" = "Bearer ${accessToken}"; "Content-Type" = "application/json" } `
            -Body $pasteDataRequestJson

        # Check if the response is successful
        if ($response) {
            Write-Output "Data added to Google Sheet."
        } else {
            Write-Output "Failed to add data to Google Sheet."
            if ($Debug) {
                Write-Output "Debug Info: $response"
            }
        }
    } else {
        Write-Output "Failed to insert a new row in Google Sheet."
        if ($Debug) {
            Write-Output "Debug Info: $response"
        }
    }
}
catch {
    # Get API Response JSON to Object
    $errorResponseObj = $_ | ConvertFrom-Json
    
    # Display the error message directly
    Write-Output "Error Message: $($_.Exception.Message)"

    if ($Debug) {
        Write-Output "====DEBUG OUTPUT===="
        Write-Output "Full Error: $_"
        $errorResponseObj.error | Format-List -Property *
        $errorResponseObj.error.details | Format-List -Property *
        Write-Output "====DEBUG OUTPUT===="
    }

    # Handle 40x errors
    if ($errorResponseObj.error.code -like "40*") {
        $errorResponseObj.error | Format-List -Property code,message,status
        $errorResponseObj.error.details | Format-List -Property reason,metadata
        if ($errorResponseObj.error.details.reason -eq "ACCESS_TOKEN_EXPIRED") {
            Write-Warning "Access token is expired. Please re-run the script with the -RefreshToken flag."
        } elseif ($errorResponseObj.error.status -eq "PERMISSION_DENIED") {
            Write-Warning "Write access to specified Google Sheet denied. Please make sure the Service Account Email has Edit access to the Google Sheet."
        }
    } else {
        $errorResponseObj.error | Format-List -Property *
    }
}


