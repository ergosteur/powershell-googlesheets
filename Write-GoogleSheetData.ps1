# Define the Google Sheets ID and the access token obtained earlier
$spreadsheetId = "your-google-sheet-id"
$accessToken = "your-access-token"
$sheetName = "Sheet1"  # The name of the sheet you want to work with

# Step 1: Retrieve the sheetId based on the sheet name
$sheetInfoEndpoint = "https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}?fields=sheets.properties"

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
$data = "$currentDateTime, $output1, $output2"

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
$insertRowEndpoint = "https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}:batchUpdate"

$response = Invoke-RestMethod -Uri $insertRowEndpoint `
    -Method Post `
    -Headers @{ "Authorization" = "Bearer ${accessToken}"; "Content-Type" = "application/json" } `
    -Body $insertRowRequestJson

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

# Use curly braces to correctly interpret the variable in the string
$pasteDataEndpoint = "https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}:batchUpdate"

$response = Invoke-RestMethod -Uri $pasteDataEndpoint `
    -Method Post `
    -Headers @{ "Authorization" = "Bearer ${accessToken}"; "Content-Type" = "application/json" } `
    -Body $pasteDataRequestJson

Write-Output "Data added to Google Sheet."