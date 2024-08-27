# API Access Token (Update to use your own token)
$accessToken = "ya29.c.c0ASRK0GbgEf1LRlXJyJ3-E7mPPRa5ZvrnRekvkNigUJg6Zze5xo3HQ1pPax0PRl6sBwdCkfym0e9jmCvYycRUo7cI5ji01DCAR5UKlQh66pRZcUPbVbmX9LBYgf510or7IXie9quLF5TPt31mJOllp9ItMHteiS36Kasja_x7JcZMEldVQrFN-MsaS0y-63aW9awtt_kaGUxqQCjsJ04t6Wuri44yjbdnr2H3nvUD1YDrJr4bYiZOnxZuQzUdDA-EkAdMRrxQZEW1MTjv7Mz_PLTPaFqJnETUDF8dU5nFa1Uu9BPkHg-Gj2Pb-uKOc0T1qGIwFRWpQ-k7aIQUUSPz-BdfOELlCmOByL4h2yJzOMOee71jJ02gysQADwT387Chw8_7wcoVdQkFvgZbYk5gcgWXin7YyglXgWr-zd47SdW5k4YY9bBQQOmfYlY_UbMnz-O1OoI3M2M0624XtyUtRbXvjzQuajrVsexb-7jXRz6-0g4qMszgs_JYn16dylRIXU3eZBpW74Rz-33Jw2soQQm7tS1wd-cQ7wibpZcSdg52dfZubwFalwp2ogrVoORd9vVlytaO6bdRaF1hwlFzm2X_yR77z_pvOVW8rZc2jcBbO1YZlYrkrBe38nphsIMORpYl1XOg4p1bqVhSZgB7s3q_1oYo6wXi7lui-Oo55s1cVqWX3gIqbeqiIOwh9J0plo93Wcw9mU9y48ShZByrMVbizqORbmozBfwVRz89My1qvO2xesIjW15mx-eYX5dbflm0rjfdhRZvY65F0iSFeoYJJe2MFQIXu6t5dy07zv14rsS2Z1eq2J-2IMa9ZhUz2z6QXey0uQbRpn_98SyMigW3hMIoiJt23pw_t0V-JtJcl8tY3YRkpJVS6dZY4q6s-ztU6UiyeiYl0gWV28guMyxOcQ2rlcF04MpgXB6Fp1wue_YQvJQYozk9jvrQ5IcndMOnl0Ybb7x51xlQZlJxRf4rgznqfk0i8SWiqo9_S88jbqpOMOOO3Ze"
$serviceAccountEmail = "powershell-scripting-test@powershell-scripting-433721.iam.gserviceaccount.com"

# Define the Google Sheets ID and the access token obtained earlier
$spreadsheetId = "1E5rUayzVrFi_rjUoW0PKKEDsljlDFvvYQ-r9ztolOyA"
$sheetName = "Sheet1"

try {
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

        $pasteDataEndpoint = "https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}:batchUpdate"

        $response = Invoke-RestMethod -Uri $pasteDataEndpoint `
            -Method Post `
            -Headers @{ "Authorization" = "Bearer ${accessToken}"; "Content-Type" = "application/json" } `
            -Body $pasteDataRequestJson

        # Check if the response is successful
        if ($response) {
            Write-Output "Data added to Google Sheet."
        } else {
            Write-Output "Failed to add data to Google Sheet."
        }
    } else {
        Write-Output "Failed to insert a new row in Google Sheet."
    }
}
catch {
    $errorMessage = $_.Exception.Message

    if ($errorMessage -like "*403*") {
        $customMessage = "The service account does not have permission to edit the Google Sheet."
        $recommendedAction = "Please ensure the sheet is shared with the service account email $serviceAccountEmail."

        Write-Error -Message "$errorMessage $customMessage" -Category PermissionDenied -RecommendedAction $recommendedAction
        Write-Output "$recommendedAction"
    } else {
        Write-Error -Message "An unexpected error occurred: $errorMessage"
    }
}
