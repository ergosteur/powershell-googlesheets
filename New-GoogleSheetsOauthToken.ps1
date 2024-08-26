# Path to your service account JSON key file
$serviceAccountKeyFile = "userdata\service-account-key.json"

# Load the JSON key file
$json = Get-Content $serviceAccountKeyFile | ConvertFrom-Json

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
$response = Invoke-RestMethod -Uri "https://oauth2.googleapis.com/token" -Method Post -ContentType "application/x-www-form-urlencoded" -Body @{
    grant_type = "urn:ietf:params:oauth:grant-type:jwt-bearer"
    assertion = $signedJWT
}

# Extract the access token
$accessToken = $response.access_token
Write-Output "Access Token: "
Write-Output $accessToken
#Write-Output $response | Format-List -Property access_token
