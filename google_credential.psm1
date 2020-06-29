set CREDENTIAL_FILE (Join-Path . "credential.json")
set SECRET_FILE (Join-Path . "client_id.json")
set DATE_FORMAT "yyyy/MM/dd HH:mm:ss"
set GMAIL_SCOPE "https://www.googleapis.com/auth/gmail.send"

function Save-GoogleCredential($credential){
    $credential | Add-Member created_at (Get-Date).ToString($DATE_FORMAT) -Force
    $credential_file = Join-Path . $CREDENTIAL_FILE
    $credential | ConvertTo-Json | Out-File $CREDENTIAL_FILE -Encoding utf8
}

function Get-GoogleCredential(){
    if (-not(Test-Path $SECRET_FILE)) {
        Write-Host "Not found client_id.json file"
        return $null
    }
    $json = Get-Content $SECRET_FILE -Encoding UTF8 -Raw | ConvertFrom-Json
    $auth = $json.installed
    if (Test-Path $CREDENTIAL_FILE) {
        $current_credential = Get-Content $CREDENTIAL_FILE -Encoding UTF8 -Raw | ConvertFrom-Json
        Write-Host $current_credential.access_token
        Write-Host $current_credential.token_type
        Write-Host $current_credential.expires_in
        Write-Host $current_credential.refresh_token
        Write-Host $current_credential.created_at
        if (-not ($current_credential.access_token -and $current_credential.token_type -and $current_credential.expires_in `
                  -and $current_credential.refresh_token -and $current_credential.created_at))
        {
            Write-Host "No credential file: $($CREDENTIAL_FILE)"
            return $null
        }
        $elapsed_seconds = ((Get-Date) - [DateTime]::ParseExact($current_credential.created_at, $DATE_FORMAT, $null)).TotalSeconds
        if ($elapsed_seconds -lt $current_credential.expires_in ) {
            Write-Host "Reuse access token..."
            return $current_credential
        }
        else{
            Write-Host "Refresh access token..."
            $refresh_body = @{
                "refresh_token" = $current_credential.refresh_token;
                "client_id" = $auth.client_id;
                "client_secret" = $auth.client_secret;
                "grant_type" = "refresh_token";
            }
            try {
                $refreshed_credential = Invoke-RestMethod -Method Post -Uri $auth.token_uri -Body $refresh_body
                $refreshed_credential | Add-Member refresh_token $refresh_body.refresh_token -Force
            }
            catch [System.Exception] {
                Write-Host $Error
                return $null
            }
            Save-GoogleCredential $refreshed_credential
            return $refreshed_credential
        }
    }

    Write-Host "New access token..."
    $gmail_scope = "https://www.googleapis.com/auth/gmail.send"
    $auth_url = "$($auth.auth_uri)?scope=$($GMAIL_SCOPE)"
    $auth_url += "&redirect_uri=$($auth.redirect_uris[0])"
    $auth_url += "&client_id=$($auth.client_id)"
    $auth_url += "&response_type=code&approval_prompt=force&access_type=offline"
    Start-Process $auth_url
    $code = Read-Host "ブラウザに表示されている認証コードを入力してください。"
    try {
        $new_body = @{
            "client_id" = $auth.client_id;
            "client_secret" = $auth.client_secret;
            "redirect_uri" = $auth.redirect_uris[0];
            "grant_type" = "authorization_code";
            "code" = $code;
        }
        $new_credential = Invoke-RestMethod -Method Post -Uri $auth.token_uri -Body $new_body
    }
    catch [System.Exception] {
        Write-Host $Error
    }
    Save-GoogleCredential $new_credential
    return $new_credential
}

Export-ModuleMember -function Get-GoogleCredential
