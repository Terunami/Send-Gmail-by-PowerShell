Param(
    [parameter(mandatory=$true)][string]$from,
    [parameter(mandatory=$true)][string]$to,
    [string]$cc,
    [string]$bcc,
    [string]$title,
    [string]$body,
    [array]$atts
)

function ConvertTo-Base64Url($str){
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($str)
    $b64str = [System.Convert]::ToBase64String($bytes)
    $without_plus = $b64str -replace '\+', '-'
    $without_slash = $without_plus -replace '/', '_'
    $without_equal = $without_slash -replace '=', ''
    return $without_equal
}

function Send-Gmail($from, $to, $cc, $bcc, $title, $body, $atts){
    $path = Join-Path . "google_credential.psm1"
    Import-Module -Name $path
    $credential = Get-GoogleCredential
    Write-Host $credential
    if (-not $credential){
        Write-Host "Not Authenticated."
        return
    }
    $dll = Join-Path . "AE.NET.Mail.dll"
    Add-Type -Path $dll
    $msg = New-Object AE.Net.Mail.MailMessage

    if (!([string]::IsNullOrEmpty($cc))) {
        $mail_cc = New-Object System.Net.Mail.MailAddress $cc
        $msg.Cc.Add($mail_cc)
    }
    if (!([string]::IsNullOrEmpty($bcc))) {
        $mail_bcc = New-Object System.Net.Mail.MailAddress $bcc
        $msg.Bcc.Add($mail_bcc)
    }
    Write-Host $from
    $mail_from = New-Object System.Net.Mail.MailAddress $from
    $mail_to = New-Object System.Net.Mail.MailAddress $to
    $msg.From = $mail_from
    $msg.To.Add($mail_to)

    Write-Host "Atts length: $($atts.length)"
    Write-Host "Atts: $($atts)"
    if (!($atts[0] -eq "")) {
        foreach($att in $atts){
            $file_path = $att
            Write-Host $file_path
            $file_name = [System.IO.Path]::GetFileName($file_path)
            $file_bytes = [System.IO.File]::ReadAllBytes($file_path)
            $file_mime = "application/octet-stream"
            $attachment = New-Object AE.Net.Mail.Attachment($file_bytes, $file_mime, $file_name)
            $attachment.Headers.Add("Content-Disposition", "inline; filename=$($file_name)")
            $msg.Attachments.Add($attachment) 
        }
    }

    $subject_encoded = ConvertTo-Base64Url $title
    $msg.Subject = "=?utf-8?B?$($subject_encoded)?="
    $msg.Body = $body

    $sw = New-Object System.IO.StringWriter
    $msg.Save($sw)
    $raw = ConvertTo-Base64Url $sw.ToString()
    $body = @{ "raw" = $raw; } | ConvertTo-Json
    Write-Host $sw.ToString()
    $user_id = "me"
    $uri = "https://www.googleapis.com/gmail/v1/users/$($user_id)/messages/send?access_token=$($credential.access_token)"

    try {
        $result = Invoke-RestMethod $uri -Method POST -ErrorAction Stop -Body $body -ContentType "application/json"
    }
    catch [System.Exception] {
        Write-Host $Error
        return
    }
    Write-Host $result
}

echo "$from, $to, $cc, $bcc, $title, $body, $atts"
$atts_arr = $atts -split "," 
Send-Gmail $from $to $cc $bcc $title $body $atts_arr