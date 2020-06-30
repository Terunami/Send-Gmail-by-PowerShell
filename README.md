# Send-Gmail-by-PowerShell
PowerShellを使ったGmailの自動送信を実装してみる。

-------------------------------------------

Param(
    [parameter(mandatory=$true)][string]$from,
    [parameter(mandatory=$true)][string]$to,
    [string]$cc,
    [string]$bcc,
    [string]$title,
    [string]$body,
    [array]$atts
)

gmail_sender.ps1 [送り元] [送り先] -title [題名] -body [本文]

ファイルの添付がうまくいかない。
他の方法を探す。