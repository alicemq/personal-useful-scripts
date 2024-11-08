# This script takes regular Outlook message (preferably as draft)
# and sends this message to email addresses

$message = "email_message.msg"

$recipients = @(
    "info@foobar.com",
    "info@contoso.com"
)


foreach ($recipient in $recipients) {
$outlook = New-Object -comObject Outlook.Application 
$mail = $outlook.Session.OpenSharedItem($message)
$mail.Forward() | Out-Null
$mail.Recipients.Add($recipient) | Out-Null
$mail.send() | Out-Null
Write-Output("Mail sent to $recipient")
}
