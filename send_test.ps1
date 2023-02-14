Add-Type -Path "C:\Program Files\PackageManagement\NuGet\Packages\MailKit.3.2.0\lib\netstandard2.0\MailKit.dll"
 Add-Type -Path "C:\Program Files\PackageManagement\NuGet\Packages\MimeKit.3.2.0\lib\netstandard2.0\MimeKit.dll"
 $SMTP     = New-OBJECT MailKit.Net.Smtp.SmtpClient
 $Message  = New-OBJECT MimeKit.MimeMessage
 $TextPart = [MimeKit.TextPart]::new("plain")
 $TextPart.Text = "This mail is only a test, do not reply."
 $Message.From.Add("trasferta1.gigapiu@gmail.com")
 $Message.To.Add("martin.circelli@gigapiu.it")
 $Message.To.Add("service@gigapiu.it")
 $Message.To.Add("")
 $Message.To.Add("")
 $Message.Subject = 'Test Email'
 $Message.Body    = $TextPart
 $SMTP.Connect('smtp.gmail.com', 587, $SecureSocketOptions.StartTlsWhenAvailable)
 $SMTP.Authenticate('trasferta1.gigapiu@gmail.com', 'pfataajppxmrwgax' )
 $SMTP.Send($Message)
 $SMTP.Disconnect($true)
 $SMTP.Dispose()
