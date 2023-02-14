cd C:\data\mail
del MailKitPath.path
del MimeKitPath.path
cd C:\Program Files\PowerShell\7
pwsh -executionpolicy remotesigned -File  C:\data\mail\MailKit\ResolveMailKitPath.ps1
pwsh -executionpolicy remotesigned -File  C:\data\mail\MailKit\ResolveMimeKitPath.ps1