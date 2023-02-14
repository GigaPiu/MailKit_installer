# Requires -RunAsAdministrator
# REGISTRAZIONE NUGET.ORG PER DOWNLOAD PACCHETTO MAILKIT
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 
register-packagesource -Name 'NuGet.org' -Location https://www.nuget.org/api/v2 -ProviderName Nuget -Trusted -Force
# INSTALLAZIONE PACCHETTI MAILKIT & MIMEKIT
install-package -Name 'MailKit' -Source 'NuGet.org' -SkipDependencies -Force
install-package -Name 'MimeKit' -Source 'NuGet.org' -SkipDependencies -Force
# INSTALLAZIONE COMANDO CHOCOLATEY PER POWERSHELL 6
Set-ExecutionPolicy Bypass -Scope Process -Force; [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; iex ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))
choco feature enable -n=allowGlobalConfirmation
# INSTALLAZIONE NUOVO TERMINALE WINDOWS
choco install microsoft-windows-terminal
# INSTALLAZIONE POWERSHELL 7
choco install powershell-core