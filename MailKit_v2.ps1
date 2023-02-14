Invoke-WebRequest -Uri "https://github.com/microsoft/winget-cli/releases/download/v1.1.12653/Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle" -OutFile "C:\Backup_programs\WinGet.msixbundle"
Add-AppxPackage "C:\Backup_programs\WinGet.msixbundle"
cd "C:\Backup_programs"
.\WinGet.msixbundle
.\WinGet.msixbundle
winget install --id Microsoft.Powershell --source 'winget'
winget install Microsoft.WindowsTerminal
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
register-packagesource -Name 'NuGet.org' -ForceLocation https://www.nuget.org/api/v2 -ProviderName Nuget -Trusted -Force
install-package -Name 'MailKit' -Source 'NuGet.org' -SkipDependencies -Force
install-package -Name 'MimeKit' -Source 'NuGet.org' -SkipDependencies -Force
	
	
	

