# GIGAPIÙ 2022 - Martin Circelli
Write-Output "--------------------------------------------------------------------------------------------------------"
Write-Output "                                    ITALMATIC - GIGAPIU' 2022                                            "
Write-Output "                    Installing MailKit software with TLS 1.2 support. Please wait...                    "
Write-Output "--------------------------------------------------------------------------------------------------------"

Start-Sleep -Seconds 2

# CHECK WINDOWS VERSION
Write-Output "Checking the version of Windows..."
Start-Sleep -Seconds 1
if ([System.Environment]::OSVersion.Version.Build -le 17763) {
	Write-Warning "Windows version doesn't meet the requirements to use MailKit. Please upgrade Windows version before installing."
	pause
	return
} else {
  Write-Output "Windows version OK."
}

Start-Sleep -Seconds 1

# ENABLE WINDOWS UPDATE
Write-Output "Checking Windows Update service status..."
Start-Sleep -Seconds 1
$DisableUpdateAtTheEnd = 0
$GetWsus = Get-Service -Name wuauserv
if($GetWsus.Status -ne 'Running') {
  $DisableUpdateAtTheEnd = 1
  Write-Output "Enabling Windows Update service..."
  NET START WUAUSERV
  if  ($LASTEXITCODE -ne 0) {
    Write-Warning "Windows Update cannot be started. Please contact service"
    pause
    return
  }
} else {
  Write-Output "Windows Update enabled"
}

Start-Sleep -Seconds 1

# CHECK INTERNET CONNECTION
if ((Test-Connection -ComputerName 8.8.8.8 -count 1 -quiet) -ne $true) {
  Write-Warning "No internet connection. Please connect the computer to internet to continue."
  pause
  if ((Test-Connection -ComputerName 8.8.8.8 -count 1 -quiet) -ne $true) {
    return
  }
}

# INSTALL WINGET COMMAND
# Install-VCLibs.140.00 Version 14.0.30035
$StoreLink = 'https://www.microsoft.com/de-de/p/app-installer/9nblggh4nns1'
$StorePackageName = 'Microsoft.VCLibs.140.00.UWPDesktop_14.0.30035.0_x64__8wekyb3d8bbwe.appx'

$RepoName = 'AppPackages'
$RepoLokation = $env:Temp
$Packagename = 'Microsoft.VCLibs.140.00'
$RepoPath = Join-Path $RepoLokation -ChildPath $RepoName 
$RepoPath = Join-Path $RepoPath -ChildPath $Packagename


#
# Function Source
# Idea from: https://serverfault.com/questions/1018220/how-do-i-install-an-app-from-windows-store-using-powershell
# modificated version. Now able to filte and return msix url's
#
function DownloadAppPackage {
[CmdletBinding()]
param (
  [string]$Uri,
  [string]$Filter = '.*' #Regex
)
   
  process {

    #$Uri=$StoreLink

    $WebResponse = Invoke-WebRequest -UseBasicParsing -Method 'POST' -Uri 'https://store.rg-adguard.net/api/GetFiles' -Body "type=url&url=$Uri&ring=Retail" -ContentType 'application/x-www-form-urlencoded'
    
    $result =$WebResponse.Links.outerHtml | Where-Object {($_ -like '*.appx*') -or ($_ -like '*.msix*')} | Where-Object {$_ -like '*_neutral_*' -or $_ -like "*_"+$env:PROCESSOR_ARCHITECTURE.Replace("AMD","X").Replace("IA","X")+"_*"} | ForEach-Object {
       $result = "" | Select-Object -Property filename, downloadurl
       
       if( $_ -match '(?<=rel="noreferrer">).+(?=</a>)' )
       {
         $result.filename = $matches.Values[0]
       }

       if( $_ -match '(?<=a href=").+(?=" r)' )
       {
         $result.downloadurl = $matches.Values[0]
       }
       $result
    } 
    
    
    $result | Where-Object -Property filename -Match $filter 
  }
}



$package = DownloadAppPackage -Uri $StoreLink  -Filter $StorePackageName

if(-not (Test-Path $RepoPath ))
{
    New-Item $RepoPath -ItemType Directory -Force
}


if(-not (Test-Path (Join-Path $RepoPath -ChildPath $package.filename )))
{
    Write-Information "Repository download..."
    Invoke-WebRequest -Uri $($package.downloadurl) -OutFile (Join-Path $RepoPath -ChildPath $package.filename )
} else 
{
    Write-Information "The file $($package.filename) already exist in the repo. Skip download"
}

#Install the Runtime
Write-Information "Installing runtime..."
Add-AppPackage (Join-Path $RepoPath -ChildPath $package.filename )


# Install-Winget Version v1.0.11692
# From github
$WinGet_Link = 'https://github.com/microsoft/winget-cli/releases/download/v1.0.11692/Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle'
$WinGet_Name = 'Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle'

$RepoName = 'AppPAckages'
$RepoLokation = $env:Temp
$Packagename = 'Winget'

$RepoPath = Join-Path $RepoLokation -ChildPath $RepoName 
$RepoPath = Join-Path $RepoPath -ChildPath $Packagename

if(-not (Test-Path $RepoPath ))
{
    New-Item $RepoPath -ItemType Directory -Force
}


if(-not (Test-Path (Join-Path $RepoPath -ChildPath  $WinGet_Name )))
{
    Write-Information "Desktop app installer download..."
    Invoke-WebRequest -Uri $WinGet_Link -OutFile (Join-Path $RepoPath -ChildPath $WinGet_Name )
} else 
{
    Write-Information "The file $WinGet_Name already exist in the repo. Skip download"
}

#Install the Package
Write-Information "Installing Winget..."
Add-AppPackage (Join-Path $RepoPath -ChildPath $WinGet_Name)

Write-Output "Installing Powershell core..."

# INSTALL POWERSHELL CORE
winget install --id Microsoft.Powershell --source 'winget'
if  ($LASTEXITCODE -ne 0) {
  Write-Warning "Powershell core not installed. Please retry, continuing to new Windows Terminal install..."
  pause
} else {
  Write-Output "PowerShell core installed. Continuing to new Windows Terminal install..."
}


# INSTALL NEW WINDOWS TERMINAL

if ([System.Environment]::OSVersion.Version.Build -ge 19041) {
	winget install --id Microsoft.WindowsTerminal --source 'winget'
} else {
  Write-Warning "The new Windows terminal cannot be installed. Continuing to MailKit install..."
}

# REGISTRATION NUGET.ORG TO DOWNLOAD MAILKIT PACKET
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
register-packagesource -Name 'NuGet.org' -Location https://www.nuget.org/api/v2 -ProviderName Nuget -Trusted -Force

# INSTALL MAILKIT & MIMEKIT PACKETS
install-package -Name 'MailKit' -Source 'NuGet.org' -SkipDependencies -Force
install-package -Name 'MimeKit' -Source 'NuGet.org' -SkipDependencies -Force

# CHECKING INSTALLATION OF THE PACKETS
$GetMailKit = Get-Package -Name MailKit
$GetMimeKit = Get-Package -Name MimeKit
if ($GetMailKit) {
  $GetMailKitVer = $GetMailKit.Version
  Write-Output "MailKit $GetMailKitVer successfully installed!"
  $CheckInstall = $CheckInstall + 1
} else {
  Write-Warning "MailKit not installed!"
}
if ($GetMimeKit) {
  $GetMimeKitVer = $GetMimeKit.Version
  Write-Output "MimeKit $GetMimeKitVer successfully installed!"
  $CheckInstall = $CheckInstall + 1
} else {
  Write-Warning "MimeKit not installed!"
}

# BLOCKING WINDOWS UPDATE
if ($DisableUpdateAtTheEnd -ne 0) {
  Write-Output "Reset Windows Update Service"
  NET STOP WUAUSERV
}

# END MESSAGE
if ($CheckInstall -ge 2) {
  Write-Output "Installation completed. Please send a test mail."
} else {
  Write-Warning "Installation failed. Please retry."
}

pause