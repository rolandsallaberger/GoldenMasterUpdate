# Install Windows Updates
Install-Module -Name PSWindowsUpdate -Force
Get-WUList
Install-WindowsUpdate -MicrosoftUpdate -AcceptAll -IgnoreReboot 

# Install O365 Update
& 'C:\Program Files\Common Files\microsoft shared\ClickToRun\OfficeC2RClient.exe' /update user displaylevel=false forceappshutdown=true

# Install Winget
# Create WinGet Folder
New-Item -Path C:\WinGet -ItemType directory -ErrorAction SilentlyContinue

# Install VCLibs
Invoke-WebRequest -Uri https://aka.ms/Microsoft.VCLibs.x64.14.00.Desktop.appx -OutFile "C:\WinGet\Microsoft.VCLibs.x64.14.00.Desktop.appx"
Add-AppxPackage "C:\WinGet\Microsoft.VCLibs.x64.14.00.Desktop.appx"

# Install Microsoft.UI.Xaml from NuGet
Invoke-WebRequest -Uri https://www.nuget.org/api/v2/package/Microsoft.UI.Xaml/2.8.6 -OutFile "C:\WinGet\Microsoft.UI.Xaml.2.8.6.zip"
Expand-Archive "C:\WinGet\Microsoft.UI.Xaml.2.8.6.zip" -DestinationPath "C:\WinGet\Microsoft.UI.Xaml.2.8.6"
Add-AppxPackage "C:\WinGet\Microsoft.UI.Xaml.2.8.6\tools\AppX\x64\Release\Microsoft.UI.Xaml.2.8.appx"

# Install latest WinGet from GitHub
Invoke-WebRequest -Uri https://github.com/microsoft/winget-cli/releases/latest/download/Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle -OutFile "C:\WinGet\Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle"
Add-AppxPackage "C:\WinGet\Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle"

# Fix Permissions
TAKEOWN /F "C:\Program Files\WindowsApps" /R /A /D Y
ICACLS "C:\Program Files\WindowsApps" /grant Administrators:F /T

# Add Environment Path
$ResolveWingetPath = Resolve-Path "C:\Program Files\WindowsApps\Microsoft.DesktopAppInstaller_*_x64__8wekyb3d8bbwe"
if ($ResolveWingetPath) {
         $WingetPath = $ResolveWingetPath[-1].Path
}
$ENV:PATH += ";$WingetPath"
$SystemEnvPath = [System.Environment]::GetEnvironmentVariable('PATH', [System.EnvironmentVariableTarget]::Machine)
$SystemEnvPath += ";$WingetPath;"
setx /M PATH "$SystemEnvPath"

TAKEOWN /F "C:\Program Files\WindowsApps" /R /A /D Y
ICACLS "C:\Program Files\WindowsApps" /grant Administrators:F /T

#Update Apps per Winget - Excluding O365 (does not accept Channel)
winget pin add  --id "Microsoft.Office"
winget pin list
winget upgrade
winget upgrade --all
