# Install Windows Updates
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Install-Module -Name PSWindowsUpdate -Force
Get-WUList
Install-WindowsUpdate -MicrosoftUpdate -AcceptAll -IgnoreReboot 

# Install O365 Update
& 'C:\Program Files\Common Files\microsoft shared\ClickToRun\OfficeC2RClient.exe' /update user displaylevel=true forceappshutdown=true

#*******************************
# Install Chocolately Package Manager
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
Set-ExecutionPolicy Bypass -Scope Process -Force; iex ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))
$env:Path = [System.Environment]::ExpandEnvironmentVariables([System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User"))

# Install WinGet from Chocolatey
choco install winget
choco install winget.powershell

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




#********************
Write-Information "This script needs be run on Windows Server 2019 or 2022"

If ($PSVersionTable.PSVersion.Major -ge 7){ Write-Error "This script needs be run by version of PowerShell prior to 7.0" }

# Define environment variables
$downloadDir = "C:\WinGet"
$gitRepo = "microsoft/winget-cli"
$msiFilenamePattern = "*.msixbundle"
$licenseFilenamePattern = "*.xml"
$releasesUri = "https://api.github.com/repos/$gitRepo/releases/latest"

# Preparing working directory
New-Item -Path $downloadDir -ItemType Directory
Push-Location $downloadDir

# Downloaing artifacts
function Install-Package {
    param (
        [string]$PackageFamilyName
    )

    Write-Host "Querying latest $PackageFamilyName version and its dependencies..."
    $response = Invoke-WebRequest `
        -Uri "https://store.rg-adguard.net/api/GetFiles" `
        -Method "POST" `
        -ContentType "application/x-www-form-urlencoded" `
        -Body "type=PackageFamilyName&url=$PackageFamilyName&ring=RP&lang=en-US" -UseBasicParsing

    Write-Host "Parsing response..."
    $regex = '<td><a href=\"([^\"]*)\"[^\>]*\>([^\<]*)<\/a>'
    $packages = (Select-String $regex -InputObject $response -AllMatches).Matches.Groups

    $result = $true
    for ($i = $packages.Count - 1; $i -ge 0; $i -= 3) {
        $url = $packages[$i - 1].Value;
        $name = $packages[$i].Value;
        $extCheck = @(".appx", ".appxbundle", ".msix", ".msixbundle") | % { $x = $false } { $x = $x -or $name.EndsWith($_) } { $x }
        $archCheck = @("x64", "neutral") | % { $x = $false } { $x = $x -or $name.Contains("_$($_)_") } { $x }

        if ($extCheck -and $archCheck) {
            # Skip if package already exists on system
            $currentPackageFamilyName = (Select-String "^[^_]+" -InputObject $name).Matches.Value
            $installedVersion = (Get-AppxPackage "$currentPackageFamilyName*").Version
            $latestVersion = (Select-String "_(\d+\.\d+.\d+.\d+)_" -InputObject $name).Matches.Value
            if ($installedVersion -and ($installedVersion -ge $latestVersion)) {
                Write-Host "${currentPackageFamilyName} is already installed, skipping..." -ForegroundColor "Yellow"
                continue
            }

            try {
                Write-Host "Downloading package: $name"
                $tempPath = "$(Get-Location)\$name"
                Invoke-WebRequest -Uri $url -Method Get -OutFile $tempPath
                Add-AppxPackage -Path $tempPath
                Write-Host "Successfully installed:" $name
            } catch {
                $result = $false
            }
        }
    }

    return $result
}

Write-Host "`n"

function Install-Package-With-Retry {
    param (
        [string]$PackageFamilyName,
        [int]$RetryCount
    )

    for ($t = 0; $t -le $RetryCount; $t++) {
        Write-Host "Attempt $($t + 1) out of $RetryCount..." -ForegroundColor "Cyan"
        if (Install-Package $PackageFamilyName) {
            return $true
        }
    }

    return $false
}

$licenseDownloadUri = ((Invoke-RestMethod -Method GET -Uri $releasesUri).assets | Where-Object name -like $licenseFilenamePattern ).browser_download_url
$licenseFilename = ((Invoke-RestMethod -Method GET -Uri $releasesUri).assets | Where-Object name -like $licenseFilenamePattern ).name
$licenseJoinPath = Join-Path -Path $downloadDir -ChildPath $licenseFilename
Invoke-WebRequest -Uri $licenseDownloadUri -OutFile ( New-Item -Path $licenseJoinPath -Force )

$result = @("Microsoft.DesktopAppInstaller_8wekyb3d8bbwe") | ForEach-Object { $x = $true } { $x = $x -and (Install-Package-With-Retry $_ 3) } { $x }

$msiFilename = ((Get-ChildItem -Path $downloadDir) | Where-Object name -like $msiFilenamePattern ).name
$msiJoinPath = Join-Path -Path $downloadDir -ChildPath $msiFilename

# Installing winget
Add-ProvisionedAppPackage -Online -PackagePath $msiJoinPath -LicensePath $licenseJoinPath -Verbose

Write-Host "`n"

# Test if winget has been successfully installed
if ($result -and (Test-Path -Path "$env:LOCALAPPDATA\Microsoft\WindowsApps\winget.exe")) {
    Write-Host "Congratulations! Windows Package Manager (winget) $(winget --version) installed successfully" -ForegroundColor "Green"
} else {
    Write-Host "Oops... Failed to install Windows Package Manager (winget)" -ForegroundColor "Red"
}

# Cleanup
Push-Location $HOME
Remove-Item $downloadDir -Recurse -Force
#********************************************************


# Install Winget
# Create WinGet Folder
New-Item -Path C:\WinGet -ItemType directory -ErrorAction SilentlyContinue

# Install VCLibs
Invoke-WebRequest -Uri https://aka.ms/Microsoft.VCLibs.x64.14.00.Desktop.appx -OutFile "C:\WinGet\Microsoft.VCLibs.x64.14.00.Desktop.appx"
Add-AppxPackage "C:\WinGet\Microsoft.VCLibs.x64.14.00.Desktop.appx"

# Install Microsoft.UI.Xaml from NuGet
Invoke-WebRequest -Uri https://www.nuget.org/api/v2/package/Microsoft.UI.Xaml/2.8.6 -OutFile "C:\WinGet\Microsoft.UI.Xaml.2.8.6.zip"
Expand-Archive "C:\WinGet\Microsoft.UI.Xaml.2.8.6.zip" -DestinationPath "C:\WinGet\Microsoft.UI.Xaml.2.8.6" -force
Add-AppxPackage "C:\WinGet\Microsoft.UI.Xaml.2.8.6\tools\AppX\x64\Release\Microsoft.UI.Xaml.2.8.appx"

# Install Winget
# Create WinGet Folder
New-Item -Path C:\WinGet -ItemType directory -ErrorAction SilentlyContinue

# Install latest WinGet from GitHub
Invoke-WebRequest -Uri https://github.com/microsoft/winget-cli/releases/latest/download/Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle -OutFile "C:\WinGet\Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle"
Add-AppxPackage "C:\WinGet\Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle"
#

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
