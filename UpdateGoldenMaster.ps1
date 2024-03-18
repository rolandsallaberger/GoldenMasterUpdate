Install-Module -Name PSWindowsUpdate -Force
Get-WUList
Install-WindowsUpdate -MicrosoftUpdate -AcceptAll -IgnoreReboot 
