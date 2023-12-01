# installs noah 4.14.2
# app version: 4.3.0
# build version: 5678

# variables
$Storage_Needed = 1000000000
$Version_Current = (Get-Package "Noah 4" | Select-Object -Property Version).Version
$Version_Newest = [version]"4.3.0"
$Version_Minimum = [version]"2.95.0"
$Dir= "C:\Temp"
$Dir_Full = "$Dir\noah_4.14.2"
$Installer_URI = "https://himsafiles.com/Noah4downloads/Noah_System_4.14.2.5678.exe"
$Requirements_URI = "https://download.visualstudio.microsoft.com/download/pr/2d6bb6b2-226a-4baa-bdec-798822606ff1/8494001c276a4b96804cde7829c04d7f/ndp48-x86-x64-allos-enu.exe"


# ---/checks/---

# current ver > install ver
if ($Version_Current -ge $Version_Newest) {
    throw "Current version is same or newer than $Version_Newest"}

# current version at least 2.95.0
if ($Version_Current -lt $Version_Minimum) {
    throw "Current version is below required version. Upgrade to $Version_Minimum first, then run this script again. Installed Version: $Version_Current"}

# enough space to install?
if ((Get-Volume -DriveLetter C | Select-Object -ExpandProperty SizeRemaining) -lt $Storage_Needed) {
		throw "Not enough space on C drive to install."}
# --------------


# backup HIMSA data
Copy-Item -Recurse C:\ProgramData\HIMSA C:\ProgramData\HIMSA_BACKUP_$(get-date -f "yyyyMMdd.hhmm")

# create dir for everything to go to
New-Item -Path $Dir -ItemType "directory"
New-Item -Path $Dir -Name noah_4.14.2 -ItemType "directory"

# check if .net version > 4.6, install if not
$Requirements = (Get-ChildItem 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full' -Recurse | Get-ItemProperty -Name version -EA 0 | Select Version).Version
if ($Requirements -lt [version]"4.6.0") {
    Invoke-WebRequest -Uri $Requirements_URI -OutFile "$Dir_Full\ndp48-x86-x64-allos-enu.exe"
    Start-Process -FilePath "$Dir_Full\ndp48-x86-x64-allos-enu.exe" -ArgumentList "/q /norestart"
    }

# download and install noah
Invoke-WebRequest -Uri $Installer_URI -OutFile "$Dir_Full\Noah4.14.2.5678MSIandMST.zip"
Expand-Archive "$Dir_Full\Noah4.14.2.5678MSIandMST.zip" -DestinationPath "$Dir_Full\Noah4.14.2.5678MSIandMST"
msiexec.exe /i "$Dir_Full\Noah4.14.2.5678MSIandMST\Noah 4.msi" EULAACCEPTED=YES TRANSFORMS="1033.MST" /qn /norestart /l*v "log.log"

# cleanup if necessary
Remove-Item -Recurse -Path $Dir_Full