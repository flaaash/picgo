###########################################################################
Write-Host "`n`nCurrent installed Copiers:" -ForegroundColor Yellow
# 1. 获取基础信息
Get-PrinterDriver | Where-Object { $_.Manufacturer -match "hp|M480F" } | Select-Object Name, Manufacturer, Provider, PrinterEnvironment
Get-PrinterPort | Where-Object { $_.PortMonitor -eq "TCPMON.DLL" } | Select-Object Name, PrinterHostAddress, PortNumber
Get-Printer | Where-Object { $_.Name -match "Canon|M480F" } | Select-Object Name, PortName, PrinterStatus, Priority


# 2. Remove old printers configuration
Remove-Printer -Name CanonC5550  2>$null
Remove-PrinterPort -Name CanonC5550 2>$null
Remove-Printer -Name CanonC5560 2>$null
Remove-PrinterPort -Name CanonC5560 2>$null

# 3.1 Download Drivers
# $urlCanonC5560 = "https://drive.usercontent.google.com/download?id=1tL4ENh8e7vvze63GB9tabCA2j0KFV4pv&export=download&authuser=0&confirm=t&uuid=c35bb61d-64a8-4268-9c1b-869c2cb0029c&at=AENtkXbHr8XwULxgWaommVPGmo86%3A1730826006974"
# $urlHP480F = "https://drive.google.com/file/d/1ofTiq_x7-syzHwNfZQN_1bD2jJ6M6tGH/view?usp=sharing"
# $zipCanonC5560 = "C:\Drivers\CNS30MA64.zip"
# $zipHP480F = "C:\Drivers\hpcad92a4_x64.zip"



# New-Item -Path "C:\Drivers" -ItemType Directory -Force
# Invoke-WebRequest -Uri $urlCanonC5560 -OutFile $zipCanonC5560
# Invoke-WebRequest -Uri $urlHP480F -OutFile $zipHP480F
# # unzip
# Expand-Archive -Path $zipCanonC5560 -DestinationPath "C:\Drivers\CNS30MA64" -Force
# Expand-Archive -Path $zipHP480F -DestinationPath "C:\Drivers\hpcad92a4_x64" -Force

    
## 3.2 Canon5560
Write-Host "`n`nInstalling C5560..." -ForegroundColor Green
# Drivers
$driverName = "Canon Generic Plus PS3"
$infPath = "d:\Software\4.Drivers\CanonC5560\GPlus_PS3_Driver_V310_W64_00\Driver\CNS30MA64.INF"
pnputil /add-driver $infPath /install
Add-PrinterDriver -Name $driverName -Verbose
# Printer
$printerName = "Canon5560"
$driverName = "Canon Generic Plus PS3"

Add-PrinterPort -Name $printerName -PrinterHostAddress "172.22.23.91" 2>$null
Start-Sleep 2
Add-Printer -Name $printerName -DriverName $driverName -PortName $printerName  -Location "Unit2200" -Comment $printerName 2>$null


## 4. HP M480F
$driverName = "HP Color LaserJet MFP M480 PCL-6 (V4)" 
$infPath = "d:\Software\4.Drivers\HP-M480F\HP_CLJM480_V4\hpcad92a4_x64.inf"
pnputil /add-driver $infPath /install
Add-PrinterDriver -Name $driverName -Verbose
# M480F-West
$printerName = "M480F-West"
Write-Host "`n`nInstalling $printerName..." -ForegroundColor Green
Add-PrinterPort -Name $printerName -PrinterHostAddress "172.22.23.92" 2>$null
Start-Sleep 2
Add-Printer -Name $printerName -DriverName $driverName -PortName $printerName  -Location "HR & Admin Team" -Comment $printerName 2>$null
# M480F-East 
$printerName = "M480F-East"
Write-Host "Installing $printerName..." -ForegroundColor Green
Add-PrinterPort -Name $printerName -PrinterHostAddress "172.22.23.93" 2>$null
Start-Sleep 2
Add-Printer -Name $printerName -DriverName $driverName -PortName $printerName  -Location "IR Team" -Comment $printerName 2>$null



# set CanonC5560 as default
rundll32 printui.dll, PrintUIEntry /y /n "Canon5560"

# 4. Output installation result
Write-Host "`n`nYou have installed Copiers below:" -ForegroundColor Yellow
Get-Printer | Where-Object { $_.Name -match "Canon|M480F" } | Select-Object Name, PortName, PrinterStatus, Priority

Write-Host "Press any key to continue..."  -ForegroundColor Yellow
[System.Console]::ReadKey() | Out-Null