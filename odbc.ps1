
Write-Host "Checking Windows Version"
$windowsVersion = (Get-WmiObject -class Win32_OperatingSystem).Caption
if($windowsVersion -like "*Windows 10*"){
    Write-Host "Current Machine Running Windows 10. Continuing"
} else {
    if($windowsVersion -like "*Windows 7*"){
        Write-Host "Current Machine does not support this script, requires Windows 10, or non garbage PC" -ForegroundColor "DarkRed" -BackgroundColor "Black"
        Write-Host "Path to database: PATH_TO_DATABASE"
        Write-Host "Database File: WorkstationDatabase.mdb"
        Start-Process -FilePath "C:\Windows\SysWOW64\odbcad32.exe"
        
        exit
    }
}
#stage 1
Write-Host "Checking for existing Database connection"
$dsn = Get-OdbcDsn -Name "Brentwood Network Database" -DsnType "System" -Platform "32-bit" -ErrorVariable ProcessError -ErrorAction SilentlyContinue
if($ProcessError){
    #the network database does not exist, so we can just create it 
    Write-Host "ODBC Connection for Midmark Databse not present on system" -ForegroundColor "Yellow" -BackgroundColor "Black"
    Write-Host "Adding Database connection for Midmark Client" -ForegroundColor "Green"
    Add-OdbcDSN -Name "Brentwood Network Database" -Platform "32-bit" -DriverName "Microsoft Access Driver (*.mdb)" -DsnType "System" -SetPropertyValue "Dbq=PATH_TO_DATABASE"
} else {
    #the network database already exists, so we must remove it before adding the new one
    if($dsn.PropertyValue -like "*Q:\*"){
        Write-Host "ODBC Connection Set to use Mapped Drive" -BackgroundColor "DarkRed"
    } else {
        if($dsn.PropertyValue -like "*\\huggins-app03*"){
            Write-Host "ODBC Connection already pointed at \\huggins-app03" -BackgroundColor "Black" -ForegroundColor "Green"
        }
    }
    Write-Host "Removing Old ODBC Connection"
    Remove-OdbcDsn -Name "Brentwood Network Database" -DsnType "System" -Platform "32-bit"
    #stage 2
    Write-Host "Adding Database connection for Midmark Client" -ForegroundColor "Green"
    Add-OdbcDSN -Name "Brentwood Network Database" -Platform "32-bit" -DriverName "Microsoft Access Driver (*.mdb)" -DsnType "System" -SetPropertyValue "Dbq=PATH_TO_DATABASE"
    Write-Host "Midmark ODBC Connection Added." -ForegroundColor "Green"
}

Write-host "New Connection Settings: "
Get-OdbcDsn -Name "Brentwood Network Database" -DsnType "System" -Platform "32-bit" -ErrorVariable ProcessError -ErrorAction SilentlyContinue

#Set-OdbcDsn -Name "Brentwood Workstation Database" -DsnType "System" -Platform "32-bit" -SetPropertyValue "Dbq=PATH_TO_DATABASE"
pause
