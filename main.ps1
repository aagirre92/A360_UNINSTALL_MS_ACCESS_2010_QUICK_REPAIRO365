<#
Author: Andoni Aguirre Aranguren

Contact: andoni_aguirre_aranguren@epam.com

Date: 14-Feb-2023

Description: This powershell script will:
    0) Uninstall "Microsoft Access Database Engine 2010" (32-bit)
	1) Quick repair O365 


Considerations:
    1) You must run this script as local admin user

#>

Try {

    # Uninstall Ms acess db engine 2010
    $ms_access_db_engine_2010 = Get-WmiObject -Class Win32_Product | Where-Object { $_.Name -eq "Microsoft Access database engine 2010 (English)" }
    if (!$ms_access_db_engine_2010) { 
        Write-Output "Microsoft Access database engine 2010 (English) NOT installed in this host"
    }
    else {
        Write-Host "Proceeding to uninstall Microsoft Access database engine 2010 (English)..."
        $ms_access_db_engine_2010.Uninstall()
        Write-Output "Microsoft Access database engine 2010 (English) uninstalled"
    }

    # Quick repair office 365 32-bit
    Set-Location "C:\Program Files\Common Files\microsoft shared\ClickToRun"
    .\OfficeClickToRun.exe scenario=Repair DisplayLevel=false RepairType=quickRepair forceappshutdown=true
    Write-Output "O365 Quick repair successful"
    Write-Output "THE END - SUCCESS"

}

Catch {
    Write-Output $Error[0]
    Write-Output "THE END - FAIL"
}