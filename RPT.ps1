########################################
## Created By:    Luke Encrapera      ##                             
## Email:   Luke.Encrapera@DCSG.com   ##                               
## Date: 11/1/2018                    ##               
## Version: 1.0                       ##            
## Description - Remote Printer Tool  ##                                                                
########################################


function Get-PrintJob($arg1, $arg2){
	cscript.exe C:\Windows\System32\Printing_Admin_Scripts\en-US\prnjobs.vbs -l -s $arg1 -p $arg2
}
function Stop-PrintJob($arg1, $arg2, $arg3){
	cscript.exe C:\Windows\System32\Printing_Admin_Scripts\en-US\prnjobs.vbs -x -p $arg1 -j $arg2 -s $arg3
}
function Stop-AllPrintJobs($arg1, $arg2){
	cscript.exe C:\Windows\System32\Printing_Admin_Scripts\en-US\prnqctl.vbs -x -p $arg1 -s $arg2
}
function Get-TestPage($arg1, $arg2){
	cscript.exe C:\Windows\System32\Printing_Admin_Scripts\en-US\prnqctl.vbs -e -s $arg1 -p $arg2
}
function Get-Printers($arg1){
	cscript.exe C:\Windows\System32\Printing_Admin_Scripts\en-US\prnmngr.vbs -l -s $arg1
}
function Remove-Printer($arg1, $arg2){
	cscript.exe C:\Windows\System32\Printing_Admin_Scripts\en-US\prnmngr.vbs -d -p $arg1 -s $arg2
}

Set-Location C:\Windows\System32\Printing_Admin_Scripts\en-US
$MachineName = Read-Host "Enter Machine Name "
$UserChoice = 0
While ($True)
{
Write-Host " Check installed printers [1],`n Check and set a default printer [2],`n Install new printer [3],`n Remove installed Printer('s) [4],`n Print Test Page [5],`n Clear Print Que [6],"
$UserChoice = Read-Host "Choose option  "

    #Check Installed printers on a machine.
        if ($UserChoice -eq 1){
            cscript PRNMNGR.vbs -l -s $MachineName
        }

    #Check and set a default printer for a machine.
        if ($UserChoice -eq 2){
            cscript prnmngr.vbs -g -s $MachineName
            # get the DEFAULT printer
            $ChangeDefault = Read-Host "Change default printer? y or n" 
            if ($ChangeDefault -eq "y"){
                # Set the DEFAULT printer
                $PrinterName =   Read-Host "Enter NAME of the Printer you want to set as Default "
#                cscript prnmngr.vbs -t -s $MachineName -p $PrinterIP
                (Get-WmiObject -ComputerName $MachineName -Class Win32_Printer -Filter "Name='HP Color'").SetDefaultPrinter($PrinterIP)
#                (New-Object -ComObject WScript.Network).SetDefaultPrinter("$PrinterName")
#                $PrinterName.setDefaultPrinter()
            }
            if ($ChangeDefault -eq "n"){
                Write-Host "No changes were made!"
            }
        }

    #Install a new printer to a device.
        if ($UserChoice -eq 3){
            $NetUSB = Read-Host "Networked type [n], USB type [u]"
            if ($NetUSB -eq "n"){
                $PrinterIP = Read-Host "Enter IP Address of the Printer "
                $PrinterIPLocation = Read-Host "Name The Printer "
                $Driver = [pscustomobject][ordered]@{
                Alias = "HP402"
                Driver = "HP LaserJet Pro M402-M403 n-dn PCL 6"
                Description = "$PrinterIPLocation $($Driver.Alias)"
                }
                cscript PRNMNGR.vbs -a -p $PrinterIP -m Driver -r $PrinterIP
            }
            if ($NetUSB -eq "u"){
                Rundll32 printui.dll PrintUIEntry /ip /u /z /c\\$MachineName
            }  
        }

    #Delete an installed printer from the machine.
        if ($UserChoice -eq 4){
            $PrinterIP = Read-Host "Enter IP Address of the Printer to remove. Enter * to delete all installed printers "
            cscript PRNMNGR.vbs -d -s $MachineName -p $PrinterIP
            #Deletes all printers from a device.
            if ($PrinterIP -eq '*'){
                cscript PRNMNGR.vbs -x -s $MachineName
            }
        }

    #Print Test page.
	    if($type -eq 5){
		    Write-Host "Start Remote Registry Service";
		    Start-Service -InputObject $(Get-Service -Computer $MachineName -Name "Remote Registry");
		    Write-host "Add Key";
		    reg add "\\$MachineName\HKLM\SOFTWARE\Policies\Microsoft\Windows NT\Printers" /v RegisterSpoolerRemoteRpcEndPoint /t REG_DWORD /d 1
		    Write-host "Restart Service";
		    Restart-Service -force -InputObject $(Get-Service -Computer $MachineName -Name "Print Spooler");
		    $out = Get-WmiObject win32_printer -ComputerName $MachineName | select caption,Portname
		    #$out = Get-Printers $MachineName
		    for($i=0; $i -lt $out.count; $i++){
			    write-host "[$i] " $out[$i].caption;
		    }
		    $PrinterIP = Read-Host "Printer:";
		    #Rundll32 printui.dll PrintUIEntry /k /n\\$MachineName\$PrinterIP
		    Get-TestPage $MachineName $out[$PrinterIP].caption
		    Write-host "Remove Key";
		    reg delete "\\$MachineName\HKLM\SOFTWARE\Policies\Microsoft\Windows NT\Printers" /v RegisterSpoolerRemoteRpcEndPoint /f
		    #Clear-Host
		    }

    #Clear print que
        if($type -eq 6){ 
		    Write-Host "Start Remote Registry Service";
		    Start-Service -InputObject $(Get-Service -Computer $MachineName -Name "Remote Registry");
		    Write-host "Add Key";
		    reg add "\\$MachineName\HKLM\SOFTWARE\Policies\Microsoft\Windows NT\Printers" /v RegisterSpoolerRemoteRpcEndPoint /t REG_DWORD /d 1
		    Write-host "Restart Service";
		    Restart-Service -force -InputObject $(Get-Service -Computer $MachineName -Name "Print Spooler");
		    $out = Get-WmiObject win32_printer -ComputerName $MachineName | select caption,Portname
		    for($i=0; $i -lt $out.count; $i++){
			    write-host "[$i] " $out[$i].caption;
		    }
		
		    $PrinterIP = Read-Host "Printer:";
		    Get-PrintJob $MachineName $PrinterIP
		    #Rundll32 printui.dll PrintUIEntry /o /n\\$MachineName\$PrinterIP
		    Write-host "Remove Key";
		    reg delete "\\$MachineName\HKLM\SOFTWARE\Policies\Microsoft\Windows NT\Printers" /v RegisterSpoolerRemoteRpcEndPoint /f
		}
}
$MachineName.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | OUT-NULL
$MachineName.UI.RawUI.Flushinputbuffer()