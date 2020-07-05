' --- KON-BOOT USB INSTALLER -------------------------------
' (c) kon-boot.com 
' ----------------------------------------------------------
Dim query 
Dim WMBIObj 
Dim AllDiskDrives 
Dim SingleDiskDrive 
Dim AllLogicalDisks 
Dim SingleLogicalDisk 
Dim AllPartitions 
Dim Partition 
Dim result
Dim textmsg
Dim wshShell
Dim FileObj
Dim Counter
Dim CDir
Dim Command
Dim Temp
Dim TempFile
Dim FileTxt

set FileObj = CreateObject("Scripting.FileSystemObject")
set wshShell = wscript.createObject("wscript.shell")
set WMBIObj = GetObject("winmgmts:\\.\root\cimv2")

Temp	= chr(34)

textmsg = 	"PLEASE DETACH ALL UNNECESSARY USB DRIVES EXCEPT THE TARGET USB FOR KON-BOOT" & VbCr & VbCr & _ 
		"ADDITIONALLY MAKE SURE THE .BAT FILE WAS RUN WITH ADMIN RIGHTS (RIGHT CLICK->RUN AS ADMINISTRATOR)" & VbCr & VbCr & _ 
		"CLICK OK TO CONTINUE"

MsgBox textmsg, vbInformation + vbOKOnly, "Kon-Boot USB Installer"


CDir = FileObj.GetAbsolutePathName(".")
TempFile = CDir & "\temp_file.txt"
'MsgBox TempFile, vbInformation + vbOKOnly, "Kon-Boot USB Installer"

Set AllDiskDrives = WMBIObj.ExecQuery("SELECT * FROM Win32_DiskDrive where InterfaceType='USB'") ' 
For Each SingleDiskDrive In AllDiskDrives 
	counter = counter + 1
    query = "ASSOCIATORS OF {Win32_DiskDrive.DeviceID='" + SingleDiskDrive.DeviceID + "'} WHERE AssocClass = Win32_DiskDriveToDiskPartition" 
    Set AllPartitions = WMBIObj.ExecQuery(query) 
    For Each Partition In AllPartitions 
        query = "ASSOCIATORS OF {Win32_DiskPartition.DeviceID='" + Partition.DeviceID + "'} WHERE AssocClass = Win32_LogicalDiskToPartition"
        Set AllLogicalDisks = WMBIObj.ExecQuery (query) 
        For Each SingleLogicalDisk In AllLogicalDisks  
			
			textmsg = 	"============================================" & VbCr & _ 
						"DeviceID: " & SingleDiskDrive.DeviceID & VbCr & _ 
						"Logical Drive: " & SingleLogicalDisk.DeviceID  & VbCr & _ 
						"Model: " & SingleDiskDrive.Model & VbCr & _ 
						"Manufacturer: " & SingleDiskDrive.Manufacturer & VbCr & _
						"============================================" & VbCr & _ 
						"Would you like to use this drive as destination?" & VbCr & _ 
						"Warning, disk data will be overwritten"
			result = MsgBox(textmsg, vbQuestion + vbOKCancel, "Kon-Boot USB Installer")
			
			if result = vbOk Then
			
			Command = 	"SELECT VOLUME " & SingleLogicalDisk.DeviceID & VbCrLf & _ 
						"CLEAN"  & VbCrLf & _ 
						"CREATE PARTITION PRIMARY"  & VbCrLf & _ 
						"SELECT PARTITION 1"  & VbCrLf & _ 
						"ACTIVE"  & VbCrLf & _ 
						"FORMAT FS=FAT32 Label=" & Temp & "KONBOOT" & Temp & " QUICK"  & VbCrLf & _
						"ASSIGN LETTER=" & SingleLogicalDisk.DeviceID & VbCrLf & _
						"EXIT" & VbCrLf
		
		  
				set FileTxt	=	FileObj.CreateTextFile(TempFile, True)
				FileTxt.WriteLine(Command)
				FileTxt.Close
				
				wshShell.Run "diskpart /s " & Temp & TempFile & Temp, 1, true
				FileObj.CopyFolder CDir & "\EFI", SingleLogicalDisk.DeviceID & "\", 1
				
				wshShell.Run "grubinst.exe --skip-mbr-test " & SingleDiskDrive.DeviceID, 1, true
			
				FileObj.CopyFile "USBFILES\grldr", SingleLogicalDisk.DeviceID & "\", 1
				FileObj.CopyFile "USBFILES\konboot.img", SingleLogicalDisk.DeviceID & "\", 1
				FileObj.CopyFile "USBFILES\menu.lst", SingleLogicalDisk.DeviceID & "\", 1
				FileObj.CopyFile "USBFILES\konbootOLD.img", SingleLogicalDisk.DeviceID & "\", 1
				
				
				MsgBox "Your Kon-Boot on USB is ready!", vbInformation + vbOKOnly, "Kon-Boot USB Installer"
				WScript.quit 
			End If
		Next
    Next 
Next 
if counter = 0 Then
	MsgBox "No USB disks detected or unknown error!", vbCritical + vbOkOnly, "Kon-Boot USB Installer"
End If