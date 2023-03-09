' Declare necessary variable
dim oShell, sProf, oFolder
dim oFSO, sProfRoot, oProfFolder
dim sWindows, sSysDrive, sDateToday, sUserTemp
dim OSType, strError, output
set oFSO=CreateObject("Scripting.FileSystemObject")
Set output = Wscript.stdout

'--------------------------------------------------

'catch errors in a log file

Sub LogError (strError) 
output.writeline Err.Number & " " & Err.Description & " " & strError
Err.Clear
End Sub

'Get Space of C:\ Drive
Sub Disk_space()
Set objWMIService = GetObject("winmgmts:\\localhost\root\CIMV2")
Set colItems = objWMIService.ExecQuery( _
"SELECT * FROM Win32_LogicalDisk WHERE DeviceID = 'C:'",,48)
For Each objItem in colItems
Bytes = objItem.FreeSpace
If Bytes >= 1073741824 Then
SetBytes = Round(FormatNumber(Bytes / 1024 / 1024 / 1024, 2), 0) & " GB"
ElseIf Bytes >= 1048576 Then
SetBytes = Round(FormatNumber(Bytes / 1024 / 1024, 2), 0) & " MB"
ElseIf Bytes >= 1024 Then
SetBytes = Round(FormatNumber(Bytes / 1024, 2), 0) & " KB"
ElseIf Bytes < 1024 Then
SetBytes = Bytes & " Bytes"
Else
SetBytes = "0 Bytes"
End If
Wscript.Echo " " & SetBytes &" Free Space"
Next
End sub
Sub DeleteThisFolder(FolderName)
If FSO.FolderExists(FolderName) Then
objshell.Run "CMD.EXE /C RD /S /Q """ & FolderName & """",0,True
End If
end Sub

'Reading registry
Function ReadReg(RegPath)
Dim objRegistry, Key
Set objRegistry = CreateObject("Wscript.shell")
Key = objRegistry.RegRead(RegPath)
ReadReg = Key
Set objRegistry = Nothing
End Function

'Delete Files on folders
public sub DeleteFolderContents(strFolder)
dim oFolder, objFile, objSubFolder
on error resume next
set oFolder=oFSO.GetFolder(strFolder)
if Err.Number0 then
end if
for each objSubFolder in oFolder.SubFolders
objSubFolder.Delete true
if Err.Number0 then
Err.Clear
DeleteFolderContents(strFolder & "\" & objSubFolder.Name)
'If Err Then Call LogError(strFolder & "\" & objSubFolder.Name)
end if
next
for each objFile in oFolder.Files
objFile.Delete true
'if Err.Number0 then Call LogError (objFile) ' In case we couldn't delete a file
next
end sub

'-------------------------------------------------------

'Clean Temp file windows
Sub Clean_temp_windows ()
DeleteFolderContents("C:\Windows\Temp")
DeleteFolderContents ("C:\Windows\Logs\CBS")
DeleteFolderContents("C:\Windows\ServiceProfiles\LocalService\AppData\Local\Temp")
DeleteFolderContents("C:\Windows\ServiceProfiles\NetworkService\AppData\Local\Temp")
DeleteFolderContents("C:\Windows\System32\DriverStore\Temp")
DeleteFolderContents("C:\Windows\WinSxS\Temp")
DeleteFolderContents("C:\Windows\SoftwareDistribution")
end Sub

'Clear profile Files and Registry
Sub Clean_profiles()
sUserTemp=userProfile & "\AppData\Local\Temp"
set oShell = CreateObject("WScript.Shell")
sWindows = oShell.ExpandEnvironmentStrings("%WINDIR%")
sSysDrive = oShell.ExpandEnvironmentStrings("%SYSTEMDRIVE%")
sProfRoot = ReadReg("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\ProfilesDirectory")
sProfRoot = Replace (LCase(sProfRoot), "%systemdrive%", sSysDrive)
set oShell=nothing
set oProfFolder=oFSO.GetFolder(sProfRoot)
for each oFolder in oProfFolder.SubFolders
sProf=sProfRoot & "\" & oFolder.Name
DeleteFolderContents sProf & sUserTemp
''If Err Then Call LogError(sProf)
next
end sub



'Empty bin
Sub E_recycle_bin()
Const RECYCLE_BIN = &Ha&
Set FSO = CreateObject("Scripting.FileSystemObject")
Set ObjShell = CreateObject("Shell.Application")
Set oFolder = ObjShell.Namespace(RECYCLE_BIN)
Set oFolderItem = oFolder.Self
Set colItems = oFolder.Items
For Each objItem in colItems
If (objItem.Type = "File folder") Then
FSO.DeleteFolder(objItem.Path)
Else
FSO.DeleteFile(objItem.Path)
End If
Next
end Sub

'Clear Cach for IE (rest Google Chrome and Firefox Modzilla)
Public Function ClearIECache()
Dim objShell
Set objShell = CreateObject("WScript.Shell")
objShell.Run "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 8"
objShell.Run "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 2"
objShell.Run "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 1"
Set objShell = Nothing
End Function
'----------------------------------------------
'Registry cleanssing 
public Sub Cleanmgr ()
Dim strKeyPath
Dim strComputer
Dim objReg
Dim arrSubKeys
Dim SubKey
Dim strValueName
Const HKLM=&H80000002
strKeyPath="SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches"
strComputer="."
strValueName="StateFlags0001"
Set objReg=GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
objReg.Enumkey HKLM ,strKeyPath,arrSubKeys
For Each SubKey In arrSubKeys
objReg.SetDWORDValue HKLM,strKeyPath & "\" & SubKey,strValueName,2
Next
'' Launch Cleaner
Dim WshShell
Set WshShell=CreateObject("WScript.Shell")
WshShell.Run "C:\WINDOWS\SYSTEM32\cleanmgr /sagerun:1"
End Sub
Public Sub Clean_Chrome ()
sUserTemp=userProfile & "\Appdata\Local\Google\Chrome\User Data\Default\Cache"
' Get user profile root folder
set oShell = CreateObject("WScript.Shell")
sWindows = oShell.ExpandEnvironmentStrings("%WINDIR%")
sSysDrive = oShell.ExpandEnvironmentStrings("%SYSTEMDRIVE%")
sProfRoot = ReadReg("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\ProfilesDirectory")
'output.writeline "Raw Registry Read Result: " & sProfRoot
sProfRoot = Replace (LCase(sProfRoot), "%systemdrive%", sSysDrive)
'output.writeline "After Variable Replacement: " & sProfRoot
set oShell=nothing
set oProfFolder=oFSO.GetFolder(sProfRoot)
for each oFolder in oProfFolder.SubFolders
''output.writeline "Processing profile: " & sProfRoot & "\" & oFolder.Name & sUserTemp
sProf=sProfRoot & "\" & oFolder.Name
DeleteFolderContents sProf & sUserTemp
''If Err Then Call LogError(sProf)
next
end sub
'' Calls shell function to clear the hibernate file
Public Sub Clear_hiberfil()
set oShell = CreateObject("WScript.Shell")
oShell.run "powercfg.exe /hibernate off"
end Sub
'' Function to go sleep in case you need it
Sub GoToSleep (iMinutes)
Dim Starting, Ending, t
Starting = Now
Ending = DateAdd("n",iMinutes,Starting)
Do
t = DateDiff("n",Now,Ending)
If t <= 0 Then Exit Do
WScript.Sleep 10000
Loop
End Sub
''Clean old OST 'NOT FINISHED yet
Sub Clean_OST()
set oShell = CreateObject("WScript.Shell")
oShell.run ("powershell -noexit -file c:\Users\!sona20232\Desktop\ost.ps1") 
'oShell.run "ost.bat"'
end Sub


''' MAIN SCRIPT ''
Clean_profiles()
Clean_temp_windows()
E_recycle_bin()
ClearIECache()
Cleanmgr()
Clean_Chrome()
Clear_hiberfil()
Disk_space()
Clean_OST()
Wscript.echo "Free Space After Cleanup ( Windows Cleanup Disktool results not included )"