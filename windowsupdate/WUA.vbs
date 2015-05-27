Set ObjFSO = CreateObject("Scripting.FileSystemObject")
Set fileLog = ObjFSO.CreateTextFile("c:\windowsupdatelog.txt")

Set updateSession = CreateObject("Microsoft.Update.Session")
updateSession.ClientApplicationID = "Wemade SE winupdate script"

Set updateSearcher = updateSession.CreateUpdateSearcher()

fileLog.writeline "[" & NOW &"] " & "Searching for updates... ==========================" 

Set searchResult = _
updateSearcher.Search("IsInstalled=0 and Type='Software' and IsHidden=0")

fileLog.writeline "[" & NOW &"] " & "List of applicable items on the machine: =========="

For I = 0 To searchResult.Updates.Count-1
    Set update = searchResult.Updates.Item(I)
    fileLog.writeline "[" & NOW &"] " &  I + 1 & "> " & update.Title
Next

If searchResult.Updates.Count = 0 Then
	fileLog.writeline "[" & NOW &"] " &  "There are no applicable updates."
	exec_reboot
	Wscript.Quit
End If

fileLog.writeline "[" & NOW &"] " &  "Creating collection of updates to download:"

Set updatesToDownload = CreateObject("Microsoft.Update.UpdateColl")

For I = 0 to searchResult.Updates.Count-1
    Set update = searchResult.Updates.Item(I)
    addThisUpdate = false
    If update.InstallationBehavior.CanRequestUserInput = true Then
fileLog.writeline "[" & NOW &"] " &  I + 1 & "> skipping: " & update.Title & _
        " because it requires user input"
    Else
        If update.EulaAccepted = false Then
fileLog.writeline "[" & NOW &"] " & I + 1 & "> note: " & update.Title & _
            " has a license agreement that must be accepted:"
fileLog.writeline "[" & NOW &"] " &  update.EulaText
fileLog.writeline "[" & NOW &"] " & "Do you accept this license agreement? (Y/N)"
            strInput = WScript.StdIn.Readline
            WScript.Echo 
            If (strInput = "Y" or strInput = "y") Then
                update.AcceptEula()
                addThisUpdate = true
            Else
                WScript.Echo I + 1 & "> skipping: " & update.Title & _
                " because the license agreement was declined"
            End If
        Else
            addThisUpdate = true
        End If
    End If
    If addThisUpdate = true Then
fileLog.writeline "[" & NOW &"] " &  I + 1 & "> adding: " & update.Title 
        updatesToDownload.Add(update)
    End If
Next

If updatesToDownload.Count = 0 Then
	fileLog.writeline "[" & NOW &"] " & "All applicable updates were skipped."
	exec_reboot
	Wscript.Quit

End If
    
fileLog.writeline "[" & NOW &"] " & "Downloading updates..."

Set downloader = updateSession.CreateUpdateDownloader() 
downloader.Updates = updatesToDownload
downloader.Download()

Set updatesToInstall = CreateObject("Microsoft.Update.UpdateColl")

rebootMayBeRequired = false

fileLog.writeline "[" & NOW &"] " &  "Successfully downloaded updates:" & searchResult.Updates.Count & " =================" 

For I = 0 To searchResult.Updates.Count-1
    set update = searchResult.Updates.Item(I)
    If update.IsDownloaded = true Then
		fileLog.writeline "[" & NOW &"] " &  I + 1 & "> " & update.Title 
        updatesToInstall.Add(update) 
        If update.InstallationBehavior.RebootBehavior > 0 Then
            rebootMayBeRequired = true
        End If
    End If
Next

If updatesToInstall.Count = 0 Then
	fileLog.writeline "[" & NOW &"] " &  "No updates were successfully downloaded."
'    WScript.Quit
End If

If rebootMayBeRequired = true Then
	fileLog.writeline "[" & NOW &"] " &  "These updates may require a reboot."
End If

	fileLog.writeline "[" & NOW &"] " &  "install updates ====================================================== "

	fileLog.writeline "[" & NOW &"] " & "Installing updates..."
    Set installer = updateSession.CreateUpdateInstaller()
    installer.Updates = updatesToInstall
    Set installationResult = installer.Install()
 
    'Output results of install
	fileLog.writeline "[" & NOW &"] " & "Installation Result: " & installationResult.ResultCode 
	fileLog.writeline "[" & NOW &"] " & "Reboot Required: " & installationResult.RebootRequired & vbCRLF 
	fileLog.writeline "[" & NOW &"] " & "Listing of updates installed " & "and individual installation results:" 
 
    For I = 0 to updatesToInstall.Count - 1
		fileLog.writeline "[" & NOW &"] " &  I + 1 & "> " & updatesToInstall.Item(i).Title & ": " & installationResult.GetUpdateResult(i).ResultCode   
    Next

exec_reboot

Function exec_reboot
	fileLog.writeline "[" & NOW &"]" & "Start Reboot"
	
	Dim ObjShell
	Set ObjShell = CreateObject("WScript.Shell")
	ObjShell.Run "%comspec% /c shutdown -c ""윈도우 업데이트를 윈해 서버가 재부팅됩니다."" -r -t 60 -f -d P:2:18", ,TRUE

	fileLog.writeline "[" & NOW &"]" & "Rebooting...."
End Function 