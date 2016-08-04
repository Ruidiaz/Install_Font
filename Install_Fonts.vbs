'------------------------------------------------------------------------------------
'	Install_Fonts.vbs
'------------------------------------------------------------------------------------

Const wbemFlagReturnImmediately = &h10
Const wbemFlagForwardOnly = &h20
Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8

On Error Resume Next
Dim objShell, objFSO
Dim WriteLogfile, strlogFile

strComputer = "."

Set objShell = CreateObject("Shell.Application")
Set wshShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.Filesystemobject")

' log file settings
WriteLogfile = 1	' change to '0' if no logging is required
LogfilePath = "C:\IT_Logs"	

'=== Logfile Path and Name ; create ===
strTimeStamp = fnGetTimeStamp()
strLogfileName = "Install_Fonts_" & strTimeStamp & ".log"	'amend logfile name
strLogfile = LogfilePath & "\" & strLogfileName

If WriteLogfile = 1 Then
	If Not objFSO.FolderExists(LogfilePath) Then
		Set createLogFolder = objFSO.CreateFolder(LogfilePath)
	End If
End If


' main part
writeLog("--------------------------------------------------------------------------------------------------")
writeLog("INFO: Install_Fonts started " & strTimeStamp & ".")
writeLog("--------------------------------------------------------------------------------------------------")

'strFontSourcePath = "\\ukcorpuktecimg1\sccm$\SCCM_2012\Apps\FONTS_Free"	'try code below instead
wshShell.CurrentDirectory = objFSO.GetParentFolderName(Wscript.ScriptFullName)
strFontSourcePath = wshShell.CurrentDirectory & "\" & "SOURCE"

If objFSO.FolderExists(strFontSourcePath) Then
	 
	Set objNameSpace = objShell.Namespace(strFontSourcePath)
	Set objFolder = objFSO.getFolder(strFontSourcePath)
	'Set objDir = "%windir%\fonts"
	
	For Each objFile In objFolder.files

		If LCase(right(objFile,4)) = ".ttf" OR LCase(right(objFile,4)) = ".otf" Then
		
				If objFSO.FileExists("C:\Windows\Fonts\" & objFile.Name) Then
					writeLog("Font already installed: " & objFile.Name)
				Else

			Set objFont = objNameSpace.ParseName(objFile.Name)
	
			  objFont.InvokeVerb("Install")
			  writeLog("Installed Font: " & objFile.Name)

			Set objFont = Nothing

		End If

  End If

 Next

Else

 writeLog("ERROR: Font Source Path does not exists")

End If


writeLog("--------------------------------------------------------------------------------------------------")
writeLog("INFO: Install_Fonts has finished")
writeLog("--------------------------------------------------------------------------------------------------")


' functions and subs
'=============================================
'	Function for logging
'=============================================
Function writeLog(strLogging)
	If WriteLogfile = 1 Then
		Set objLog = objFSO.OpenTextFile(strLogfile, ForAppending, True)
		objlog.writeline Date & " " & Time & " " & " <" & strLogging & ">"
		objLog.close
	End If
End Function

'=============================================
'	Functions for formatting MsgBoxes
'=============================================
Function MsgBoxInfo(strTXT)
	MsgBox strTXT, vbInformation, "INFO"
End Function

Function MsgBoxExcl(strTXT)
	MsgBox strTXT, vbExclamation, "ATTENTION"
End Function

Function MsgBoxErr(strTXT)
	MsgBox strTXT, vbCritical, "ERROR"
End Function

'=============================================
'	Function for formatting TimeStamps
'=============================================
Function fnGetTimeStamp()
	t = Now
	
	y = Year(t)
	m = Month(t)
	d = Day(t)
	h = Hour(t)
	mn = Minute(t)
	s = Second(t) 
	
	strArr = Array(m,d,h,mn,s)
	
	For i=0 To 4
		If strArr(i) < 10 Then
			strArr(i) = "0" & strArr(i)
		End If
	Next
	
	fnGetTimeStamp = y & strArr(0) & strArr(1) & "_" & strArr(2) & strArr(3) & strArr(4)
End Function
