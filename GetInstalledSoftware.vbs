'-----------------------------------------------------------------------------------------------------------------------
' IT License
' 
' Copyright (c) 2016 Daniel Popiniuc
' 
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
' 
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
' 
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.
'-----------------------------------------------------------------------------------------------------------------------
Const ForReading = 1 
Const ForAppending = 8 
Const strFieldSeparator = ";" 
Const strVersionPrefix = "v" 
Const strResultFileName = "WindowsSoftwareList" 
Const strResultFileType = ".csv" 
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set WshShell = WScript.CreateObject("WScript.Shell") 
strCurDir = WshShell.CurrentDirectory
Set SrvList = objFSO.OpenTextFile(strCurDir & "\WindowsComputerList.txt", ForReading) 
Set ReportFile = objFSO.OpenTextFile(strCurDir & "\" & strResultFileName & strResultFileType, ForAppending, True) 
Set objNetwork = CreateObject("Wscript.Network") 
OsType = WshShell.RegRead("HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment\PROCESSOR_ARCHITECTURE")
wscript.echo "Please wait until script is completed..." 
Do Until SrvList.AtEndOfStream 
	strComputer = LCase(SrvList.ReadLine) 
	If checkServerResponse(strComputer) then 
		srvIP = checkSoftware(strComputer, True, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\")
		If OsType = "AMD64" Then
			srvIP = checkSoftware(strComputer, False, "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\")
		End If
	Else 
		ReportFile.write strComputer & "," & "Server is unreachable." 
		ReportFile.writeline 
	End If 
Loop 
Wscript.echo "Script completed, consult [" & strCurDir & "\" & strResultFileName & strResultFileType & "] for captured software list currently installed." & vbNewLine & vbNewLine & "Thank you for using this script!" 

Function checkSoftware(strComputer, bolWriteHeader, strKey) 
    Const HKLM = &H80000002 'HKEY_LOCAL_MACHINE 
    strEntryDisplayName = "DisplayName" 
    strEntryQuietDisplayName = "QuietDisplayName" 
    strEntryPublisher = "Publisher" 
    strEntryInstallLocation = "InstallLocation" 
    strEntryInstallDate = "InstallDate" 
    strEntryVersionMajor = "VersionMajor" 
    strEntryVersionMinor = "VersionMinor" 
    strEntryEstimatedSize = "EstimatedSize" 
    strEntryDisplayVersion = "DisplayVersion" 
    strEntryURLInfoAbout = "URLInfoAbout" 
    Set objReg = GetObject("winmgmts://" & strComputer & "/root/default:StdRegProv") 
	If bolWriteHeader Then
		ReportFile.writeline "Evaluation timestamp" & _ 
			strFieldSeparator  & "HostName" & _
			strFieldSeparator  & "Publisher" & _
			strFieldSeparator  & "Software" & _
			strFieldSeparator & "Install Location" & _
			strFieldSeparator & "Installation Date" & _
			strFieldSeparator & "Size [bytes]" & _
			strFieldSeparator & "Version (major.minor)" & _
			strFieldSeparator & "Full Version Cleaned" & _
			strFieldSeparator & "URL Info About" & _
			strFieldSeparator & "Registry SubKey"
	End If
    objReg.EnumKey HKLM, strKey, arrSubkeys 
    For Each strSubkey In arrSubkeys 
        intReturnN = objReg.GetStringValue(HKLM, strKey & strSubkey, strEntryDisplayName, strDisplayName) 
		If intReturnN <> 0 Then
			objReg.GetStringValue HKLM, strKey & strSubkey, strEntryQuietDisplayName, strDisplayName
		End If
        If strDisplayName <> "" Then 
            intReturnP = objReg.GetStringValue(HKLM, strKey & strSubkey, strEntryPublisher, strPublisher) 
            intReturnL = objReg.GetStringValue(HKLM, strKey & strSubkey, strEntryInstallLocation, strInstallLocation) 
            objReg.GetStringValue HKLM, strKey & strSubkey, strEntryInstallDate, strInstallDate 
            objReg.GetDWORDValue HKLM, strKey & strSubkey, strEntryVersionMajor, intValueVersionMajor 
            objReg.GetDWORDValue HKLM, strKey & strSubkey, strEntryVersionMinor, intValueVersionMinor 
            objReg.GetDWORDValue HKLM, strKey & strSubkey, strEntryEstimatedSize, intEstimatedSize 
			intReturnV = objReg.GetStringValue(HKLM, strKey & strSubkey, strEntryDisplayVersion, strDisplayVersion) 
            intReturnU = objReg.GetStringValue(HKLM, strKey & strSubkey, strEntryURLInfoAbout, strURLInfoAbout) 
			If intReturnP <> 0 Then
				strPublisher = "_unknown publisher_"
			End If
			If (intReturnL <> 0) Or (Len(Trim(strInstallLocation)) = 0) Then
				strInstallLocation = "_unknown install location_"
			End If
            If strInstallDate > 0 Then 
				strDateYMD = Mid(strInstallDate, 1, 4) & _
					"-" & Mid(strInstallDate, 5, 2) & _
					"-" & Mid(strInstallDate, 7, 2)
			Else
				strDateYMD = "NULL"
            End If
			If intEstimatedSize <> "" Then
				strSizeInBytes = CStr(Replace(intEstimatedSize, ",", "."))
			Else
				strSizeInBytes = "0"
			End If
			If intValueVersionMajor >= 0 Then  
				If intValueVersionMinor >= 0 Then 
					strVersionMajorMinor = CStr(intValueVersionMajor) & "." & CStr(intValueVersionMinor)
				Else
					strVersionMajorMinor = CStr(intValueVersionMajor) & ".0"
				End If
			Else
				strVersionMajorMinor = "0.0"
			End If
			If intReturnV <> 0 Then
				strDisplayVersion = strVersionPrefix & "0.0.0"
				strDisplayVersionCleaned = strVersionPrefix & "0.0.0"
			Else
				strDisplayVersion = Replace(strDisplayVersion, " beta ", ".")
				' In some cases DisplayVersion has a date before the version so we're going to take in consideration only the very last group of continuous string splitted by space
				strDisplayVersionPieces = Split(CStr(strDisplayVersion), " ")
				For Each strDisplayVersionPieceValue In strDisplayVersionPieces 
					strDisplayVersionCleaned = strVersionPrefix & strDisplayVersionPieceValue
				Next
				strDisplayVersion = strVersionPrefix & strVersionMajorMinor
			End If
			If (intReturnU <> 0) Or (Len(Trim(strURLInfoAbout)) = 0) Then
				strURLInfoAbout = "NULL"
			Else
				If ((Left(strURLInfoAbout, 4) <> "http") And (Left(strURLInfoAbout, 4) = "www.")) Then
					strURLInfoAbout = "http://" & strURLInfoAbout
				End If
				If (Right(strURLInfoAbout, 1) = "/") Then
					strURLInfoAbout = Left(strURLInfoAbout, (Len(strURLInfoAbout) - 1))
				End If
			End If
			ReportFile.writeline Now() & _ 
				strFieldSeparator & strComputer & _
				strFieldSeparator & strPublisher & _
				strFieldSeparator & strDisplayName & _
				strFieldSeparator & strInstallLocation  & _
				strFieldSeparator & strDateYMD  & _
				strFieldSeparator & strSizeInBytes & _
				strFieldSeparator & strDisplayVersion & _
				strFieldSeparator & strDisplayVersionCleaned & _
				strFieldSeparator & strURLInfoAbout & _
				strFieldSeparator & strSubkey
        End If 
    Next 
End Function 
Function checkServerResponse(serverName) 
	strTarget = serverName 
	Set objShell = CreateObject("WScript.Shell") 
	Set objExec = objShell.Exec("ping -n 1 -w 1000 " & strTarget) 
	strPingResults = LCase(objExec.StdOut.ReadAll) 
	If InStr(strPingResults, "reply from") Then 
		checkServerResponse = True 
	Else 
		checkServerResponse = False 
	End If 
End Function 
