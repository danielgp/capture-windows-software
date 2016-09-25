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
If (objFSO.FileExists(strCurDir & "\" & strResultFileName & strResultFileType)) Then
    bolFileHeaderToAdd = False
Else
    bolFileHeaderToAdd = True
End If
Set ReportFile = objFSO.OpenTextFile(strCurDir & "\" & strResultFileName & strResultFileType, ForAppending, True) 
Set objNetwork = CreateObject("Wscript.Network") 
OsType = WshShell.RegRead("HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment\PROCESSOR_ARCHITECTURE")
wscript.echo "Please wait until script is completed..." 
Set SrvList = objFSO.OpenTextFile(strCurDir & "\WindowsComputerList.txt", ForReading) 
Do Until SrvList.AtEndOfStream 
    strComputer = LCase(SrvList.ReadLine) 
    If checkServerResponse(strComputer) then 
        srvIP = checkSoftware(strComputer, bolFileHeaderToAdd, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\")
        If OsType = "AMD64" Then
            srvIP = checkSoftware(strComputer, False, "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\")
        End If
    Else 
        ReportFile.write strComputer & "," & "Server is unreachable." 
        ReportFile.writeline 
    End If 
Loop 
Wscript.echo "Script completed, consult [" & strCurDir & "\" & strResultFileName & strResultFileType & "] for captured software list currently installed." & vbNewLine & vbNewLine & "Thank you for using this script!" 

Function Number2Digits(InputNo)
	If InputNo < 10 Then
		Number2Digits = "0" & InputNo
	Else
		Number2Digits = InputNo
	End If
End Function
Function CurrentDate2SqlFormat()
	CurrentDate2SqlFormat = DatePart("yyyy", Now()) & _
        "-" & Number2Digits(DatePart("m", Now())) & _
        "-" & Number2Digits(DatePart("d", Now()))
End Function
Function CurrentDateTime2SqlFormat()
	CurrentDateTime2SqlFormat = CurrentDate2SqlFormat() & _
        " " & Number2Digits(DatePart("h", Now())) & _
        ":" & Number2Digits(DatePart("n", Now())) & _
        ":" & Number2Digits(DatePart("s", Now()))
End Function
Function CleanStringStartEnd(strFullStringToClean, strStartCleanSubString, strEndCleanSubString)
    intStartCleanPosition = InStr(1, strFullStringToClean, strStartCleanSubString, vbTextCompare)
    intEndCleanPosition = InStr(1, strFullStringToClean, strEndCleanSubString, vbTextCompare)
    If ((intStartCleanPosition > 0) And (intEndCleanPosition > 0)) Then
        strCleanedString = Trim(Replace(Left(strFullStringToClean, (intStartCleanPosition - 1)) & " " & Right(strFullStringToClean, (Len(strFullStringToClean) - intEndCleanPosition)), "  ", " "))
        strCleanedString = CleanStringStartEnd(strCleanedString, strStartCleanSubString, strEndCleanSubString) ' to ensure if more then 1 occurence of identifiers is being taken care of
    Else
        strCleanedString = strFullStringToClean
    End If
    CleanStringStartEnd = strCleanedString
End Function
Function CleanStringWithBlacklistArray(strFullStringToClean, aryBlackList, strStringToReplaceWith)
    strCleanedString = strFullStringToClean
    For Each strBlackListPiece In aryBlackList 
        strCleanedString = Replace(Replace(strCleanedString, strBlackListPiece, strStringToReplaceWith), "  ", " ")
    Next
    CleanStringWithBlacklistArray = Trim(strCleanedString)
End Function
Function InArray(Haystack, GivenArray)
	Dim bReturn
	bReturn = False
	For Each elmnt In GivenArray
		If cStr(Haystack) = elmnt Then 
			bReturn = True
		End If
	Next
	InArray = bReturn
End Function
Function CleanStringOfNumericPiece(strFullStringToClean)
    ' break entire string into pieces with space as separator
    aryFullStringToClean = Split(strFullStringToClean, " ")
    strCleanedString = ""
    For Each strCurrentPiece In aryFullStringToClean 
        ' if strCurrentPiece is whitelisted as not to be removed
        If (InArray(strCurrentPiece, Array("360", "365"))) Then
            bolCurrentPieceToKeep = True
        Else
            intFirstCharacterCodeOfCurrentPiece = Asc(Left(strCurrentPiece, 1))
            intLastCharacterCodeOfCurrentPiece = Asc(Right(strCurrentPiece, 1))
            bolCurrentPieceToKeep = True
            ' test if 1st character is numeric
            If ((intFirstCharacterCodeOfCurrentPiece >= Asc("0")) And (intFirstCharacterCodeOfCurrentPiece <= Asc("9"))) Then
                ' test if last character is numeric
                If ((intLastCharacterCodeOfCurrentPiece >= Asc("0")) And (intLastCharacterCodeOfCurrentPiece <= Asc("9"))) Then
                    bolCurrentPieceToKeep = False
                End If
            End If
        End If
        If (bolCurrentPieceToKeep) Then
            strCleanedString = Trim(strCleanedString & " " & strCurrentPiece)
        End If
    Next
    CleanStringOfNumericPiece = strCleanedString
End Function
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
            strFieldSeparator  & "Software name cleaned" & _
            strFieldSeparator & "Install Location" & _
            strFieldSeparator & "Installation Date" & _
            strFieldSeparator & "Size [bytes]" & _
            strFieldSeparator & "Version (major.minor)" & _
            strFieldSeparator & "Full Version Cleaned" & _
            strFieldSeparator & "URL Info About" & _
            strFieldSeparator & "Registry Key Trunk" & _
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
            strSoftwareNameCleaned = CleanStringStartEnd(strDisplayName, " (", ")")
            aryBlackListToClean = Array("(x86_x64)", "_WHQL", "_X64", "_X86", "64-bit", "beta", "en-us", "for x64", "for x86", "SP1", "SP2", "SP3", "Update", "version", "VS2005", "x64", "x86")
            strSoftwareNameCleaned = CleanStringWithBlacklistArray(strSoftwareNameCleaned, aryBlackListToClean, "")
            aryBlackListToReplaceWithSpace = Array(" -")
            strSoftwareNameCleaned = CleanStringWithBlacklistArray(strSoftwareNameCleaned, aryBlackListToReplaceWithSpace, " ")
            aryBlackListToReplaceWithOriginal = Array("(R)")
            strSoftwareNameCleaned = CleanStringWithBlacklistArray(strSoftwareNameCleaned, aryBlackListToReplaceWithOriginal, Chr(174))
            strSoftwareNameCleaned = CleanStringOfNumericPiece(strSoftwareNameCleaned)
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
            strSubkeyPieces = Split(strKey, "\")
            ReportFile.writeline CurrentDateTime2SqlFormat() & _ 
                strFieldSeparator & strComputer & _
                strFieldSeparator & strPublisher & _
                strFieldSeparator & strDisplayName & _
                strFieldSeparator & strSoftwareNameCleaned & _
                strFieldSeparator & strInstallLocation  & _
                strFieldSeparator & strDateYMD  & _
                strFieldSeparator & strSizeInBytes & _
                strFieldSeparator & strDisplayVersion & _
                strFieldSeparator & strDisplayVersionCleaned & _
                strFieldSeparator & strURLInfoAbout & _
                strFieldSeparator & strSubkeyPieces(1) & _
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
