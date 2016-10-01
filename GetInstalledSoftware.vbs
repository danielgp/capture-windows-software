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
Dim strResultFileType
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set WshShell = WScript.CreateObject("WScript.Shell") 
'-----------------------------------------------------------------------------------------------------------------------
MsgBox "This script will read from Windows Registry the entire list of installed software and export it in a file with a pre-configured name!" & vbNewLine & vbNewLine & "please wait until script is completed...", vbOKOnly + vbInformation, "Start feedback" 
InputResultType = MsgBox("This is a script intended to read from Windows Registry all your installed software applications under current Windows installation!" & vbNewLine & vbNewLine & "Do you want to store obtained results into CSV format file?" & vbNewLine & vbNewLine & "if you choose No a SQL file will be used instead" & vbNewLine & "otherwise choosing Cancel will end current script without any processing and result.", vbYesNoCancel + vbQuestion, "Choose processing result type")
If (InputResultType = vbCancel) Then
    MsgBox "This is a script intended to read from Windows Registry all your installed software applications under current Windows installation!" & vbNewLine & vbNewLine & "You have chosen to terminate execution without any processing and no result, should you arrive at this point by mistake just re-execute it and pay greater attention to previous options dialogue, otherwise thanks for your attention!", vbOKOnly + vbExclamation, "Script end"
Else
    StartTime = Timer()
    Select Case InputResultType
        Case vbYes
            strResultFileType = ".csv"
        Case vbNo 
            strResultFileType = ".sql"
    End Select
    OsType = WshShell.RegRead("HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment\PROCESSOR_ARCHITECTURE")
    strCurDir = WshShell.CurrentDirectory
    Set SrvListFile = objFSO.OpenTextFile(strCurDir & "\WindowsComputerList.txt", ForReading) 
    If (objFSO.FileExists(strCurDir & "\" & strResultFileName & strResultFileType)) Then
        bolFileHeaderToAdd = False
    Else
        bolFileHeaderToAdd = True
    End If
    Set ReportFile = objFSO.OpenTextFile(strCurDir & "\" & strResultFileName & strResultFileType, ForAppending, True) 
    Do Until SrvListFile.AtEndOfStream 
        strComputer = LCase(SrvListFile.ReadLine) 
        checkSoftware(strComputer, bolFileHeaderToAdd, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\")
        If (OsType = "AMD64") Then
            checkSoftware(strComputer, False, "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\")
        End If
    Loop 
    SrvListFile.Close
    ReportFile.Close
    EndTime = Timer()
    MsgBox "This script has completed processing from Windows Registry entire list of installed software under current Windows installation (in just " & FormatNumber(EndTime - StartTime, 0) & " seconds), please consult generated file [" & strCurDir & "\" & strResultFileName & strResultFileType & "]." & vbNewLine & vbNewLine & "Thank you for using this script, hope to see you back soon!", vbOKOnly + vbInformation, "Script end"
End If
'-----------------------------------------------------------------------------------------------------------------------
Function Number2Digits(InputNo)
    If (InputNo < 10) Then
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
        If (cStr(Haystack) = elmnt) Then 
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
        ' if strCurrentPiece is amoung whitelisted values, does not have to be removed
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
Function CleanStringBeforeOrAfterNumber(strFullStringToClean, strBeforeOrAfter, aryBlackList, strStringToReplaceWith)
    ' break entire string into pieces with space as separator
    aryFullStringToClean = Split(strFullStringToClean, " ")
    strCleanedString = ""
    intPieceCounter = 0
    intLastPieceNumber = UBound(aryFullStringToClean)
    For Each strBlackListPiece In aryBlackList 
        For Each strCurrentPiece In aryFullStringToClean 
            ' first or last piece does not need any cleaning as cannot be followed by a numbers or anything else
            If ((intPieceCounter = 0) Or (intPieceCounter = intLastPieceNumber)) Then
                bolCurrentPieceToKeep = True
            Else
                bolCurrentPieceToKeep = True
                If (strCurrentPiece = strBlackListPiece) Then
                    Select Case strBeforeOrAfter
                        Case "After"
                            If (IsNumeric(aryFullStringToClean((intPieceCounter - 1)))) Then
                                bolCurrentPieceToKeep = False
                            End If
                        Case "Before"
                            If (IsNumeric(aryFullStringToClean((intPieceCounter + 1)))) Then
                                bolCurrentPieceToKeep = False
                            End If
                    End Select
                End If
            End If
            If (bolCurrentPieceToKeep) Then
                strCleanedString = Trim(strCleanedString & " " & strCurrentPiece)
            End If
            intPieceCounter = intPieceCounter + 1
        Next
    Next
    CleanStringBeforeOrAfterNumber = strCleanedString
End Function
Function checkSoftware(strComputer, bolWriteHeader, strKey) 
    Dim aryJSONinformationCSV(5)
    Dim aryJSONinformationSQL(5)
    Const HKLM = &H80000002 'HKEY_LOCAL_MACHINE
    aryInformationToExpose = Array(_
        "Evaluation Timestamp", _
        "Host Name", _
        "Publisher Name", _
        "Software Name", _
        "Install Location", _
        "Installation Date", _
        "Size [bytes]", _
        "Full Version", _
        "Other Info", _
        "Registry Key Trunk", _
        "Registry SubKey" _
    )
    aryInformationToExposeOtherInfo = Array(_
        "Publisher Name", _
        "Software", _
        "Version Major", _
        "Version Minor", _
        "Version Displayed", _
        "URL Info About" _
    )
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
    If (LCase(strResultFileType) = ".csv") Then
        If (bolWriteHeader) Then
            ReportFile.writeline Join(aryInformationToExpose, strFieldSeparator)
        End If
    End If
    objReg.EnumKey HKLM, strKey, arrSubkeys 
    For Each strSubkey In arrSubkeys 
        intReturnN = objReg.GetStringValue(HKLM, strKey & strSubkey, strEntryDisplayName, strDisplayName) 
        If (intReturnN <> 0) Then
            objReg.GetStringValue HKLM, strKey & strSubkey, strEntryQuietDisplayName, strDisplayName
        End If
        If (strDisplayName <> "") Then 
            intReturnP = objReg.GetStringValue(HKLM, strKey & strSubkey, strEntryPublisher, strPublisher) 
            intReturnL = objReg.GetStringValue(HKLM, strKey & strSubkey, strEntryInstallLocation, strInstallLocation) 
            intReturnD = objReg.GetStringValue(HKLM, strKey & strSubkey, strEntryInstallDate, strInstallDate) 
            objReg.GetDWORDValue HKLM, strKey & strSubkey, strEntryVersionMajor, intValueVersionMajor 
            objReg.GetDWORDValue HKLM, strKey & strSubkey, strEntryVersionMinor, intValueVersionMinor 
            objReg.GetDWORDValue HKLM, strKey & strSubkey, strEntryEstimatedSize, intEstimatedSize 
            intReturnV = objReg.GetStringValue(HKLM, strKey & strSubkey, strEntryDisplayVersion, strDisplayVersion) 
            intReturnU = objReg.GetStringValue(HKLM, strKey & strSubkey, strEntryURLInfoAbout, strURLInfoAbout) 
            If (intReturnP = 0) Then
                strPublisherName = PublishersHarmonized(strPublisher)
            Else
                strPublisherName = "NULL"
            End If
            strSoftwareNameCleaned = CleanStringStartEnd(strDisplayName, " (", ")")
            aryBlackListToRemoveBetweenNumbers = Array("Update") ' to properly clean "Java <No> Update <No>" software name
            strSoftwareNameCleaned = CleanStringBeforeOrAfterNumber(strSoftwareNameCleaned, "Before", aryBlackListToRemoveBetweenNumbers, "")
            aryBlackListToRemoveBetweenNumbers = Array("R2") ' to properly clean "Microsoft SQL Server <No> R2 Native Client" software name
            strSoftwareNameCleaned = CleanStringBeforeOrAfterNumber(strSoftwareNameCleaned, "After", aryBlackListToRemoveBetweenNumbers, "")
            aryBlackListToClean = Array("(x86_x64)", "(x64)", "(x86)", "_WHQL", "_X64", "_X86", "64-bit", "beta", "en-us", "for x64", "for x86", "SP1", "SP2", "SP3", "version", "VS2005", "x64", "x86")
            strSoftwareNameCleaned = CleanStringWithBlacklistArray(strSoftwareNameCleaned, aryBlackListToClean, "")
            aryBlackListToReplaceWithSpace = Array(" -")
            strSoftwareNameCleaned = CleanStringWithBlacklistArray(strSoftwareNameCleaned, aryBlackListToReplaceWithSpace, " ")
            aryBlackListToReplaceWithOriginal = Array("(R)")
            strSoftwareNameCleaned = CleanStringWithBlacklistArray(strSoftwareNameCleaned, aryBlackListToReplaceWithOriginal, Chr(174))
            strSoftwareNameCleaned = CleanStringOfNumericPiece(strSoftwareNameCleaned)
            If ((intReturnL <> 0) Or (Len(Trim(strInstallLocation)) = 0)) Then
                strInstallLocation = "NULL"
            End If
            If (intReturnD <> 0) Then
                strDateYMD = "NULL"
            Else
                If (IsNumeric(strInstallDate)) Then
                    If (strInstallDate > 0) Then
                        strDateYMD = Mid(strInstallDate, 1, 4) & _
                            "-" & Mid(strInstallDate, 5, 2) & _
                            "-" & Mid(strInstallDate, 7, 2)
                    Else
                        strDateYMD = "NULL"
                    End If
                Else
                    strDateYMD = "NULL"
                End If
            End If
            If (intEstimatedSize <> "") Then
                strSizeInBytes = CStr(Replace(intEstimatedSize, ",", "."))
            Else
                strSizeInBytes = "NULL"
            End If
            If (intValueVersionMajor > 0) Then
                intValueVersionMajor = CStr(intValueVersionMajor)
            Else
                intValueVersionMajor = "-"
            End If
            If (intValueVersionMinor > 0) Then
                intValueVersionMinor = CStr(intValueVersionMinor)
            Else
                intValueVersionMinor = "-"
            End If
            If (intReturnV <> 0) Then
                strDisplayVersion = "-"
                strDisplayVersionCleaned = "NULL"
            Else
                ' In some cases DisplayVersion has a date before the version so we're going to take in consideration only the very last group of continuous string splitted by space
                strDisplayVersionPieces = Split(CStr(Replace(Replace(strDisplayVersion, "a", "."), " beta ", ".")), " ")
                For Each strDisplayVersionPieceValue In strDisplayVersionPieces
                    If (IsNumeric(strDisplayVersionPieceValue)) Then
                        strDisplayVersionCleaned = strVersionPrefix & strDisplayVersionPieceValue
                    End If
                Next
            End If
            If ((intReturnU <> 0) Or (IsNull(strURLInfoAbout)) Or (Len(Trim(strURLInfoAbout)) = 0)) Then
                strURLInfoAbout = "-"
            Else
                If ((Left(strURLInfoAbout, 4) <> "http") And (Left(strURLInfoAbout, 4) = "www.")) Then
                    strURLInfoAbout = "http://" & strURLInfoAbout
                End If
                If (Right(strURLInfoAbout, 1) = "/") Then
                    strURLInfoAbout = Left(strURLInfoAbout, (Len(strURLInfoAbout) - 1))
                End If
            End If
            strSubkeyPieces = Split(strKey, "\")
            aryValuesToExposeOtherInfo = Array(_
                strPublisher, _
                strDisplayName, _
                intValueVersionMajor, _
                intValueVersionMinor, _
                strDisplayVersion, _
                strURLInfoAbout _
            )
            Select Case LCase(strResultFileType)
                Case ".csv"
                    intCounter = 0
                    For Each crtOtherInfo in aryInformationToExposeOtherInfo
                        aryJSONinformationCSV(intCounter) = crtOtherInfo & ": " & _
                            aryValuesToExposeOtherInfo(intCounter)
                        intCounter = intCounter + 1
                    Next
                    strOtherInfo = Join(aryJSONinformationCSV, " | ")
                Case ".sql"
                    intCounter = 0
                    For Each crtOtherInfo in aryInformationToExposeOtherInfo
                        aryJSONinformationSQL(intCounter) = """" & crtOtherInfo & """: " & _
                            """" & aryValuesToExposeOtherInfo(intCounter) & """"
                        intCounter = intCounter + 1
                    Next
                    strOtherInfo = "{ " & Join(aryJSONinformationSQL, ", ") & " }"
            End Select
            aryValuesToExpose = Array(_
                CurrentDateTime2SqlFormat(), _ 
                strComputer, _
                strPublisherName, _
                strSoftwareNameCleaned, _
                strInstallLocation , _
                strDateYMD , _
                strSizeInBytes, _
                strDisplayVersionCleaned, _
                strOtherInfo, _
                strSubkeyPieces(1), _
                strSubkey _
            )
            Select Case LCase(strResultFileType)
                Case ".csv"
                    ReportFile.WriteLine Join(aryValuesToExpose, strFieldSeparator)
                Case ".sql"
                    ReportFile.WriteLine "INSERT INTO `in_windows_software_list` (" & _
                        BuildInsertOrUpdateSQLstructure(aryInformationToExpose, aryValuesToExpose, "InsertFields") & _
                        ") VALUES(" & _
                        BuildInsertOrUpdateSQLstructure(aryInformationToExpose, aryValuesToExpose, "InsertValues") & _
                        ") ON DUPLICATE KEY UPDATE " & _
                        BuildInsertOrUpdateSQLstructure(aryInformationToExpose, aryValuesToExpose, "Update") & _
                        ";"
            End Select
        End If 
    Next 
    If (LCase(strResultFileType) = ".sql") Then
        ReportFile.WriteLine "INSERT `version_details` (`FullVersion`) SELECT `FullVersion` FROM `in_windows_software_list` WHERE (`HostName` = '" & strComputer & "') AND (`FullVersion` IS NOT NULL) AND (`FullVersion` NOT IN (SELECT `FullVersion` FROM `version_details` GROUP BY `FullVersion`)) GROUP BY `FullVersion`;"
    End If
End Function 
Function CSVfieldNamesIntoSQLfieldName(aryFieldNames)
    strFieldListForMySQLinsert = Join(aryFieldNames, "|")
    strFieldListForMySQLinsert = CleanStringWithBlacklistArray(strFieldListForMySQLinsert, Array("[bytes]"), "Bytes")
    strFieldListForMySQLinsert = CleanStringWithBlacklistArray(strFieldListForMySQLinsert, Array("(major.minor)"), "Major Minor")
    strFieldListForMySQLinsert = CleanStringWithBlacklistArray(strFieldListForMySQLinsert, Array(" "), "")
    CSVfieldNamesIntoSQLfieldName = strFieldListForMySQLinsert
End Function
Function BuildInsertOrUpdateSQLstructure(aryFieldNames, aryFieldValues, strInsertOrUpdate)
    Counter = 0
    strUpdateSQLstructure = ""
    Select Case strInsertOrUpdate
        Case "InsertFields"
            aryFieldValuesMySQL = Split(CSVfieldNamesIntoSQLfieldName(aryFieldNames), "|")
            strUpdateSQLstructure = "`" & Join(aryFieldValuesMySQL, "`, `") & "`"
        Case "InsertValues"
            strUpdateSQLstructure = "'" & Join(aryFieldValues, "', '") & "'"
        Case "Update"
            aryFieldValuesMySQL = Split(CSVfieldNamesIntoSQLfieldName(aryFieldNames), "|")
            intFieldNumbered = UBound(aryFieldValuesMySQL)
            For Each strFieldName In aryFieldValuesMySQL 
                If (Counter <= (intFieldNumbered - 2)) Then 'last 2 fields are part of PK so will not be used
                    If (Counter > 0) Then
                        strUpdateSQLstructure = strUpdateSQLstructure & ", "
                    End If
                    strUpdateSQLstructure = strUpdateSQLstructure & "`" & strFieldName & "` = '" & aryFieldValues(Counter) & "'"
                End If
                Counter = Counter + 1
            Next
    End Select
    BuildInsertOrUpdateSQLstructure = Replace(Replace(strUpdateSQLstructure, "'NULL'", "NULL"), "\", "\\")
End Function 
Function PublishersHarmonized(strPublisherName)
    aryPublishersTemplate = Array(_
        Array("Dell", "Dell Inc."), _
        Array("Dell Packaging Team", "Dell Inc."), _
        Array("Dell Products, LP", "Dell Inc."), _
        Array("Dell SecureWorks", "Dell Inc."), _
        Array("Google", "Google Inc."), _
        Array("Informatica", "Informatica Corporation"), _
        Array("Informatica Co.", "Informatica Corporation"), _
        Array("Intel", "Intel Corporation"), _
        Array("Intel(R) Corporation", "Intel Corporation"), _
        Array("Lumension", "Lumension Security, Inc."), _
        Array("McAfee", "McAfee, Inc."), _
        Array("Microsoft", "Microsoft Corporation"), _
        Array("Oracle", "Oracle Corporation"), _
        Array("Qualcomm Atheros", "Qualcomm Atheros Communications"), _
        Array("SAP", "SAP AG"), _
        Array("SAP SE", "SAP AG"), _
        Array("Symantec Corp.", "Symantec Corporation") _
    )
    strPublishersHarmonized = ""
    For Each strCurrentPublisherHarmonized In aryPublishersTemplate 
        If (strPublisherName = strCurrentPublisherHarmonized(0)) Then
            strPublishersHarmonized = strCurrentPublisherHarmonized(1)
        End If
    Next
    If (strPublishersHarmonized = "") Then
        PublishersHarmonized = strPublisherName
    Else
        PublishersHarmonized = strPublishersHarmonized
    End If
End Function
'-----------------------------------------------------------------------------------------------------------------------
