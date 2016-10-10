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
Const ForAppending = 8
Const ForReading = 1
Const HKLM = &H80000002 'HKEY_LOCAL_MACHINE
Const strFieldSeparator = ";"
Const strVersionPrefix = "v"
Const strOutputResultType = "SQL" ' another supported options is CSV
Const strConfigurationWindowsComputerList = "ConfigurationWindowsComputerList.txt"
Const strResultFileNameSoftware = "ResultWindowsSoftwareInstalled"
Const strResultFileNameDeviceDetails = "ResultWindowsDeviceDetails"
Const strResultFileNameDeviceVolumes = "ResultWindowsDeviceVolumes"
Const strConfigurationPortableSoftware = "ConfigurationPortableSoftwareList.txt"
Const strResultFileNamePortableSoftware = "ResultWindowsSoftwarePortable"
Const strConfigurationSecurityRiskComponents = "ConfigurationSecurityRiskComponents.txt"
Const strResultFileNameSecurityRiskComponents = "ResultSecurityRiskComponents"
Dim strResultFileType
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set WshShell = WScript.CreateObject("WScript.Shell")
Set dtmConvertedDate = CreateObject("WbemScripting.SWbemDateTime")
'-----------------------------------------------------------------------------------------------------------------------
Const strScriptIntroduction = "This script will read from Windows Management Instrumentation (WMI) current Device Details and from Windows Registry the entire list of installed/portable software applications and export it in a file with a pre-configured name!"
strInput = InputBox(strScriptIntroduction & vbNewLine & vbNewLine & _
    "Input one or multiple choices from list below (no separators required)" & vbNewLine & _
	"  a = Device Details" & vbNewLine & _
	"  b = Disk Volumes" & vbNewLine & _
	"  c = Installed Software" & vbNewLine & _
	"  d = Portable Software" & vbNewLine & _
	"-------------------------------------------------" & vbNewLine & _
	"  x = a through c" & vbNewLine & _
	"  y = a through d" & vbNewLine & _
	"-------------------------------------------------" & vbNewLine & _
	"  h = High Security Components Scanning" & vbNewLine & _
	"-------------------------------------------------" & vbNewLine & _
	"  z = all choices (a+b+c+d+h) in same order" _
	, "Capture Windows Software - start")
If ((InStr(1, strInput, "a", vbTextCompare) = 0) And (InStr(1, strInput, "b", vbTextCompare) = 0) And (InStr(1, strInput, "c", vbTextCompare) = 0) And (InStr(1, strInput, "d", vbTextCompare) = 0) And (InStr(1, strInput, "h", vbTextCompare) = 0) And (InStr(1, strInput, "x", vbTextCompare) = 0) And (InStr(1, strInput, "y", vbTextCompare) = 0) And (InStr(1, strInput, "z", vbTextCompare) = 0)) Then
    MsgBox Replace(strScriptIntroduction, " will ", " was intended to ") & vbNewLine & vbNewLine & "You have chosen to terminate execution without any processing and no result, should you arrive at this point by mistake just re-execute it and pay greater attention to previous options dialogue, otherwise thanks for your attention!", vbOKOnly + vbExclamation, "Capture Windows Software - cancelled"
Else
    StartTime = Timer()
    strCurDir = WshShell.CurrentDirectory
    Set SrvListFile = objFSO.OpenTextFile(strCurDir & "\" & strConfigurationWindowsComputerList, ForReading)
    Do Until SrvListFile.AtEndOfStream
        strComputer = CurrentComputerName(LCase(SrvListFile.ReadLine))
        Select Case UCase(strOutputResultType)
            Case "CSV"
                strResultFileType = ".csv"
                If (objFSO.FileExists(strCurDir & "\" & strResultFileNameDeviceDetails & strResultFileType)) Then
                    bolFileDeviceHeaderToAdd = False
                Else
                    bolFileDeviceHeaderToAdd = True
                End If
                If (objFSO.FileExists(strCurDir & "\" & strResultFileNameSoftware & strResultFileType)) Then
                    bolFileSoftwareHeaderToAdd = False
                Else
                    bolFileSoftwareHeaderToAdd = True
                End If
            Case "SQL"
                strResultFileType = ".sql"
        End Select
        Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
        strFilesResulted = ""
        If ((InStr(1, strInput, "a", vbTextCompare) > 0) Or (InStr(1, strInput, "x", vbTextCompare) > 0) Or (InStr(1, strInput, "y", vbTextCompare) > 0) Or (InStr(1, strInput, "z", vbTextCompare) > 0)) Then
            ReadWMI_All objWMIService, strComputer, strResultFileType, strFieldSeparator, objFSO, strCurDir, strResultFileNameDeviceDetails, ForAppending, bolFileDeviceHeaderToAdd
            strFilesResulted = strFilesResulted & "  - " & strCurDir & "\" & strResultFileNameDeviceDetails & strResultFileType & vbNewLine
        End If
        If ((InStr(1, strInput, "b", vbTextCompare) > 0) Or (InStr(1, strInput, "x", vbTextCompare) > 0) Or (InStr(1, strInput, "y", vbTextCompare) > 0) Or (InStr(1, strInput, "z", vbTextCompare) > 0)) Then
            ReadWMI_DeviceVolumes objWMIService, strComputer, strResultFileType, strFieldSeparator, objFSO, strCurDir, strResultFileNameDeviceVolumes, ForAppending
            strFilesResulted = strFilesResulted & "  - " & strCurDir & "\" & strResultFileNameDeviceVolumes & strResultFileType & vbNewLine
        End If
        If ((InStr(1, strInput, "c", vbTextCompare) > 0) Or (InStr(1, strInput, "x", vbTextCompare) > 0) Or (InStr(1, strInput, "y", vbTextCompare) > 0) Or (InStr(1, strInput, "z", vbTextCompare) > 0)) Then
            ReadRegistry_SofwareInstalled strComputer, strResultFileType, strFieldSeparator, objFSO, strCurDir, strResultFileNameSoftware, ForAppending
            strFilesResulted = strFilesResulted & "  - " & strCurDir & "\" & strResultFileNameSoftware & strResultFileType & vbNewLine
        End If
        If ((InStr(1, strInput, "d", vbTextCompare) > 0) Or (InStr(1, strInput, "y", vbTextCompare) > 0) Or (InStr(1, strInput, "z", vbTextCompare) > 0)) Then
            ReadLogicalDisk_PortableSoftware objFSO, objWMIService, strCurDir, ForReading, ForAppending, strResultFileNamePortableSoftware, strResultFileType, strFieldSeparator, strConfigurationPortableSoftware, "in_windows_software_portable"
            strFilesResulted = strFilesResulted & "  - " & strCurDir & "\" & strResultFileNamePortableSoftware & strResultFileType & vbNewLine
        End If
        If ((InStr(1, strInput, "h", vbTextCompare) > 0) Or (InStr(1, strInput, "z", vbTextCompare) > 0)) Then
            ReadLogicalDisk_PortableSoftware objFSO, objWMIService, strCurDir, ForReading, ForAppending, strResultFileNameSecurityRiskComponents, strResultFileType, strFieldSeparator, strConfigurationSecurityRiskComponents, "in_windows_security_risk_components"
            strFilesResulted = strFilesResulted & "  - " & strCurDir & "\" & strResultFileNameSecurityRiskComponents & strResultFileType & vbNewLine
        End If
    Loop
    SrvListFile.Close
    EndTime = Timer()
    MsgBox Replace(strScriptIntroduction, " will ", " has ") & _
        vbNewLine & vbNewLine & _
        "Entire evaluation took " & FormatNumber(EndTime - StartTime, 0) & " seconds." & _
        vbNewLine & vbNewLine & _
        "Consult results stored within following file(s):" & vbNewLine & _
        strFilesResulted & vbNewLine & _
        "Thank you for using this script and hope to see you back soon!", _
        vbOKOnly + vbInformation, "Capture Windows Software - finish"
End If
'-----------------------------------------------------------------------------------------------------------------------
Function AdjustEmptyValueWithinArrayAndGlueIt(aryEntryArray, strValueToReplace, strGlue)
    strFinal = strGlue
    Counter = 0
    For Each crtValue in aryEntryArray
        strFinal = strFinal & strGlue
        If ((crtValue = "") Or (IsNull(crtValue))) Then
            strFinal = strFinal & strValueToReplace
        Else
            strFinal = strFinal & Trim(Replace(crtValue, "||", "|"))
        End If
        Counter = Counter + 1
    Next
    AdjustEmptyValueWithinArrayAndGlueIt = Replace(strFinal, strGlue & strGlue, "")
End Function
Function ApplySoftwareNormalizationForLogicalDisks(ReportFile)
    ReportFile.WriteLine "ALTER TABLE `device_details` AUTO_INCREMENT = 1; /* Ensure no end gaps are present in the auto incrementing sequence for Device */"
    ReportFile.WriteLine "INSERT INTO `device_details` " & _
        "(`DeviceParrentName`, `DeviceName`, `DeviceOSdetails`, `DeviceHardwareDetails`) " & _
        "SELECT (CASE WHEN (JSON_EXTRACT(`dv`.`DetailedInformation`, '$.""Drive Type Code""') = 3) THEN REPLACE(JSON_EXTRACT(`dv`.`DetailedInformation`, '$.""System Name""'), '""', '') ELSE NULL END), `dv`.`VolumeSerialNumber`, NULL, `dv`.`DetailedInformation` " & _
        "FROM `device_volumes` `dv` ON DUPLICATE KEY UPDATE `DeviceParrentName` = (CASE WHEN (JSON_EXTRACT(`dv`.`DetailedInformation`, '$.""Drive Type Code""') = 3) THEN REPLACE(JSON_EXTRACT(`dv`.`DetailedInformation`, '$.""System Name""'), '""', '') ELSE NULL END), `DeviceHardwareDetails` = `dv`.`DetailedInformation`;"
    ReportFile.WriteLine "ALTER TABLE `device_details` AUTO_INCREMENT = 1; /* Ensure no end gaps are present in the auto incrementing sequence for Device */"
End Function
Function ApplySoftwareNormalizationForSoftwareInstalled(strComputer, ReportFile)
    ReportFile.WriteLine "/* Following sequence of MySQL queries will ensure Software List normalization and data retention, so a complete traceability would be ensured for each hostnames/devices included */"
    ReportFile.WriteLine "ALTER TABLE `publisher_details` AUTO_INCREMENT = 1; /* Making sure the Publisher table do not have a ending gap for ID auto numbering sequence */"
    ReportFile.WriteLine "INSERT `publisher_details` (`PublisherName`) SELECT `PublisherName` FROM `in_windows_software_installed` WHERE (`HostName` = '" & strComputer & "') AND (`PublisherName` IS NOT NULL) AND (`PublisherName` NOT IN (SELECT `PublisherName` FROM `publisher_details` GROUP BY `PublisherName`)) GROUP BY `PublisherName`; /* Publishers consolidation */"
    ReportFile.WriteLine "ALTER TABLE `software_details` AUTO_INCREMENT = 1; /* Making sure the Software table do not have a ending gap for ID auto numbering sequence */"
    ReportFile.WriteLine "INSERT `software_details` (`SoftwareName`) SELECT `SoftwareName` FROM `in_windows_software_installed` WHERE (`HostName` = '" & strComputer & "') AND (`SoftwareName` IS NOT NULL) AND (`SoftwareName` NOT IN (SELECT `SoftwareName` FROM `software_details` GROUP BY `SoftwareName`)) GROUP BY `SoftwareName`; /* Software consolidation */"
    ReportFile.WriteLine "ALTER TABLE `version_details` AUTO_INCREMENT = 1; /* Making sure the Version table do not have a ending gap for ID auto numbering sequence */"
    ReportFile.WriteLine "INSERT `version_details` (`FullVersion`) SELECT `FullVersion` FROM `in_windows_software_installed` WHERE (`HostName` = '" & strComputer & "') AND (`FullVersion` IS NOT NULL) AND (`FullVersion` NOT IN (SELECT `FullVersion` FROM `version_details` GROUP BY `FullVersion`)) GROUP BY `FullVersion`; /* Version consolidation */"
    ReportFile.WriteLine "ALTER TABLE `evaluation_headers` AUTO_INCREMENT = 1; /* Making sure the Evaluation header table do not have a ending gap for ID auto numbering sequence */"
    ReportFile.WriteLine "INSERT INTO `evaluation_headers` (`DeviceId`, `DateOfGatheringTimestampFirst`, `DateOfGatheringTimestampLast`) SELECT `dd`.`DeviceId`, MIN(`iwsl`.`EvaluationTimestamp`), MAX(`iwsl`.`EvaluationTimestamp`) FROM `device_details` `dd` INNER JOIN `in_windows_software_installed` `iwsl` ON `dd`.`DeviceName` = `iwsl`.`HostName` WHERE (`dd`.`DeviceName` = '" & strComputer & "'); /* Evaluation initiation */"
    ReportFile.WriteLine "SELECT LAST_INSERT_ID() INTO @EvaluationId; /* Capture the Id for current evaluation captured to be used later on */"
    ReportFile.WriteLine "INSERT INTO `evaluation_lines` (`EvaluationId`, `PublisherId`, `SoftwareId`, `VersionId`, `InstallationDate`) SELECT DISTINCT @EvaluationId, `PublisherId`, `SoftwareId`, `VersionId`, MAX(`InstallationDate`) FROM `in_windows_software_installed` `iwsl` INNER JOIN `publisher_details` `pd` ON `iwsl`.`PublisherName` = `pd`.`PublisherName` INNER JOIN `software_details` `sd` ON `iwsl`.`SoftwareName` = `sd`.`SoftwareName` INNER JOIN `version_details` `vd` ON `iwsl`.`FullVersion` = `vd`.`FullVersion` WHERE (`iwsl`.`HostName` = '" & strComputer & "') GROUP BY `pd`.`PublisherId`, `sd`.`SoftwareId`, `vd`.`VersionId` ORDER BY `pd`.`PublisherId`, `sd`.`SoftwareId`, `vd`.`VersionId`; /* Populate the software installed in the structure that ensures traceability over time */"
    ReportFile.WriteLine "UPDATE `device_details` SET `MostRecentEvaluationId` = @EvaluationId, `LastSeen` = `LastSeen` WHERE (`DeviceName` = '" & strComputer & "'); /* Sets the most recent evaluation against relevant device to make easier software version comparison(s) between considering the latest details */"
End Function
Function ApplySoftwareNormalizationForSoftwarePortable(ReportFile)
    ReportFile.WriteLine "/* Following sequence of MySQL queries will ensure Portable Software List normalization and data retention, so a complete traceability would be ensured for each hostnames/devices included */"
    ReportFile.WriteLine "ALTER TABLE `publisher_details` AUTO_INCREMENT = 1; /* Making sure the Publisher table do not have a ending gap for ID auto numbering sequence */"
    ReportFile.WriteLine "INSERT `publisher_details` (`PublisherName`) SELECT `sf`.`PublisherName` FROM `in_windows_software_portable` `iwsp` LEFT JOIN  `software_files` `sf` ON `iwsp`.`FileNameSearched` = `sf`.`SoftwareFileName` WHERE (`sf`.`PublisherName` IS NOT NULL) AND (`sf`.`PublisherName` NOT IN (SELECT `PublisherName` FROM `publisher_details` GROUP BY `PublisherName`)) GROUP BY `sf`.`PublisherName`; /* Publishers consolidation */"
    ReportFile.WriteLine "ALTER TABLE `software_details` AUTO_INCREMENT = 1; /* Making sure the Software table do not have a ending gap for ID auto numbering sequence */"
    ReportFile.WriteLine "INSERT `software_details` (`SoftwareName`) SELECT `sf`.`SoftwareName` FROM `in_windows_software_portable` `iwsp` LEFT JOIN  `software_files` `sf` ON `iwsp`.`FileNameSearched` = `sf`.`SoftwareFileName` WHERE (`sf`.`SoftwareName` IS NOT NULL) AND (`sf`.`SoftwareName` NOT IN (SELECT `SoftwareName` FROM `software_details` GROUP BY `SoftwareName`)) GROUP BY `sf`.`SoftwareName`; /* Software consolidation */"
    ReportFile.WriteLine "ALTER TABLE `version_details` AUTO_INCREMENT = 1; /* Making sure the Version table do not have a ending gap for ID auto numbering sequence */"
    ReportFile.WriteLine "INSERT `version_details` (`FullVersion`) SELECT `FileVersionFound` FROM `in_windows_software_portable` WHERE (`FileVersionFound` IS NOT NULL) AND (`FileVersionFound` NOT IN('', 'v', 'v0', 'v0.0', 'v0.0.0', 'v0.0.0.0')) AND (`FileVersionFound` NOT IN (SELECT `FullVersion` FROM `version_details` GROUP BY `FullVersion`)) GROUP BY `FileVersionFound`; /* Version consolidation */"
    ReportFile.WriteLine "ALTER TABLE `evaluation_headers` AUTO_INCREMENT = 1; /* Making sure the Evaluation header table do not have a ending gap for ID auto numbering sequence */"
    ReportFile.WriteLine "INSERT INTO `evaluation_headers` (`DeviceId`, `DateOfGatheringTimestampFirst`, `DateOfGatheringTimestampLast`) SELECT `dd`.`DeviceId`, MIN(`EvaluationTimestamp`), MAX(`EvaluationTimestamp`) FROM `device_details` `dd` INNER JOIN `in_windows_software_portable` `iwsp` ON `dd`.`DeviceName` = `iwsp`.`VolumeSerialNumber` GROUP BY `dd`.`DeviceName`; /* Evaluation initiation */"
    ReportFile.WriteLine "INSERT INTO `evaluation_lines` (`EvaluationId`, `PublisherId`, `SoftwareId`, `VersionId`, `InstallationDate`, `Folders`) SELECT `eh`.`EvaluationId`, `pd`.`PublisherId`, `sd`.`SoftwareId`, `vd`.`VersionId`, CAST(MAX(CASE WHEN `iwsp`.`FileDateCreated` < `iwsp`.`FileDateLastModified` THEN `iwsp`.`FileDateCreated` ELSE `iwsp`.`FileDateLastModified` END) AS DATE), GROUP_CONCAT(DISTINCT `FilePathFound` ORDER BY `FilePathFound` SEPARATOR '; ') FROM `in_windows_software_portable` `iwsp` INNER JOIN `version_details` `vd` ON `iwsp`.`FileVersionFound` = `vd`.`FullVersion` INNER JOIN `software_files` `sf` ON ((`sf`.`SoftwareFileName` = `iwsp`.`FileNameSearched`) AND (`vd`.`FullVersionNumeric` BETWEEN `sf`.`SoftwareFileVersionNumericFirst` AND `sf`.`SoftwareFileVersionNumericLast`)) INNER JOIN `device_details` `dd` ON `iwsp`.`VolumeSerialNumber` = `dd`.`DeviceName` INNER JOIN `evaluation_headers` `eh` ON `dd`.`DeviceId` = `eh`.`DeviceId` LEFT JOIN `software_details` `sd` ON `sf`.`SoftwareName` = `sd`.`SoftwareName` LEFT JOIN `publisher_details` `pd` ON `sf`.`PublisherName` = `pd`.`PublisherName` WHERE (`iwsp`.`EvaluationTimestamp` BETWEEN `eh`.`DateOfGatheringTimestampFirst` AND  `eh`.`DateOfGatheringTimestampLast`) GROUP BY `eh`.`EvaluationId`, `pd`.`PublisherId`, `sd`.`SoftwareId`, `vd`.`VersionId` ORDER BY `eh`.`EvaluationId`, `pd`.`PublisherId`, `sd`.`SoftwareId`, `vd`.`VersionId`; /* Evaluation details */"
    ReportFile.WriteLine "CALL `pr_MatchLatestEvaluationForSoftwarePortable`(); /* Sets the most recent evaluation against relevant device to make easier software version comparison(s) between considering the latest details */"
End Function
Function ApplySoftwareNormalizationForSecurityRiskComponents(ReportFile)
    ReportFile.WriteLine "/* Following sequence of MySQL queries will ensure Portable Software List normalization and data retention, so a complete traceability would be ensured for each hostnames/devices included */"
    ReportFile.WriteLine "ALTER TABLE `publisher_details` AUTO_INCREMENT = 1; /* Making sure the Publisher table do not have a ending gap for ID auto numbering sequence */"
    ReportFile.WriteLine "INSERT `publisher_details` (`PublisherName`) SELECT `sf`.`PublisherName` FROM `in_windows_security_risk_components` `iwsrc` LEFT JOIN  `software_files` `sf` ON `iwsrc`.`FileNameSearched` = `sf`.`SoftwareFileName` WHERE (`sf`.`PublisherName` IS NOT NULL) AND (`sf`.`PublisherName` NOT IN (SELECT `PublisherName` FROM `publisher_details` GROUP BY `PublisherName`)) GROUP BY `sf`.`PublisherName`; /* Publishers consolidation */"
    ReportFile.WriteLine "ALTER TABLE `software_details` AUTO_INCREMENT = 1; /* Making sure the Software table do not have a ending gap for ID auto numbering sequence */"
    ReportFile.WriteLine "INSERT `software_details` (`SoftwareName`) SELECT `sf`.`SoftwareName` FROM `in_windows_security_risk_components` `iwsrc` LEFT JOIN  `software_files` `sf` ON `iwsrc`.`FileNameSearched` = `sf`.`SoftwareFileName` WHERE (`sf`.`SoftwareName` IS NOT NULL) AND (`sf`.`SoftwareName` NOT IN (SELECT `SoftwareName` FROM `software_details` GROUP BY `SoftwareName`)) GROUP BY `sf`.`SoftwareName`; /* Software consolidation */"
    ReportFile.WriteLine "ALTER TABLE `version_details` AUTO_INCREMENT = 1; /* Making sure the Version table do not have a ending gap for ID auto numbering sequence */"
    ReportFile.WriteLine "INSERT `version_details` (`FullVersion`) SELECT `FileVersionFound` FROM `in_windows_security_risk_components` WHERE (`FileVersionFound` IS NOT NULL) AND (`FileVersionFound` NOT IN('', 'v', 'v0', 'v0.0', 'v0.0.0', 'v0.0.0.0')) AND (`FileVersionFound` NOT IN (SELECT `FullVersion` FROM `version_details` GROUP BY `FullVersion`)) GROUP BY `FileVersionFound`; /* Version consolidation */"
    ReportFile.WriteLine "ALTER TABLE `security_evaluation_headers` AUTO_INCREMENT = 1; /* Making sure the Evaluation header table do not have a ending gap for ID auto numbering sequence */"
    ReportFile.WriteLine "INSERT INTO `security_evaluation_headers` (`DeviceId`, `DateOfGatheringTimestampFirst`, `DateOfGatheringTimestampLast`) SELECT `dd`.`DeviceId`, MIN(`EvaluationTimestamp`), MAX(`EvaluationTimestamp`) FROM `device_details` `dd` INNER JOIN `in_windows_security_risk_components` `iwsrc` ON `dd`.`DeviceName` = `iwsrc`.`VolumeSerialNumber` GROUP BY `dd`.`DeviceName`; /* Evaluation initiation */"
    ReportFile.WriteLine "INSERT INTO `security_evaluation_lines` (`SecurityEvaluationId`, `PublisherId`, `SoftwareId`, `VersionId`, `InstallationDate`, `Folders`) SELECT `seh`.`SecurityEvaluationId`, `pd`.`PublisherId`, `sd`.`SoftwareId`, `vd`.`VersionId`, CAST(MAX(CASE WHEN `iwsrc`.`FileDateCreated` < `iwsrc`.`FileDateLastModified` THEN `iwsrc`.`FileDateCreated` ELSE `iwsrc`.`FileDateLastModified` END) AS DATE), GROUP_CONCAT(DISTINCT `FilePathFound` ORDER BY `FilePathFound` SEPARATOR '; ') FROM `in_windows_security_risk_components` `iwsrc` INNER JOIN `version_details` `vd` ON `iwsrc`.`FileVersionFound` = `vd`.`FullVersion` INNER JOIN `software_files` `sf` ON ((`sf`.`SoftwareFileName` = `iwsrc`.`FileNameSearched`) AND (`vd`.`FullVersionNumeric` BETWEEN `sf`.`SoftwareFileVersionNumericFirst` AND `sf`.`SoftwareFileVersionNumericLast`)) INNER JOIN `device_details` `dd` ON `iwsrc`.`VolumeSerialNumber` = `dd`.`DeviceName` INNER JOIN `security_evaluation_headers` `seh` ON `dd`.`DeviceId` = `seh`.`DeviceId` LEFT JOIN `software_details` `sd` ON `sf`.`SoftwareName` = `sd`.`SoftwareName` LEFT JOIN `publisher_details` `pd` ON `sf`.`PublisherName` = `pd`.`PublisherName` WHERE (`iwsrc`.`EvaluationTimestamp` BETWEEN `seh`.`DateOfGatheringTimestampFirst` AND  `seh`.`DateOfGatheringTimestampLast`) GROUP BY `seh`.`SecurityEvaluationId`, `pd`.`PublisherId`, `sd`.`SoftwareId`, `vd`.`VersionId` ORDER BY `seh`.`SecurityEvaluationId`, `pd`.`PublisherId`, `sd`.`SoftwareId`, `vd`.`VersionId`; /* Evaluation details */"
    ReportFile.WriteLine "CALL `pr_MatchLatestEvaluationForSecurityRiskComponents`(); /* Sets the most recent evaluation against relevant device to make easier software version comparison(s) between considering the latest details */"
End Function
Function BuildInsertOrUpdateSQLstructure(aryFieldNames, aryFieldValues, strInsertOrUpdate, intFirstNumberOfFieldsToIgnore, intLastNumberOfFieldsToIgnore)
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
                If ((Counter >= intFirstNumberOfFieldsToIgnore) And (Counter <= (intFieldNumbered - intLastNumberOfFieldsToIgnore))) Then
                    If (Counter > intFirstNumberOfFieldsToIgnore) Then
                        strUpdateSQLstructure = strUpdateSQLstructure & ", "
                    End If
                    strUpdateSQLstructure = strUpdateSQLstructure & "`" & strFieldName & "` = '" & aryFieldValues(Counter) & "'"
                End If
                Counter = Counter + 1
            Next
    End Select
    BuildInsertOrUpdateSQLstructure = Replace(Replace(Replace(strUpdateSQLstructure, "'NULL'", "NULL"), "\", "\\"), "t's", "t\'s")
End Function
Function CheckSoftware(strComputer, bolWriteHeader, ReportFile, objReg, strKey)
    Dim aryJSONinformationCSV(7)
    Dim aryJSONinformationSQL(7)
    aryInformationToExpose = Array(_
        "Evaluation Timestamp", _
        "Host Name", _
        "Publisher Name", _
        "Software Name", _
        "Full Version", _
        "Installation Date", _
        "Other Info", _
        "Registry Key Trunk", _
        "Registry SubKey" _
    )
    aryInformationToExposeOtherInfo = Array(_
        "Install Location", _
        "Publisher Name", _
        "Size [bytes]", _
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
                strPublisherName = HarmonizedPublisher(strPublisher)
            Else
                strPublisherName = "NULL"
            End If
            strSoftwareNameCleaned = HarmonizedSoftwareName(strDisplayName)
            If ((intReturnL <> 0) Or (Len(Trim(strInstallLocation)) = 0)) Then
                strInstallLocation = "NULL"
            Else
                strInstallLocation = Replace(strInstallLocation, "\", "\\")
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
                ' as LAME software is an just an encoder for MP3 the version seem to require special handling for both FullVersion and Publisher
                Select Case strSoftwareNameCleaned
                    Case "LAME"
                        aryDisplayName = Split(strDisplayName, " ")
                        strPublisherName = aryDisplayName(1)
                        strDisplayVersionCleaned = aryDisplayName(2)
                    Case "Double Commander"
                        strPublisherName = "alexx2000"
                        aryDisplayName = Split(strDisplayName, " ")
                        strDisplayVersionCleaned = aryDisplayName(2)
                End Select
            Else
                Select Case strSoftwareNameCleaned
                    ' AIMP seems to store DisplayVersion as "vX.XX.XXXX, Date" so needs special handling to get only the first part without the ","
                    Case "AIMP"
                        aryDisplayVersion = Split(strDisplayVersion, ", ")
                        strDisplayVersionCleaned = aryDisplayVersion(0)
                    Case Else
                        ' In some cases DisplayVersion has a date before the version so we're going to take in consideration only the very last group of continuous string splitted by space
                        strDisplayVersionPieces = Split(CStr(Replace(Replace(strDisplayVersion, " beta ", "."), "a", ".")), " ")
                        For Each strDisplayVersionPieceValue In strDisplayVersionPieces
                            If (IsNumeric(strDisplayVersionPieceValue)) Then
                                strDisplayVersionCleaned = strVersionPrefix & strDisplayVersionPieceValue
                            End If
                        Next
                End Select
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
                strInstallLocation, _
                strPublisher, _
                strSizeInBytes, _
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
                        crtValue = Trim(aryValuesToExposeOtherInfo(intCounter))
                        If (IsNull(crtValue)) Then
                            crtValue = "-"
                        End If
                        If (IsNumericExtended(crtValue)) Then
                            aryJSONinformationSQL(intCounter) = """" & crtOtherInfo & """: " & crtValue
                        Else
                            aryJSONinformationSQL(intCounter) = """" & crtOtherInfo & """: " & """" & crtValue & """"
                        End If
                        intCounter = intCounter + 1
                    Next
                    strOtherInfo = "{ " & Join(aryJSONinformationSQL, ", ") & " }"
            End Select
            aryValuesToExpose = Array(_
                CurrentDateTimeToSqlFormat(), _
                strComputer, _
                strPublisherName, _
                strSoftwareNameCleaned, _
                strDisplayVersionCleaned, _
                strDateYMD , _
                strOtherInfo, _
                strSubkeyPieces(1), _
                strSubkey _
            )
            Select Case LCase(strResultFileType)
                Case ".csv"
                    ReportFile.WriteLine Join(aryValuesToExpose, strFieldSeparator)
                Case ".sql"
                    ReportFile.WriteLine "INSERT INTO `in_windows_software_installed` (" & _
                        BuildInsertOrUpdateSQLstructure(aryInformationToExpose, aryValuesToExpose, "InsertFields", 0, 2) & _
                        ") VALUES(" & _
                        BuildInsertOrUpdateSQLstructure(aryInformationToExpose, aryValuesToExpose, "InsertValues", 0, 2) & _
                        ") ON DUPLICATE KEY UPDATE " & _
                        BuildInsertOrUpdateSQLstructure(aryInformationToExpose, aryValuesToExpose, "Update", 0, 2) & _
                        ";"
            End Select
        End If
    Next
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
Function CleanStringBeforeOrAfterNumber(strFullStringToClean, strBeforeOrAfter, aryBlackList, strStringToReplaceWith)
    intPieceCounter = 0
    For Each strBlackListPiece In aryBlackList
        ' break entire string into pieces with space as separator
        aryFullStringToClean = Split(strFullStringToClean, " ")
        intLastPieceNumber = UBound(aryFullStringToClean)
        strCleanedString = ""
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
                strFullStringToClean = strCleanedString
            End If
            intPieceCounter = intPieceCounter + 1
        Next
    Next
    CleanStringBeforeOrAfterNumber = strCleanedString
End Function
Function ConvertDateTimeToSqlFormat(dtGivenDate)
    ConvertDateTimeToSqlFormat = ConvertDateToSqlFormat(dtGivenDate) & _
        " " & NumberWithTwoDigits(DatePart("h", dtGivenDate)) & _
        ":" & NumberWithTwoDigits(DatePart("n", dtGivenDate)) & _
        ":" & NumberWithTwoDigits(DatePart("s", dtGivenDate))
End Function
Function ConvertDateToSqlFormat(dtGivenDate)
    ConvertDateToSqlFormat = DatePart("yyyy", dtGivenDate) & _
        "-" & NumberWithTwoDigits(DatePart("m", dtGivenDate)) & _
        "-" & NumberWithTwoDigits(DatePart("d", dtGivenDate))
End Function
Function CSVfieldNamesIntoSQLfieldName(aryFieldNames)
    strFieldListForMySQLinsert = Join(aryFieldNames, "|")
    strFieldListForMySQLinsert = CleanStringWithBlacklistArray(strFieldListForMySQLinsert, Array("[bytes]"), "Bytes")
    strFieldListForMySQLinsert = CleanStringWithBlacklistArray(strFieldListForMySQLinsert, Array("(major.minor)"), "Major Minor")
    strFieldListForMySQLinsert = CleanStringWithBlacklistArray(strFieldListForMySQLinsert, Array(" "), "")
    CSVfieldNamesIntoSQLfieldName = strFieldListForMySQLinsert
End Function
Function CurrentDateTimeToSqlFormat()
    CurrentDateTimeToSqlFormat = ConvertDateTimeToSqlFormat(Now())
End Function
Function CurrentDateToSqlFormat()
    CurrentDateToSqlFormat = ConvertDateToSqlFormat(Now())
End Function
Function CurrentOperatingSystemVersionForComparison()
    intOSVersion = 0
    ' only required to be able to differentiate a few attributes present only in modern OS versions
    Set objOperatingSystem = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
    For Each crtObjOS in objOperatingSystem
        aryVersionParts = Split(crtObjOS.Version, ".")
        intOSVersion = CInt(aryVersionParts(0)) * 10 + aryVersionParts(1)
    Next
    CurrentOperatingSystemVersionForComparison = intOSVersion
End Function
Function CurrentComputerName(strGivenComputerName)
    If ((LCase(strGivenComputerName) = "localhost") Or (strGivenComputerName = "127.0.0.1") Or (strGivenComputerName = "::1")) Then
        Set objSysInfo = CreateObject("WinNTSystemInfo")
        CurrentComputerName = objSysInfo.ComputerName
    Else
        CurrentComputerName = strGivenComputerName
    End If
End Function
Function FolderHasSubfolders(oFolder)
    On Error Resume Next
    FolderHasSubfolders = (oFolder.SubFolders.Count >= 0)
End Function
Function IsNumericExtended(InValueToEvaluate)
    If (IsNumeric(InValueToEvaluate)) Then
        intLengthDelta = Len(InValueToEvaluate) - Len(Replace(InValueToEvaluate, ".", ""))
        If ((Left(InValueToEvaluate, 1) <> "0") And (Left(InValueToEvaluate, 1) <> "F") And (Left(InValueToEvaluate, 1) <> "T") And (intLengthDelta < 2)) Then
            IsNumericExtended = True
        Else
            IsNumericExtended = False
        End If
    Else
        IsNumericExtended = False
    End If
End Function
Function HarmonizedPublisher(strPublisherName)
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
        Array("Realtek", "Realtek Semiconductor Corp."), _
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
        HarmonizedPublisher = strPublisherName
    Else
        HarmonizedPublisher = strPublishersHarmonized
    End If
End Function
Function HarmonizedSoftwareName(strSoftwareName)
    Dim strSoftwareNameWIP
    strSoftwareNameWIP = CleanStringStartEnd(strSoftwareName, " (", ")")
    aryBlackListToRemove = Array("Update") ' to properly clean "Java <No> Update <No>" software name
    strSoftwareNameWIP = CleanStringBeforeOrAfterNumber(strSoftwareNameWIP, "Before", aryBlackListToRemove, "")
    aryBlackListToRemoveAfter = Array("R2", "LAME") ' to properly clean "Microsoft SQL Server <No> R2 Native Client" software name
    strSoftwareNameWIP = CleanStringBeforeOrAfterNumber(strSoftwareNameWIP, "After", aryBlackListToRemoveAfter, "")
    aryBlackListToClean = Array("(x86_x64)", "(x64)", "(x86)", "_WHQL", "_X64", "_X86", "64-bit", "beta", "en-us", "for x64", "for x86", "SP1", "SP2", "SP3", "version", "VS2005", "VS2008", "VS2010", "VS2012", "VS2015", "x64", "x86")
    strSoftwareNameWIP = CleanStringWithBlacklistArray(strSoftwareNameWIP, aryBlackListToClean, "")
    aryBlackListToReplaceWithSpace = Array(" -")
    strSoftwareNameWIP = CleanStringWithBlacklistArray(strSoftwareNameWIP, aryBlackListToReplaceWithSpace, " ")
    aryBlackListToReplaceWithOriginal = Array("(R)")
    strSoftwareNameWIP = CleanStringWithBlacklistArray(strSoftwareNameWIP, aryBlackListToReplaceWithOriginal, Chr(174))
    strSoftwareNameWIP = CleanStringOfNumericPiece(strSoftwareNameWIP)
    HarmonizedSoftwareName = strSoftwareNameWIP
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
Function MappingDriveTypeCodeInDescriptionOut(InputDriveType)
    aryDriveTypes = Array(_
        Array(0, "Unknown"), _
        Array(1, "No Root Directory"), _
        Array(2, "Removable Disk"), _
        Array(3, "Local Disk"), _
        Array(4, "Network Drive"), _
        Array(5, "Compact Disc"), _
        Array(6, "RAM Disk") _
    )
    For Each crtDriveType In aryDriveTypes
        If (InputDriveType = crtDriveType(0)) Then
            MappingDriveTypeCodeInDescriptionOut = crtDriveType(1)
        End If
    Next
End Function
Function MappingLanguageLCIDinDescriptionOut(GivenElement, GivenValue, FeedbackElement)
    aryLanguageCodes = Array(_
        Array("10241", "2801", "Arabic - Syria"), _
        Array("10249", "2809", "English - Belize"), _
        Array("1025", "0401", "Arabic - Saudi Arabia"), _
        Array("10250", "280a", "Spanish - Peru"), _
        Array("10252", "280c", "French - Senegal"), _
        Array("1026", "0402", "Bulgarian"), _
        Array("1027", "0403", "Catalan"), _
        Array("1028", "0404", "Chinese - Taiwan"), _
        Array("1029", "0405", "Czech"), _
        Array("1030", "0406", "Danish"), _
        Array("1031", "0407", "German - Germany"), _
        Array("1032", "0408", "Greek"), _
        Array("1033", "0409", "English - United States"), _
        Array("1034", "040a", "Spanish - Spain Array(Traditional Sort)"), _
        Array("1035", "040b", "Finnish"), _
        Array("1036", "040c", "French - France"), _
        Array("1037", "040d", "Hebrew"), _
        Array("1038", "040e", "Hungarian"), _
        Array("1039", "040f", "Icelandic"), _
        Array("1040", "0410", "Italian - Italy"), _
        Array("1041", "0411", "Japanese"), _
        Array("1042", "0412", "Korean"), _
        Array("1043", "0413", "Dutch - Netherlands"), _
        Array("1044", "0414", "Norwegian Array(Bokmål)"), _
        Array("1045", "0415", "Polish"), _
        Array("1046", "0416", "Portuguese - Brazil"), _
        Array("1047", "0417", "Rhaeto-Romanic"), _
        Array("1048", "0418", "Romanian"), _
        Array("1049", "0419", "Russian"), _
        Array("1050", "041a", "Croatian"), _
        Array("1051", "041b", "Slovak"), _
        Array("1052", "041c", "Albanian - Albania"), _
        Array("1053", "041d", "Swedish"), _
        Array("1054", "041e", "Thai"), _
        Array("1055", "041f", "Turkish"), _
        Array("1056", "0420", "Urdu"), _
        Array("1057", "0421", "Indonesian"), _
        Array("1058", "0422", "Ukrainian"), _
        Array("1059", "0423", "Belarusian"), _
        Array("1060", "0424", "Slovenian"), _
        Array("1061", "0425", "Estonian"), _
        Array("1062", "0426", "Latvian"), _
        Array("1063", "0427", "Lithuanian"), _
        Array("1064", "0428", "Tajik"), _
        Array("1065", "0429", "Farsi"), _
        Array("1066", "042a", "Vietnamese"), _
        Array("1067", "042b", "Armenian - Armenia"), _
        Array("1068", "042c", "Azeri Array(Latin)"), _
        Array("1069", "042d", "Basque"), _
        Array("1070", "042e", "Sorbian"), _
        Array("1071", "042f", "FYRO Macedonian"), _
        Array("1072", "0430", "Sutu"), _
        Array("1073", "0431", "Tsonga"), _
        Array("1074", "0432", "Tswana"), _
        Array("1075", "0433", "Venda"), _
        Array("1076", "0434", "Xhosa"), _
        Array("1077", "0435", "Zulu"), _
        Array("1078", "0436", "Afrikaans - South Africa"), _
        Array("1079", "0437", "Georgian"), _
        Array("1080", "0438", "Faroese"), _
        Array("1081", "0439", "Hindi"), _
        Array("1082", "043a", "Maltese"), _
        Array("1083", "043b", "Sami Array(Lappish)"), _
        Array("1084", "043c", "Scottish Gaelic"), _
        Array("1085", "043d", "Yiddish"), _
        Array("1086", "043e", "Malay - Malaysia"), _
        Array("1087", "043f", "Kazakh"), _
        Array("1088", "0440", "Kyrgyz Array(Cyrillic)"), _
        Array("1089", "0441", "Swahili"), _
        Array("1090", "0442", "Turkmen"), _
        Array("1091", "0443", "Uzbek Array(Latin)"), _
        Array("1092", "0444", "Tatar"), _
        Array("1093", "0445", "Bengali Array(India)"), _
        Array("1094", "0446", "Punjabi"), _
        Array("1095", "0447", "Gujarati"), _
        Array("1096", "0448", "Oriya"), _
        Array("1097", "0449", "Tamil"), _
        Array("1098", "044a", "Telugu"), _
        Array("1099", "044b", "Kannada"), _
        Array("1100", "044c", "Malayalam"), _
        Array("1101", "044d", "Assamese"), _
        Array("1102", "044e", "Marathi"), _
        Array("1103", "044f", "Sanskrit"), _
        Array("1104", "0450", "Mongolian Array(Cyrillic)"), _
        Array("1105", "0451", "Tibetan - People's Republic of China"), _
        Array("1106", "0452", "Welsh"), _
        Array("1107", "0453", "Khmer"), _
        Array("1108", "0454", "Lao"), _
        Array("1109", "0455", "Burmese"), _
        Array("1110", "0456", "Galician"), _
        Array("1111", "0457", "Konkani"), _
        Array("1112", "0458", "Manipuri"), _
        Array("1113", "0459", "Sindhi - India"), _
        Array("1114", "045a", "Syriac"), _
        Array("1115", "045b", "Sinhalese - Sri Lanka"), _
        Array("1116", "045c", "Cherokee - United States"), _
        Array("1117", "045d", "Inuktitut"), _
        Array("1118", "045e", "Amharic - Ethiopia"), _
        Array("1119", "045f", "Tamazight Array(Arabic)"), _
        Array("1120", "0460", "Kashmiri Array(Arabic)"), _
        Array("1121", "0461", "Nepali"), _
        Array("1122", "0462", "Frisian - Netherlands"), _
        Array("1123", "0463", "Pashto"), _
        Array("1124", "0464", "Filipino"), _
        Array("1125", "0465", "Divehi"), _
        Array("1126", "0466", "Edo"), _
        Array("11265", "2c01", "Arabic - Jordan"), _
        Array("1127", "0467", "Fulfulde - Nigeria"), _
        Array("11273", "2c09", "English - Trinidad"), _
        Array("11274", "2c0a", "Spanish - Argentina"), _
        Array("11276", "2c0c", "French - Cameroon"), _
        Array("1128", "0468", "Hausa - Nigeria"), _
        Array("1129", "0469", "Ibibio - Nigeria"), _
        Array("1130", "046a", "Yoruba"), _
        Array("1131", "046B", "Quecha - Bolivia"), _
        Array("1132", "046c", "Sepedi"), _
        Array("1133", "046d", "Bashkir"), _
        Array("1134", "046e", "Luxembourgish"), _
        Array("1135", "046f", "Greenlandic"), _
        Array("1136", "0470", "Igbo - Nigeria"), _
        Array("1137", "0471", "Kanuri - Nigeria"), _
        Array("1138", "0472", "Oromo"), _
        Array("1139", "0473", "Tigrigna - Ethiopia"), _
        Array("1140", "0474", "Guarani - Paraguay"), _
        Array("1141", "0475", "Hawaiian - United States"), _
        Array("1142", "0476", "Latin"), _
        Array("1143", "0477", "Somali"), _
        Array("1144", "0478", "Yi"), _
        Array("1145", "0479", "Papiamentu"), _
        Array("1146", "0471", "Mapudungun"), _
        Array("1148", "047c", "Mohawk"), _
        Array("1150", "047e", "Breton"), _
        Array("1152", "0480", "Uighur - China"), _
        Array("1153", "0481", "Maori - New Zealand"), _
        Array("1154", "0482", "Occitan"), _
        Array("1155", "0483", "Corsican"), _
        Array("1156", "0484", "Alsatian"), _
        Array("1157", "0485", "Yakut"), _
        Array("1158", "0486", "K'iche"), _
        Array("1159", "0487", "Kinyarwanda"), _
        Array("1160", "0488", "Wolof"), _
        Array("1164", "048c", "Dari"), _
        Array("12289", "3001", "Arabic - Lebanon"), _
        Array("12297", "3009", "English - Zimbabwe"), _
        Array("12298", "300a", "Spanish - Ecuador"), _
        Array("12300", "300c", "French - Cote d'Ivoire"), _
        Array("1279", "04ff", "HID Array(Human Interface Device)"), _
        Array("13313", "3401", "Arabic - Kuwait"), _
        Array("13321", "3409", "English - Philippines"), _
        Array("13322", "340a", "Spanish - Chile"), _
        Array("13324", "340c", "French - Mali"), _
        Array("14337", "3801", "Arabic - U.A.E."), _
        Array("14345", "3809", "English - Indonesia"), _
        Array("14346", "380a", "Spanish - Uruguay"), _
        Array("14348", "380c", "French - Morocco"), _
        Array("15361", "3c01", "Arabic - Bahrain"), _
        Array("15369", "3c09", "English - Hong Kong SAR"), _
        Array("15370", "3c0a", "Spanish - Paraguay"), _
        Array("15372", "3c0c", "French - Haiti"), _
        Array("16385", "4001", "Arabic - Qatar"), _
        Array("16393", "4009", "English - India"), _
        Array("16394", "400a", "Spanish - Bolivia"), _
        Array("17417", "4409", "English - Malaysia"), _
        Array("17418", "440a", "Spanish - El Salvador"), _
        Array("18441", "4809", "English - Singapore"), _
        Array("18442", "480a", "Spanish - Honduras"), _
        Array("19466", "4c0a", "Spanish - Nicaragua"), _
        Array("2049", "0801", "Arabic - Iraq"), _
        Array("20490", "500a", "Spanish - Puerto Rico"), _
        Array("2052", "0804", "Chinese - People's Republic of China"), _
        Array("2055", "0807", "German - Switzerland"), _
        Array("2057", "0809", "English - United Kingdom"), _
        Array("2058", "080a", "Spanish - Mexico"), _
        Array("2060", "080c", "French - Belgium"), _
        Array("2064", "0810", "Italian - Switzerland"), _
        Array("2067", "0813", "Dutch - Belgium"), _
        Array("2068", "0814", "Norwegian Array(Nynorsk)"), _
        Array("2070", "0816", "Portuguese - Portugal"), _
        Array("2072", "0818", "Romanian - Moldava"), _
        Array("2073", "0819", "Russian - Moldava"), _
        Array("2074", "081a", "Serbian Array(Latin)"), _
        Array("2077", "081d", "Swedish - Finland"), _
        Array("2080", "0820", "Urdu - India"), _
        Array("2092", "082c", "Azeri Array(Cyrillic)"), _
        Array("2108", "083c", "Irish"), _
        Array("2110", "083e", "Malay - Brunei Darussalam"), _
        Array("2115", "0843", "Uzbek Array(Cyrillic)"), _
        Array("2117", "0845", "Bengali Array(Bangladesh)"), _
        Array("2118", "0846", "Punjabi Array(Pakistan)"), _
        Array("2128", "0850", "Mongolian Array(Mongolian)"), _
        Array("2129", "0851", "Tibetan - Bhutan"), _
        Array("2137", "0859", "Sindhi - Pakistan"), _
        Array("2143", "085f", "Tamazight Array(Latin)"), _
        Array("2144", "0860", "Kashmiri"), _
        Array("2145", "0861", "Nepali - India"), _
        Array("21514", "540a", "Spanish - United States"), _
        Array("2155", "086B", "Quecha - Ecuador"), _
        Array("2163", "0873", "Tigrigna - Eritrea"), _
        Array("22538", "580a", "Spanish - Latin America"), _
        Array("3073", "0c01", "Arabic - Egypt"), _
        Array("3076", "0c04", "Chinese - Hong Kong SAR"), _
        Array("3079", "0c07", "German - Austria"), _
        Array("3081", "0c09", "English - Australia"), _
        Array("3082", "0c0a", "Spanish - Spain Array(Modern Sort)"), _
        Array("3084", "0c0c", "French - Canada"), _
        Array("3098", "0c1a", "Serbian Array(Cyrillic)"), _
        Array("3179", "0C6B", "Quecha - Peru"), _
        Array("4097", "1001", "Arabic - Libya"), _
        Array("4100", "1004", "Chinese - Singapore"), _
        Array("4103", "1007", "German - Luxembourg"), _
        Array("4105", "1009", "English - Canada"), _
        Array("4106", "100a", "Spanish - Guatemala"), _
        Array("4108", "100c", "French - Switzerland"), _
        Array("4122", "101a", "Croatian Array(Bosnia/Herzegovina)"), _
        Array("5121", "1401", "Arabic - Algeria"), _
        Array("5124", "1404", "Chinese - Macao SAR"), _
        Array("5127", "1407", "German - Liechtenstein"), _
        Array("5129", "1409", "English - New Zealand"), _
        Array("5130", "140a", "Spanish - Costa Rica"), _
        Array("5132", "140c", "French - Luxembourg"), _
        Array("5146", "141A", "Bosnian Array(Bosnia/Herzegovina)"), _
        Array("58380", "e40c", "French - North Africa"), _
        Array("6145", "1801", "Arabic - Morocco"), _
        Array("6153", "1809", "English - Ireland"), _
        Array("6154", "180a", "Spanish - Panama"), _
        Array("6156", "180c", "French - Monaco"), _
        Array("7169", "1c01", "Arabic - Tunisia"), _
        Array("7177", "1c09", "English - South Africa"), _
        Array("7178", "1c0a", "Spanish - Dominican Republic"), _
        Array("7180", "1c0c", "French - West Indies"), _
        Array("8193", "2001", "Arabic - Oman"), _
        Array("8201", "2009", "English - Jamaica"), _
        Array("8202", "200a", "Spanish - Venezuela"), _
        Array("8204", "200c", "French - Reunion"), _
        Array("9217", "2401", "Arabic - Yemen"), _
        Array("9225", "2409", "English - Caribbean"), _
        Array("9226", "240a", "Spanish - Colombia"), _
        Array("9228", "240c", "French - Democratic Rep. of Congo") _
    )
    For Each CurrentLanguageCode In aryLanguageCodes
        Select Case GivenElement
            Case "Language - Country/Region"
                Select Case FeedbackElement
                    Case "LCID Decimal"
                        If (CStr(GivenValue) = CurrentLanguageCode(2)) Then
                            MappingLanguageLCIDinDescriptionOut = CurrentLanguageCode(0)
                        End If
                    Case "LCID Hexadecimal"
                        If (CStr(GivenValue) = CurrentLanguageCode(2)) Then
                            MappingLanguageLCIDinDescriptionOut = CurrentLanguageCode(1)
                        End If
                End Select
            Case "LCID Decimal"
                Select Case FeedbackElement
                    Case "Language - Country/Region"
                        If (CStr(GivenValue) = CurrentLanguageCode(0)) Then
                            MappingLanguageLCIDinDescriptionOut = CurrentLanguageCode(2)
                        End If
                    Case "LCID Hexadecimal"
                        If (CStr(GivenValue) = CurrentLanguageCode(0)) Then
                            MappingLanguageLCIDinDescriptionOut = CurrentLanguageCode(1)
                        End If
                End Select
            Case "LCID Hexadecimal"
                Select Case FeedbackElement
                    Case "Language - Country/Region"
                        If (CStr(GivenValue) = CurrentLanguageCode(1)) Then
                            MappingLanguageLCIDinDescriptionOut = CurrentLanguageCode(2)
                        End If
                    Case "LCID Decimal"
                        If (CStr(GivenValue) = CurrentLanguageCode(1)) Then
                            MappingLanguageLCIDinDescriptionOut = CurrentLanguageCode(0)
                        End If
                End Select
        End Select
    Next
End Function
Function MappingLogicalDiskIdToSerialNumberOrViceVersa(objWMIService, strSearchString, strSearchDesired)
    strInitialLocalDisk = ReadWMI__Win32_LogicalDisk_LightInformationGlued(objWMIService)
    aryTwoParts = Split(strInitialLocalDisk, "||")
    aryDeviceId = Split(aryTwoParts(0), " ")
    aryVolumeSerialNumber = Split(aryTwoParts(1), " ")
    strValueToReturn = Null
    intCounter = 0
    Select Case strSearchDesired
        Case "Device ID"
            For Each crtVolumeSerialNumber in aryVolumeSerialNumber
                If (crtVolumeSerialNumber = strSearchString) Then
                    strValueToReturn = aryDeviceId(intCounter)
                End If
                intCounter = intCounter + 1
            Next
        Case "Volume Serial Number"
            For Each crtDeviceId in aryDeviceId
                If (crtDeviceId = strSearchString) Then
                    strValueToReturn = aryVolumeSerialNumber(intCounter)
                End If
                intCounter = intCounter + 1
            Next
    End Select
    MappingLogicalDiskIdToSerialNumberOrViceVersa = strValueToReturn
End Function
Function MappingMediaTypeCodeInDescriptionOut(InputMediaType)
    aryMediaTypes = Array(_
        Array(0, "Format is unknown"), _
        Array(1, "5¼-Inch Floppy Disk - 1.2 MB - 512 bytes/sector"), _
        Array(2, "3½-Inch Floppy Disk - 1.44 MB -512 bytes/sector"), _
        Array(3, "3½-Inch Floppy Disk - 2.88 MB - 512 bytes/sector"), _
        Array(4, "3½-Inch Floppy Disk - 20.8 MB - 512 bytes/sector"), _
        Array(5, "3½-Inch Floppy Disk - 720 KB - 512 bytes/sector"), _
        Array(6, "5¼-Inch Floppy Disk - 360 KB - 512 bytes/sector"), _
        Array(7, "5¼-Inch Floppy Disk - 360 KB - 512 bytes/sector"), _
        Array(8, "5¼-Inch Floppy Disk - 360 KB - 1024 bytes/sector"), _
        Array(9, "5¼-Inch Floppy Disk - 180 KB - 512 bytes/sector"), _
        Array(10, "5¼-Inch Floppy Disk - 160 KB - 512 bytes/sector"), _
        Array(11, "Removable media other than floppy"), _
        Array(12, "Fixed hard disk media"), _
        Array(13, "3½-Inch Floppy Disk - 120 MB - 512 bytes/sector"), _
        Array(14, "3½-Inch Floppy Disk - 640 KB - 512 bytes/sector"), _
        Array(15, "5¼-Inch Floppy Disk - 640 KB - 512 bytes/sector"), _
        Array(16, "5¼-Inch Floppy Disk - 720 KB - 512 bytes/sector"), _
        Array(17, "3½-Inch Floppy Disk - 1.2 MB - 512 bytes/sector"), _
        Array(18, "3½-Inch Floppy Disk - 1.23 MB - 1024 bytes/sector"), _
        Array(19, "5¼-Inch Floppy Disk - 1.23 MB - 1024 bytes/sector"), _
        Array(20, "3½-Inch Floppy Disk - 128 MB - 512 bytes/sector"), _
        Array(21, "3½-Inch Floppy Disk - 230 MB - 512 bytes/sector"), _
        Array(22, "8-Inch Floppy Disk - 256 KB - 128 bytes/sector") _
    )
    For Each crtMediaType In aryMediaTypes
        If (InputMediaType = crtMediaType(0)) Then
            MappingMediaTypeCodeInDescriptionOut = crtMediaType(1)
        End If
    Next
End Function
Function MappingOSTypeCodeInDescriptionOut(InputOSTypeCode)
    aryOSTypeInfos = Array(_
        Array(0, "Unknown"), _
        Array(1, "Other"), _
        Array(2, "MACOS"), _
        Array(3, "ATTUNIX"), _
        Array(4, "DGUX"), _
        Array(5, "DECNT"), _
        Array(6, "Digital Unix"), _
        Array(7, "OpenVMS"), _
        Array(8, "HPUX"), _
        Array(9, "AIX"), _
        Array(10, "MVS"), _
        Array(11, "OS400"), _
        Array(12, "OS/2"), _
        Array(13, "JavaVM"), _
        Array(14, "MSDOS"), _
        Array(15, "WIN3x"), _
        Array(16, "WIN95"), _
        Array(17, "WIN98"), _
        Array(18, "WINNT"), _
        Array(19, "WINCE"), _
        Array(20, "NCR3000"), _
        Array(21, "NetWare"), _
        Array(22, "OSF"), _
        Array(23, "DC/OS"), _
        Array(24, "Reliant UNIX"), _
        Array(25, "SCO UnixWare"), _
        Array(26, "SCO OpenServer"), _
        Array(27, "Sequent"), _
        Array(28, "IRIX"), _
        Array(29, "Solaris"), _
        Array(30, "SunOS"), _
        Array(31, "U6000"), _
        Array(32, "ASERIES"), _
        Array(33, "TandemNSK"), _
        Array(34, "TandemNT"), _
        Array(35, "BS2000"), _
        Array(36, "LINUX"), _
        Array(37, "Lynx"), _
        Array(38, "XENIX"), _
        Array(39, "VM/ESA"), _
        Array(40, "Interactive UNIX"), _
        Array(41, "BSDUNIX"), _
        Array(42, "FreeBSD"), _
        Array(43, "NetBSD"), _
        Array(44, "GNU Hurd"), _
        Array(45, "OS9"), _
        Array(46, "MACH Kernel"), _
        Array(47, "Inferno"), _
        Array(48, "QNX"), _
        Array(49, "EPOC"), _
        Array(50, "IxWorks"), _
        Array(51, "VxWorks"), _
        Array(52, "MiNT"), _
        Array(53, "BeOS"), _
        Array(54, "HP MPE"), _
        Array(55, "NextStep"), _
        Array(56, "PalmPilot"), _
        Array(57, "Rhapsody"), _
        Array(58, "Windows 2000"), _
        Array(59, "Dedicated"), _
        Array(60, "OS/390"), _
        Array(61, "VSE"), _
        Array(62, "TPF") _
    )
    For Each CurrentOSTypeInfo In aryOSTypeInfos
        If (InputOSTypeCode = CurrentOSTypeInfo(0)) Then
            MappingOSTypeCodeInDescriptionOut = CurrentOSTypeInfo(1)
        End If
    Next
End Function
Function NumberWithTwoDigits(InputNo)
    If (InputNo < 10) Then
        NumberWithTwoDigits = "0" & InputNo
    Else
        NumberWithTwoDigits = InputNo
    End If
End Function
Function ReadLogicalDisk_PortableSoftware(objFSO, objWMIService, strCurDir, ForReading, ForAppending, strResultFileNamePortableSoftware, strResultFileType, strFieldSeparator, strConfigurationFileName, strTableResults)
    If (LCase(strResultFileType) = ".csv") Then
        If (objFSO.FileExists(strCurDir & "\" & strResultFileNamePortableSoftware & strResultFileType)) Then
            bolFilePortableSoftwareHeaderToAdd = False
        Else
            bolFilePortableSoftwareHeaderToAdd = True
        End If
    End If
    aryFieldsPortableSoftware = Array(_
        "Evaluation Timestamp", _
        "Volume Serial Number", _
        "File Name Searched", _
        "Method To Find", _
        "File Path Found", _
        "File Name Found", _
        "File Date Created", _
        "File Date Last Modified", _
        "File Version Found", _
        "File Size Found", _
        "Files Checked For Match Until Found" _
    )
    strFieldsGlued = Join(aryFieldsPortableSoftware, "||")
    Set objResultPortableSoftware = objFSO.OpenTextFile(strCurDir & "\" & strResultFileNamePortableSoftware & strResultFileType, ForAppending, True)
    Select Case LCase(strResultFileType)
        Case ".csv"
            If (bolFilePortableSoftwareHeaderToAdd) Then
                objResultPortableSoftware.WriteLine Join(aryFieldsPortableSoftware, strFieldSeparator)
            End If
        Case ".sql"
            objResultPortableSoftware.WriteLine "DELETE FROM `" & strTableResults & "`;"
    End Select
    Set PortableSoftwareList = objFSO.OpenTextFile(strCurDir & "\" & strConfigurationFileName, ForReading)
    Do Until PortableSoftwareList.AtEndOfStream
        strFileToAnalyze = PortableSoftwareList.ReadLine
        ' will only consider lines with at least 10 characters AND not starting with ' used for keeping comments
        If ((Left(strFileToAnalyze, 1) <> "'") And (Len(strFileToAnalyze) >= 10)) Then
            strFilePieces = Split(strFileToAnalyze, "\")
            If (Right(Left(strFileToAnalyze, 2), 1) = ":") Then
                strDeviceId = Left(strFileToAnalyze, 2)
                crtVolumeSerialNumber = MappingLogicalDiskIdToSerialNumberOrViceVersa(objWMIService, strDeviceId, "Volume Serial Number")
            Else
                crtVolumeSerialNumber = strFilePieces(0)
                strDeviceId = MappingLogicalDiskIdToSerialNumberOrViceVersa(objWMIService, crtVolumeSerialNumber, "Device ID")
            End If
            PieceCounter = 0
            PiecesCounted = UBound(strFilePieces)
            strFilePath = strDeviceId
            For Each strFilePiece In strFilePieces
                If ((PieceCounter < PiecesCounted) And (PieceCounter > 0)) Then
                    strFilePath = strFilePath & "\" & Trim(strFilePiece)
                End If
                PieceCounter = PieceCounter + 1
                strFileNameToSearch = Trim(strFilePiece)
            Next
            If (PiecesCounted = 1) Then
                strFilePath = strFilePath & "\"
            End If
            If ((Len(strDeviceId) > 0) And (Len(strFileNameToSearch) > 0)) Then
                If (objFSO.FolderExists(strFilePath)) Then
                    Set strFolderToSearch = objFSO.GetFolder(strFilePath)
                    intFilesCheckedForMatchUntilFound = 0
                    If (InStr(1, strFileNameToSearch, "*", vbTextCompare) > 0) Then
                        aryFileNamePieces = Split(strFileNameToSearch, "*")
                        If (UBound(aryFileNamePieces) = 1) Then ' only 1 single * is supported
                            RecursiveFileSearchToFileOutput strFolderToSearch, strFileNameToSearch, strResultFileType, objResultPortableSoftware, strFieldsGlued, strFieldSeparator, crtVolumeSerialNumber, intFilesCheckedForMatchUntilFound, strTableResults
                        End If
                    Else
                        RecursiveFileSearchToFileOutput strFolderToSearch, strFileNameToSearch, strResultFileType, objResultPortableSoftware, strFieldsGlued, strFieldSeparator, crtVolumeSerialNumber, intFilesCheckedForMatchUntilFound, strTableResults
                    End If
                End If
            End If
        End If
    Loop
    If (LCase(strResultFileType) = ".sql") Then
        Select Case strTableResults
            Case "in_windows_software_portable"
                ApplySoftwareNormalizationForSoftwarePortable objResultPortableSoftware
            Case "in_windows_security_risk_components"
                ApplySoftwareNormalizationForSecurityRiskComponents objResultPortableSoftware
        End Select
    End If
    objResultPortableSoftware.Close
    PortableSoftwareList.Close
End Function
Function ReadRegistry_SofwareInstalled(strComputer, strResultFileType, strFieldSeparator, objFSO, strCurDir, strResultFileNameSoftware, ForAppending)
    Set objResultSoftware = objFSO.OpenTextFile(strCurDir & "\" & strResultFileNameSoftware & strResultFileType, ForAppending, True)
    If (LCase(strResultFileType) = ".sql") Then
        objResultSoftware.WriteLine "DELETE FROM `in_windows_software_installed` WHERE (`HostName` = '" & strComputer & "');"
    End If
    Set objRegistry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
    objRegistry.GetStringValue HKLM, "SYSTEM\CurrentControlSet\Control\Session Manager\Environment", "PROCESSOR_ARCHITECTURE", strOStype
    CheckSoftware strComputer, bolFileSoftwareHeaderToAdd, objResultSoftware, objRegistry, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"
    ' if Windows is 64-bit an additional Registry Key has to be analyzed
    If (strOStype = "AMD64") Then
        CheckSoftware strComputer, False, objResultSoftware, objRegistry, "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\"
    End If
    If (LCase(strResultFileType) = ".sql") Then
        ApplySoftwareNormalizationForSoftwareInstalled strComputer, objResultSoftware
    End If
    objResultSoftware.Close
End Function
Function ReadWMI__Win32_BaseBoard(objWMIService, strComputer, strResultFileType, strFieldSeparator)
    Dim aryDetailsToReturn(1)
    Dim aryJSONinformationSQL(7)
    aryFieldsBaseBoard = Array(_
        "Caption", _
        "Creation Class Name", _
        "Description", _
        "Manufacturer", _
        "Name", _
        "Status", _
        "Tag", _
        "Version" _
    )
    Set objBaseBoard = objWMIService.ExecQuery("Select * from Win32_BaseBoard")
    For Each crtObjBaseBoard in objBaseBoard
        aryValuesBaseBoard = Array(_
            crtObjBaseBoard.Caption, _
            crtObjBaseBoard.CreationClassName, _
            crtObjBaseBoard.Description, _
            crtObjBaseBoard.Manufacturer, _
            crtObjBaseBoard.Name, _
            crtObjBaseBoard.Status, _
            crtObjBaseBoard.Tag, _
            crtObjBaseBoard.Version _
        )
        Select Case LCase(strResultFileType)
            Case ".csv"
                aryDetailsToReturn(0) = "MBR " & Join(aryFieldsBaseBoard, strFieldSeparator & "MBR ")
                aryDetailsToReturn(1) = AdjustEmptyValueWithinArrayAndGlueIt(aryValuesBaseBoard, "-", strFieldSeparator)
            Case ".sql"
                aryDetailsToReturn(0) = "/* " & strComputer & " - Details MBR results for MySQL */"
                intCounter = 0
                For Each crtField in aryFieldsBaseBoard
                    crtValue = Trim(aryValuesBaseBoard(intCounter))
                    If (IsNull(crtValue)) Then
                        crtValue = "-"
                    End If
                    If (IsNumericExtended(crtValue)) Then
                        aryJSONinformationSQL(intCounter) = """" & crtField & """: " & crtValue
                    Else
                        aryJSONinformationSQL(intCounter) = """" & crtField & """: " & """" & crtValue & """"
                    End If
                    intCounter = intCounter + 1
                Next
                aryDetailsToReturn(1) = Replace(Join(aryJSONinformationSQL, ", "), "\", "\\\\")
        End Select
    Next
    ReadWMI__Win32_BaseBoard = Join(aryDetailsToReturn, "||")
End Function
Function ReadWMI__Win32_BIOS(objWMIService, strComputer, strResultFileType, strFieldSeparator)
    Dim aryDetailsToReturn(1)
    Dim aryJSONinformationSQL(17)
    aryFieldsBIOS = Array(_
        "Build Number", _
        "Caption", _
        "Current Language", _
        "Description", _
        "Identification Code", _
        "Installable Languages", _
        "Manufacturer", _
        "Name", _
        "Primary BIOS", _
        "Serial Number", _
        "SMBIOS BIOS Version", _
        "SMBIOS Major Version", _
        "SMBIOS Minor Version", _
        "SMBIOS Present", _
        "Status", _
        "System Bios Major Version", _
        "System Bios Minor Version", _
        "Version" _
    )
    Set objBIOS = objWMIService.ExecQuery("Select * from Win32_BIOS")
    For Each crtObjBIOS in objBIOS
        strSystemBiosMajorVersion = "N/A"
        strSystemBiosMinorVersion = "N/A"
        ' for Windows 10 and Server 2016 or newer
        If (intOSVersion >= 100) Then
            strSystemBiosMajorVersion = crtObjBIOS.SystemBiosMajorVersion
            strSystemBiosMinorVersion = crtObjBIOS.SystemBiosMinorVersion
        End If
        aryValuesBIOS = Array(_
            crtObjBIOS.BuildNumber, _
            crtObjBIOS.Caption, _
            crtObjBIOS.CurrentLanguage, _
            crtObjBIOS.Description, _
            crtObjBIOS.IdentificationCode, _
            crtObjBIOS.InstallableLanguages, _
            crtObjBIOS.Manufacturer, _
            crtObjBIOS.Name, _
            crtObjBIOS.PrimaryBIOS, _
            crtObjBIOS.SerialNumber, _
            crtObjBIOS.SMBIOSBIOSVersion, _
            crtObjBIOS.SMBIOSMajorVersion, _
            crtObjBIOS.SMBIOSMinorVersion, _
            crtObjBIOS.SMBIOSPresent, _
            crtObjBIOS.Status, _
            strSystemBiosMajorVersion, _
            strSystemBiosMinorVersion, _
            crtObjBIOS.Version _
        )
        Select Case LCase(strResultFileType)
            Case ".csv"
                aryDetailsToReturn(0) = "BIOS " & Join(aryFieldsBIOS, strFieldSeparator & "BIOS ")
                aryDetailsToReturn(1) = AdjustEmptyValueWithinArrayAndGlueIt(aryValuesBIOS, "-", strFieldSeparator)
            Case ".sql"
                aryDetailsToReturn(0) = "/* " & strComputer & " - Details BIOS results for MySQL */"
                intCounter = 0
                For Each crtField in aryFieldsBIOS
                    crtValue = Trim(aryValuesBIOS(intCounter))
                    If (IsNull(crtValue)) Then
                        crtValue = "-"
                    End If
                    If (IsNumericExtended(crtValue)) Then
                        aryJSONinformationSQL(intCounter) = """" & crtField & """: " & crtValue
                    Else
                        aryJSONinformationSQL(intCounter) = """" & crtField & """: " & """" & crtValue & """"
                    End If
                    intCounter = intCounter + 1
                Next
                aryDetailsToReturn(1) = Replace(Join(aryJSONinformationSQL, ", "), "\", "\\\\")
        End Select
    Next
    ReadWMI__Win32_BIOS = Join(aryDetailsToReturn, "||")
End Function
Function ReadWMI__Win32_ComputerSystem(objWMIService, strComputer, strResultFileType, strFieldSeparator)
    Dim aryDetailsToReturn(1)
    Dim aryJSONinformationSQL(29)
    aryFieldsComputerSystem = Array(_
        "Boot Device", _
        "Build Number", _
        "Build Type", _
        "Caption", _
        "Code Set", _
        "Country Code", _
        "Current Time Zone Code", _
        "Current Time Zone Description", _
        "Encryption Level", _
        "Foreground Application Boost", _
        "Install Date", _
        "Locale Code", _
        "Locale Description", _
        "Manufacturer", _
        "Organization", _
        "OS Architecture", _
        "OS Language Code", _
        "OS Language Description", _
        "OS Product Suite", _
        "OS Type Code", _
        "OS Type Description", _
        "Primary", _
        "Registered User", _
        "Serial Number", _
        "System Drive", _
        "System Directory", _
        "Total Virtual Memory [MB]", _
        "Total Visible Memory [MB]", _
        "Version", _
        "Windows Directory" _
    )
    Set objComputerSystem = objWMIService.ExecQuery("Select * from Win32_ComputerSystem")
    Set objOperatingSystem = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
    Set colTimeZone = objWMIService.ExecQuery("Select * from Win32_TimeZone")
    For Each crtObjCS in objComputerSystem
        For Each crtObjOS in objOperatingSystem
            For Each crtObjTZ in colTimeZone
                dtmConvertedDate.Value = crtObjOS.InstallDate
                aryValuesCS = Array(_
                    crtObjOS.BootDevice, _
                    crtObjOS.BuildNumber, _
                    crtObjOS.BuildType, _
                    crtObjOS.Caption, _
                    crtObjOS.CodeSet, _
                    crtObjOS.CountryCode, _
                    crtObjOS.CurrentTimeZone, _
                    crtObjTZ.Description, _
                    crtObjOS.EncryptionLevel, _
                    crtObjOS.ForegroundApplicationBoost, _
                    dtmConvertedDate.GetVarDate, _
                    crtObjOS.Locale, _
                    MappingLanguageLCIDinDescriptionOut("LCID Hexadecimal", crtObjOS.Locale, "Language - Country/Region"), _
                    crtObjOS.Manufacturer, _
                    crtObjOS.Organization, _
                    crtObjOS.OSArchitecture, _
                    crtObjOS.OSLanguage, _
                    MappingLanguageLCIDinDescriptionOut("LCID Decimal", crtObjOS.OSLanguage, "Language - Country/Region"), _
                    crtObjOS.OSProductSuite, _
                    crtObjOS.OSType, _
                    MappingOSTypeCodeInDescriptionOut(crtObjOS.OSType), _
                    crtObjOS.Primary, _
                    crtObjOS.RegisteredUser, _
                    crtObjOS.SerialNumber, _
                    crtObjOS.SystemDrive, _
                    crtObjOS.SystemDirectory, _
                    Round((crtObjOS.TotalVirtualMemorySize / 1024), 0), _
                    Round((crtObjOS.TotalVisibleMemorySize / 1024), 0), _
                    crtObjOS.Version, _
                    crtObjOS.WindowsDirectory _
                )
                Select Case LCase(strResultFileType)
                    Case ".csv"
                        aryDetailsToReturn(0) = Join(aryFieldsComputerSystem, strFieldSeparator)
                        aryDetailsToReturn(1) = AdjustEmptyValueWithinArrayAndGlueIt(aryValuesCS, "-", strFieldSeparator)
                    Case ".sql"
                        aryDetailsToReturn(0) = "/* " & strComputer & " - Details Computer System results for MySQL */"
                        intCounter = 0
                        For Each crtField in aryFieldsComputerSystem
                            crtValue = Trim(aryValuesCS(intCounter))
                            If (IsNull(crtValue)) Then
                                crtValue = "-"
                            End If
                            If (IsNumericExtended(crtValue)) Then
                                aryJSONinformationSQL(intCounter) = """" & crtField & """: " & crtValue
                            Else
                                aryJSONinformationSQL(intCounter) = """" & crtField & """: " & """" & crtValue & """"
                            End If
                            intCounter = intCounter + 1
                        Next
                        aryDetailsToReturn(1) = Replace(Join(aryJSONinformationSQL, ", "), "\", "\\\\")
                End Select
            Next
        Next
    Next
    ReadWMI__Win32_ComputerSystem = Join(aryDetailsToReturn, "||")
End Function
Function ReadWMI__Win32_DiskDrive(objWMIService, strComputer, strResultFileType, strFieldSeparator)
    Dim aryDetailsToReturn(1)
    Dim aryJSONinformationSQL(18)
    aryFieldsDiskDrive = Array(_
        "Bytes PerSector", _
        "Caption", _
        "Description", _
        "FirmwareRevision", _
        "InterfaceType", _
        "Manufacturer", _
        "Model", _
        "Name", _
        "Partitions", _
        "Sectors PerTrack", _
        "Serial Number", _
        "Signature", _
        "Size [GB]", _
        "Status", _
        "Total Cylinders", _
        "Total Heads", _
        "Total Sectors", _
        "Total Tracks", _
        "Tracks Per Cylinder" _
    )
    Set objDiskDrive = objWMIService.ExecQuery("Select * from Win32_DiskDrive")
    aryDetailsToReturn(0) = ""
    aryDetailsToReturn(1) = ""
    For Each crtObjDiskDrive in objDiskDrive
        If (IsNull(crtObjDiskDrive.Signature) Or (crtObjDiskDrive.Signature = "")) Then
            strSignatureSafe = "-"
        Else
            StrSignatureSafe = crtObjDiskDrive.Signature
        End If
        aryValuesDiskDrive = Array(_
            crtObjDiskDrive.BytesPerSector, _
            crtObjDiskDrive.Caption, _
            crtObjDiskDrive.Description, _
            crtObjDiskDrive.FirmwareRevision, _
            crtObjDiskDrive.InterfaceType, _
            crtObjDiskDrive.Manufacturer, _
            crtObjDiskDrive.Model, _
            crtObjDiskDrive.Name, _
            crtObjDiskDrive.Partitions, _
            crtObjDiskDrive.SectorsPerTrack, _
            crtObjDiskDrive.SerialNumber, _
            strSignatureSafe, _
            Round((crtObjDiskDrive.Size /(1024 * 1024 * 1024)), 0), _
            crtObjDiskDrive.Status, _
            crtObjDiskDrive.TotalCylinders, _
            crtObjDiskDrive.TotalHeads, _
            crtObjDiskDrive.TotalSectors, _
            crtObjDiskDrive.TotalTracks, _
            crtObjDiskDrive.TracksPerCylinder _
        )
        strDiskNameCleaned = Replace(Replace(crtObjDiskDrive.Name, "\", ""), ".", "")
        strDiskNumber = Replace(strDiskNameCleaned, "PHYSICALDRIVE", "")
        Select Case LCase(strResultFileType)
            Case ".csv"
                If (aryDetailsToReturn(0) = "") Then
                    aryDetailsToReturn(0) = _
                        "Disk" & strDiskNumber & " " & _
                        Join(aryFieldsDiskDrive, strFieldSeparator & "Disk" & strDiskNumber & " ")
                Else
                    aryDetailsToReturn(0) = aryDetailsToReturn(0) & strFieldSeparator & _
                        "Disk" & strDiskNumber & " " & _
                        Join(aryFieldsDiskDrive, strFieldSeparator & "Disk" & strDiskNumber & " ")
                End If
                If (aryDetailsToReturn(1) = "") Then
                    aryDetailsToReturn(1) = _
                        AdjustEmptyValueWithinArrayAndGlueIt(aryValuesDiskDrive, "-", strFieldSeparator)
                Else
                    aryDetailsToReturn(1) = aryDetailsToReturn(1) & strFieldSeparator & _
                        AdjustEmptyValueWithinArrayAndGlueIt(aryValuesDiskDrive, "-", strFieldSeparator)
                End If
            Case ".sql"
                aryDetailsToReturn(0) = "/* " & strComputer & " - Details Disk results for MySQL */"
                intCounter = 0
                For Each crtField in aryFieldsDiskDrive
                    crtValue = Trim(aryValuesDiskDrive(intCounter))
                    If (IsNull(crtValue)) Then
                        crtValue = "-"
                    End If
                    Select Case crtField
                        Case "Name"
                            aryJSONinformationSQL(intCounter) = """" & crtField & """: " & _
                                """" & Trim(strDiskNameCleaned) & """"
                        Case "Serial Number"
                            If ((StrComp(crtValue, "", 0) <> 0) And (StrComp(crtValue, "", 0) <> 0)) Then
                                aryJSONinformationSQL(intCounter) = """" & crtField & """: " & """" & crtValue & """"
                            Else
                                aryJSONinformationSQL(intCounter) = """" & crtField & """: " & """-"""
                            End If
                        Case Else
                            If (IsNumericExtended(crtValue)) Then
                                aryJSONinformationSQL(intCounter) = """" & crtField & """: " & crtValue
                            Else
                                aryJSONinformationSQL(intCounter) = """" & crtField & """: " & """" & crtValue & """"
                            End If
                    End Select
                    If (crtField = "Name") Then
                    Else
                    End If
                    intCounter = intCounter + 1
                Next
                strDiskNameCleanedNice = Replace(strDiskNameCleaned, "PHYSICALDRIVE", "Physical Drive ")
                If (aryDetailsToReturn(1) = "") Then
                    aryDetailsToReturn(1) = _
                        """" & strDiskNameCleanedNice & """: { " & _
                        Replace(Join(aryJSONinformationSQL, ", "), "\", "\\\\") & _
                        " }"
                Else
                    aryDetailsToReturn(1) = aryDetailsToReturn(1) & ", " & _
                        """" & strDiskNameCleanedNice & """: { " & _
                        Replace(Join(aryJSONinformationSQL, ", "), "\", "\\\\") & _
                        " }"
                End If
        End Select
    Next
    ReadWMI__Win32_DiskDrive = Join(aryDetailsToReturn, "||")
End Function
Function ReadWMI__Win32_LogicalDisk_LightInformationGlued(objWMIService)
    Dim aryDetailsToReturn(1)
    aryDetailsToReturn(0) = ""
    aryDetailsToReturn(1) = ""
    Set objLogicalDisk = objWMIService.ExecQuery("Select * from Win32_LogicalDisk")
    For Each crtObjLogicalDisk in objLogicalDisk
        aryDetailsToReturn(0) = Trim(aryDetailsToReturn(0) & " " & crtObjLogicalDisk.DeviceID)
        aryDetailsToReturn(1) = Trim(aryDetailsToReturn(1) & " " & crtObjLogicalDisk.VolumeSerialNumber)
    Next
    ReadWMI__Win32_LogicalDisk_LightInformationGlued = Join(aryDetailsToReturn, "||")
End Function
Function ReadWMI__Win32_LogicalDisk(objWMIService, strComputer, strResultFileType, strFieldSeparator, ReportFile, bolFileDeviceVolumeHeaderToAdd)
    Dim aryDetailsToReturn(1)
    Dim aryJSONinformationSQL(39)
    aryFieldsLogicalDiskMain = Array(_
        "Volume Serial Number", _
        "Detailed Information" _
    )
    aryFieldsLogicalDisk = Array(_
        "Access", _
        "Availability", _
        "Block Size", _
        "Caption", _
        "Compressed", _
        "Config Manager Error Code", _
        "Config Manager User Config", _
        "Creation Class Name", _
        "Description", _
        "Device ID", _
        "Drive Type Code", _
        "Drive Type Description", _
        "Error Cleared", _
        "Error Description", _
        "Error Methodology", _
        "File System", _
        "Free Space", _
        "Install Date", _
        "Last Error Code", _
        "Maximum Component Length", _
        "Media Type Code", _
        "Media Type Description", _
        "Name", _
        "Number Of Blocks", _
        "PNP Device ID", _
        "Power Management Supported", _
        "Provider Name", _
        "Purpose", _
        "Quotas Disabled", _
        "Quotas Incomplete", _
        "Quotas Rebuilding", _
        "Size", _
        "Status", _
        "Status Info", _
        "Supports Disk Quotas", _
        "Supports File Based Compression", _
        "System Name", _
        "Volume Dirty", _
        "Volume Name", _
        "Volume Serial Number" _
    )
    Set objLogicalDisk = objWMIService.ExecQuery("Select * from Win32_LogicalDisk")
    aryDetailsToReturn(0) = ""
    aryDetailsToReturn(1) = ""
    For Each crtObjLogicalDisk in objLogicalDisk
        aryValuesLogicalDisk = Array(_
            crtObjLogicalDisk.Access, _
            crtObjLogicalDisk.Availability, _
            crtObjLogicalDisk.BlockSize, _
            crtObjLogicalDisk.Caption, _
            crtObjLogicalDisk.Compressed, _
            crtObjLogicalDisk.ConfigManagerErrorCode, _
            crtObjLogicalDisk.ConfigManagerUserConfig, _
            crtObjLogicalDisk.CreationClassName, _
            crtObjLogicalDisk.Description, _
            crtObjLogicalDisk.DeviceID, _
            crtObjLogicalDisk.DriveType, _
            MappingDriveTypeCodeInDescriptionOut(crtObjLogicalDisk.DriveType), _
            crtObjLogicalDisk.ErrorCleared, _
            crtObjLogicalDisk.ErrorDescription, _
            crtObjLogicalDisk.ErrorMethodology, _
            crtObjLogicalDisk.FileSystem, _
            crtObjLogicalDisk.FreeSpace, _
            crtObjLogicalDisk.InstallDate, _
            crtObjLogicalDisk.LastErrorCode, _
            crtObjLogicalDisk.MaximumComponentLength, _
            crtObjLogicalDisk.MediaType, _
            MappingMediaTypeCodeInDescriptionOut(crtObjLogicalDisk.MediaType), _
            crtObjLogicalDisk.Name, _
            crtObjLogicalDisk.NumberOfBlocks, _
            crtObjLogicalDisk.PNPDeviceID, _
            crtObjLogicalDisk.PowerManagementSupported, _
            crtObjLogicalDisk.ProviderName, _
            crtObjLogicalDisk.Purpose, _
            crtObjLogicalDisk.QuotasDisabled, _
            crtObjLogicalDisk.QuotasIncomplete, _
            crtObjLogicalDisk.QuotasRebuilding, _
            crtObjLogicalDisk.Size, _
            crtObjLogicalDisk.Status, _
            crtObjLogicalDisk.StatusInfo, _
            crtObjLogicalDisk.SupportsDiskQuotas, _
            crtObjLogicalDisk.SupportsFileBasedCompression, _
            crtObjLogicalDisk.SystemName, _
            crtObjLogicalDisk.VolumeDirty, _
            crtObjLogicalDisk.VolumeName, _
            crtObjLogicalDisk.VolumeSerialNumber _
        )
        strDiskNameCleaned = crtObjLogicalDisk.VolumeSerialNumber
        Select Case LCase(strResultFileType)
            Case ".csv"
                If (bolFileDeviceVolumeHeaderToAdd) Then
                    ReportFile.WriteLine Join(aryFieldsLogicalDisk, strFieldSeparator)
                    bolFileDeviceVolumeHeaderToAdd = False
                End If
                ReportFile.WriteLine AdjustEmptyValueWithinArrayAndGlueIt(aryValuesLogicalDisk, "-", strFieldSeparator)
            Case ".sql"
                intCounter = 0
                For Each crtField in aryFieldsLogicalDisk
                    crtValue = Trim(aryValuesLogicalDisk(intCounter))
                    Select Case crtField
                        Case "Volume Serial Number"
                            aryJSONinformationSQL(intCounter) = """" & crtField & """: " & """" & crtValue & """"
                        Case Else
                            If (IsNull(crtValue)) Then
                                crtValue = "-"
                            End If
                            If (IsNumericExtended(crtValue)) Then
                                aryJSONinformationSQL(intCounter) = """" & crtField & """: " & crtValue
                            Else
                                aryJSONinformationSQL(intCounter) = """" & crtField & """: " & """" & crtValue & """"
                            End If
                    End Select
                    intCounter = intCounter + 1
                Next
                aryValuesToExpose = Null
                aryValuesToExpose = Array(_
                    crtObjLogicalDisk.VolumeSerialNumber, _
                    "{ " & Replace(Join(aryJSONinformationSQL, ", "), "\", "\\\\") & " }" _
                )
                ReportFile.WriteLine "ALTER TABLE `device_volumes` AUTO_INCREMENT = 1; /* Ensure no end gaps are present in the auto incrementing sequence for Device Volumes */"
                ReportFile.WriteLine "INSERT INTO `device_volumes` (" & _
                    BuildInsertOrUpdateSQLstructure(aryFieldsLogicalDiskMain, aryValuesToExpose, "InsertFields", 0, 0) & _
                    ") VALUES(" & _
                    BuildInsertOrUpdateSQLstructure(aryFieldsLogicalDiskMain, aryValuesToExpose, "InsertValues", 0, 0) & _
                    ") ON DUPLICATE KEY UPDATE " & _
                    BuildInsertOrUpdateSQLstructure(aryFieldsLogicalDiskMain, aryValuesToExpose, "Update", 1, 0) & _
                    ";"
                ReportFile.WriteLine "ALTER TABLE `device_volumes` AUTO_INCREMENT = 1; /* Ensure no end gaps are present in the auto incrementing sequence for Device Volumes */"
        End Select
    Next
End Function
Function ReadWMI__Win32_PhysicalMemoryArray(objWMIService, strComputer, strResultFileType, strFieldSeparator)
    Dim aryDetailsToReturn(1)
    Dim aryJSONinformationSQL(18)
    aryFieldsPMA = Array(_
        "Caption", _
        "Depth", _
        "Description", _
        "Height", _
        "Hot Swappable", _
        "Install Date", _
        "Manufacturer", _
        "Max Capacity", _
        "Memory Devices", _
        "Memory Error Correction", _
        "Model", _
        "Name", _
        "Other Identifying Info", _
        "Serial Number", _
        "SKU", _
        "Status", _
        "Version", _
        "Weight", _
        "Width" _
    )
    ' only required to be able to differentiate a few attributes present only in modern OS versions
    intOSVersion = CurrentOperatingSystemVersionForComparison()
    ' actual Win32_Processor determination
    Set objPMA = objWMIService.ExecQuery("Select * from Win32_PhysicalMemoryArray")
    For Each crtObjPMA in objPMA
        strMaxCapacityEx = "N/A"
        ' for Windows 7 and Server 2012 or newer
        If (intOSVersion >= 61) Then
            strMaxCapacityEx = crtObjPMA.MaxCapacityEx
        End If
        aryValuesPMA = Array(_
            crtObjPMA.Caption, _
            crtObjPMA.Depth, _
            crtObjPMA.Description, _
            crtObjPMA.Height, _
            crtObjPMA.HotSwappable, _
            crtObjPMA.InstallDate, _
            crtObjPMA.Manufacturer, _
            strMaxCapacityEx, _
            crtObjPMA.MemoryDevices, _
            crtObjPMA.MemoryErrorCorrection, _
            crtObjPMA.Model, _
            crtObjPMA.Name, _
            crtObjPMA.OtherIdentifyingInfo, _
            crtObjPMA.SerialNumber, _
            crtObjPMA.SKU, _
            crtObjPMA.Status, _
            crtObjPMA.Version, _
            crtObjPMA.Weight, _
            crtObjPMA.Width _
        )
        Select Case LCase(strResultFileType)
            Case ".csv"
                aryDetailsToReturn(0) = "RAM " & Join(aryFieldsPMA, strFieldSeparator & "RAM ")
                aryDetailsToReturn(1) = AdjustEmptyValueWithinArrayAndGlueIt(aryValuesPMA, "-", strFieldSeparator)
            Case ".sql"
                aryDetailsToReturn(0) = "/* " & strComputer & " - Details RAM details for MySQL */"
                intCounter = 0
                For Each crtField in aryFieldsPMA
                    crtValue = Trim(aryValuesPMA(intCounter))
                    If (IsNull(crtValue)) Then
                        crtValue = "-"
                    End If
                    If (IsNumericExtended(crtValue)) Then
                        aryJSONinformationSQL(intCounter) = """" & crtField & """: " & crtValue
                    Else
                        aryJSONinformationSQL(intCounter) = """" & crtField & """: " & """" & crtValue & """"
                    End If
                    intCounter = intCounter + 1
                Next
                aryDetailsToReturn(1) = Replace(Join(aryJSONinformationSQL, ", "), "\", "\\\\")
        End Select
    Next
    ReadWMI__Win32_PhysicalMemoryArray = Join(aryDetailsToReturn, "||")
End Function
Function ReadWMI__Win32_Processor(objWMIService, strComputer, strResultFileType, strFieldSeparator)
    Dim aryDetailsToReturn(1)
    Dim aryJSONinformationSQL(34)
    aryFieldsCPU = Array(_
        "Address Width", _
        "Architecture", _
        "Availability", _
        "Characteristics", _
        "CPU Status", _
        "Current Clock Speed", _
        "Current Voltage", _
        "Data Width", _
        "Description", _
        "Device ID", _
        "External Clock", _
        "Family", _
        "L2 Cache Size", _
        "L3 Cache Size", _
        "Level", _
        "Load Percentage", _
        "Manufacturer", _
        "Maximum Clock Speed", _
        "Name", _
        "Number Of Cores", _
        "Number Of Enabled Core", _
        "Number Of Logical Processors", _
        "PartNumber", _
        "Processor ID", _
        "Processor Type", _
        "Revision", _
        "Role", _
        "Second Level Address Translation Extensions", _
        "Serial Number", _
        "Socket Designation", _
        "Status Information", _
        "ThreadCount", _
        "Upgrade Method", _
        "Virtualization Firmware Enabled", _
        "VMMonitor Mode Extensions" _
    )
    ' only required to be able to differentiate a few attributes present only in modern OS versions
    intOSVersion = CurrentOperatingSystemVersionForComparison()
    ' actual Win32_Processor determination
    Set objCPU = objWMIService.ExecQuery("Select * from Win32_Processor")
    For Each crtObjCPU in objCPU
        strSecondLevelAddressTranslationExtensions = "N/A"
        strVirtualizationFirmwareEnabled = "N/A"
        strVMMonitorModeExtensions = "N/A"
        strCharacteristics = "N/A"
        strNumberOfEnabledCore = "N/A"
        strPartNumber = "N/A"
        strThreadCount = "N/A"
        strSerialNumber = "N/A"
        ' for Windows 8 and Server 2012 R2 or newer
        If (intOSVersion >= 62) Then
            strSecondLevelAddressTranslationExtensions = crtObjCPU.SecondLevelAddressTranslationExtensions
            strVirtualizationFirmwareEnabled = crtObjCPU.VirtualizationFirmwareEnabled
            strVMMonitorModeExtensions = crtObjCPU.VMMonitorModeExtensions
            ' for Windows 10 and Server 2016 or newer
            If (intOSVersion >= 100) Then
                strCharacteristics = crtObjCPU.Characteristics
                strNumberOfEnabledCore = crtObjCPU.NumberOfEnabledCore
                strPartNumber = crtObjCPU.PartNumber
                strThreadCount = crtObjCPU.ThreadCount
                strSerialNumber = crtObjCPU.SerialNumber
            End If
        End If
        aryValuesCPU = Array(_
            crtObjCPU.AddressWidth, _
            crtObjCPU.Architecture, _
            crtObjCPU.Availability, _
            strCharacteristics, _
            crtObjCPU.CpuStatus, _
            crtObjCPU.CurrentClockSpeed, _
            crtObjCPU.CurrentVoltage, _
            crtObjCPU.DataWidth, _
            crtObjCPU.Description, _
            crtObjCPU.DeviceID, _
            crtObjCPU.ExtClock, _
            crtObjCPU.Family, _
            crtObjCPU.L2CacheSize, _
            crtObjCPU.L3CacheSize, _
            crtObjCPU.Level, _
            crtObjCPU.LoadPercentage, _
            crtObjCPU.Manufacturer, _
            crtObjCPU.MaxClockSpeed, _
            crtObjCPU.Name, _
            crtObjCPU.NumberOfCores, _
            strNumberOfEnabledCore, _
            crtObjCPU.NumberOfLogicalProcessors, _
            strPartNumber, _
            crtObjCPU.ProcessorId, _
            crtObjCPU.ProcessorType, _
            crtObjCPU.Revision, _
            crtObjCPU.Role, _
            strSecondLevelAddressTranslationExtensions, _
            strSerialNumber, _
            crtObjCPU.SocketDesignation, _
            crtObjCPU.StatusInfo, _
            strThreadCount, _
            crtObjCPU.UpgradeMethod, _
            strVirtualizationFirmwareEnabled, _
            strVMMonitorModeExtensions _
        )
        Select Case LCase(strResultFileType)
            Case ".csv"
                aryDetailsToReturn(0) = "CPU " & Join(aryFieldsCPU, strFieldSeparator & "CPU ")
                aryDetailsToReturn(1) = AdjustEmptyValueWithinArrayAndGlueIt(aryValuesCPU, "-", strFieldSeparator)
            Case ".sql"
                aryDetailsToReturn(0) = "/* " & strComputer & " - Details CPU results for MySQL */"
                intCounter = 0
                For Each crtField in aryFieldsCPU
                    crtValue = Trim(aryValuesCPU(intCounter))
                    If (IsNull(crtValue)) Then
                        crtValue = "-"
                    End If
                    If (IsNumericExtended(crtValue)) Then
                        aryJSONinformationSQL(intCounter) = """" & crtField & """: " & crtValue
                    Else
                        aryJSONinformationSQL(intCounter) = """" & crtField & """: " & """" & crtValue & """"
                    End If
                    intCounter = intCounter + 1
                Next
                aryDetailsToReturn(1) = Replace(Join(aryJSONinformationSQL, ", "), "\", "\\\\")
        End Select
    Next
    ReadWMI__Win32_Processor = Join(aryDetailsToReturn, "||")
End Function
Function ReadWMI__Win32_VideoController(objWMIService, strComputer, strResultFileType, strFieldSeparator)
    Dim aryDetailsToReturn(1)
    Dim aryJSONinformationSQL(34)
    aryFieldsVideoController = Array(_
        "Adapter Compatibility", _
        "Adapter DAC Type", _
        "Adapter RAM", _
        "Availability", _
        "Caption", _
        "Color Table Entries", _
        "Current Bits PerPixel", _
        "Current Horizontal Resolution", _
        "Current Number Of Colors", _
        "Current Refresh Rate", _
        "Current Scan Mode", _
        "Current Vertical Resolution", _
        "Description", _
        "Device ID", _
        "Dither Type", _
        "Driver Date", _
        "Driver Version", _
        "ICM Intent", _
        "ICM Method", _
        "Inf Filename", _
        "Inf Section", _
        "Install Date", _
        "Installed Display Drivers", _
        "Max Memory Supported [MB]", _
        "Max Refresh Rate", _
        "Min Refresh Rate", _
        "Monochrome", _
        "Number Of Color Planes", _
        "Number Of Video Pages", _
        "Power Management Supported", _
        "Specification Version", _
        "Video Architecture", _
        "Video Memory Type", _
        "Video Mode Description", _
        "Video Processor" _
    )
    aryDetailsToReturn(0) = ""
    aryDetailsToReturn(1) = ""
    intVideoController = 0
    Set objVideoController = objWMIService.ExecQuery("Select * from Win32_VideoController")
    For Each crtObjVideoController in objVideoController
        If (IsNull(crtObjVideoController.MaxMemorySupported)) Then
            strMaxMemorySupported = 0
        Else
            strMaxMemorySupported = Round(crtObjVideoController.MaxMemorySupported / 1024 / 1024, 0)
        End If
        aryValuesVideoController = Array(_
            crtObjVideoController.AdapterCompatibility, _
            crtObjVideoController.AdapterDACType, _
            crtObjVideoController.AdapterRAM, _
            crtObjVideoController.Availability, _
            crtObjVideoController.Caption, _
            crtObjVideoController.ColorTableEntries, _
            crtObjVideoController.CurrentBitsPerPixel, _
            crtObjVideoController.CurrentHorizontalResolution, _
            crtObjVideoController.CurrentNumberOfColors, _
            crtObjVideoController.CurrentRefreshRate, _
            crtObjVideoController.CurrentScanMode, _
            crtObjVideoController.CurrentVerticalResolution, _
            crtObjVideoController.Description, _
            crtObjVideoController.DeviceID, _
            crtObjVideoController.DitherType, _
            crtObjVideoController.DriverDate, _
            crtObjVideoController.DriverVersion, _
            crtObjVideoController.ICMIntent, _
            crtObjVideoController.ICMMethod, _
            crtObjVideoController.InfFilename, _
            crtObjVideoController.InfSection, _
            crtObjVideoController.InstallDate, _
            crtObjVideoController.InstalledDisplayDrivers, _
            strMaxMemorySupported, _
            crtObjVideoController.MaxRefreshRate, _
            crtObjVideoController.MinRefreshRate, _
            crtObjVideoController.Monochrome, _
            crtObjVideoController.NumberOfColorPlanes, _
            crtObjVideoController.NumberOfVideoPages, _
            crtObjVideoController.PowerManagementSupported, _
            crtObjVideoController.SpecificationVersion, _
            crtObjVideoController.VideoArchitecture, _
            crtObjVideoController.VideoMemoryType, _
            crtObjVideoController.VideoModeDescription, _
            crtObjVideoController.VideoProcessor _
        )
        Select Case LCase(strResultFileType)
            Case ".csv"
                If (aryDetailsToReturn(0) = "") Then
                    aryDetailsToReturn(0) = _
                        "Video " & Join(aryFieldsVideoController, strFieldSeparator & "Video ")
                    aryDetailsToReturn(1) = _
                        AdjustEmptyValueWithinArrayAndGlueIt(aryValuesVideoController, "-", strFieldSeparator)
                Else
                    aryDetailsToReturn(0) = aryDetailsToReturn(0) & strFieldSeparator & _
                        "Video " & Join(aryFieldsVideoController, strFieldSeparator & "Video ")
                    aryDetailsToReturn(1) = aryDetailsToReturn(1) & strFieldSeparator & _
                        AdjustEmptyValueWithinArrayAndGlueIt(aryValuesVideoController, "-", strFieldSeparator)
                End If
            Case ".sql"
                aryDetailsToReturn(0) = "/* " & strComputer & " - Details RAM details for MySQL */"
                intCounter = 0
                For Each crtField in aryFieldsVideoController
                    crtValue = Trim(aryValuesVideoController(intCounter))
                    If (IsNull(crtValue)) Then
                        crtValue = "-"
                    End If
                    If (IsNumericExtended(crtValue)) Then
                        aryJSONinformationSQL(intCounter) = """" & crtField & """: " & crtValue
                    Else
                        aryJSONinformationSQL(intCounter) = """" & crtField & """: " & """" & crtValue & """"
                    End If
                    intCounter = intCounter + 1
                Next
                If (aryDetailsToReturn(1) = "") Then
                    aryDetailsToReturn(1) = _
                        """Video " & intVideoController & """: { " & _
                        Replace(Join(aryJSONinformationSQL, ", "), "\", "\\\\") & _
                        " }"
                Else
                    aryDetailsToReturn(1) = aryDetailsToReturn(1) & ", " & _
                        """Video " & intVideoController & """: { " & _
                        Replace(Join(aryJSONinformationSQL, ", "), "\", "\\\\") & _
                        " }"
                End If
        End Select
        intVideoController = intVideoController + 1
    Next
    ReadWMI__Win32_VideoController = Join(aryDetailsToReturn, "||")
End Function
Function ReadWMI_All(objWMIService, strComputer, strResultFileType, strFieldSeparator, objFSO, strCurDir, strResultFileNameDeviceDetails, ForAppending, bolFileDeviceHeaderToAdd)
    strDetailsCS = Split(ReadWMI__Win32_ComputerSystem(objWMIService, strComputer, strResultFileType, strFieldSeparator), "||")
    strDetailsCPU = Split(ReadWMI__Win32_Processor(objWMIService, strComputer, strResultFileType, strFieldSeparator), "||")
    strDetailsBaseBoard = Split(ReadWMI__Win32_BaseBoard(objWMIService, strComputer, strResultFileType, strFieldSeparator), "||")
    strDetailsBIOS = Split(ReadWMI__Win32_BIOS(objWMIService, strComputer, strResultFileType, strFieldSeparator), "||")
    strDetailsDiskDrive = Split(ReadWMI__Win32_DiskDrive(objWMIService, strComputer, strResultFileType, strFieldSeparator), "||")
    strDetailsRAM = Split(ReadWMI__Win32_PhysicalMemoryArray(objWMIService, strComputer, strResultFileType, strFieldSeparator), "||")
    strDetailsVideoController = Split(ReadWMI__Win32_VideoController(objWMIService, strComputer, strResultFileType, strFieldSeparator), "||")
    Set objResultDeviceDetails = objFSO.OpenTextFile(strCurDir & "\" & strResultFileNameDeviceDetails & strResultFileType, ForAppending, True)
    Select Case LCase(strResultFileType)
        Case ".csv"
            If (bolFileDeviceHeaderToAdd) Then
                objResultDeviceDetails.WriteLine strDetailsCS(0) & _
                    strFieldSeparator & strDetailsCPU(0) & _
                    strFieldSeparator & strDetailsBaseBoard(0) & _
                    strFieldSeparator & strDetailsBIOS(0) & _
                    strFieldSeparator & strDetailsDiskDrive(0) & _
                    strFieldSeparator & strDetailsRAM(0) & _
                    strFieldSeparator & strDetailsVideoController(0)
            End If
            objResultDeviceDetails.WriteLine strDetailsCS(1) & _
                strFieldSeparator & strDetailsCPU(1) & _
                strFieldSeparator & strDetailsBaseBoard(1) & _
                strFieldSeparator & strDetailsBIOS(1) & _
                strFieldSeparator & strDetailsDiskDrive(1) & _
                strFieldSeparator & strDetailsRAM(1) & _
                strFieldSeparator & strDetailsVideoController(1)
        Case ".sql"
            objResultDeviceDetails.WriteLine strDetailsCS(0)
            JSONinformationComputerSystemSQL = "{ " & strDetailsCS(1) & " }"
            objResultDeviceDetails.WriteLine strDetailsCPU(0)
            JSONinformationHardwareSQL = "{ ""CPU"": { " & strDetailsCPU(1) & " }" & _
                ", ""Motherboard"": { " & strDetailsBaseBoard(1) & " }" & _
                ", ""BIOS"": { " & strDetailsBIOS(1) & " }" & _
                ", ""Disk Drive"": { " & strDetailsDiskDrive(1) & " }" & _
                ", ""RAM"": { " & strDetailsRAM(1) & " }" & _
                ", ""Video Controller"": { " & strDetailsVideoController(1) & " }" & _
                " }"
            objResultDeviceDetails.WriteLine "ALTER TABLE `device_details` AUTO_INCREMENT = 1; /* Ensure no end gaps are present in the auto incrementing sequence for Device */"
            objResultDeviceDetails.WriteLine "INSERT INTO `device_details` " & _
                "(`DeviceParrentName`, `DeviceName`, `DeviceOSdetails`, `DeviceHardwareDetails`) " & _
                "VALUES('" & strComputer & "', '" & strComputer & "', " & _
                "'" & JSONinformationComputerSystemSQL & "', " & _
                "'" & JSONinformationHardwareSQL & "') " & _
                "ON DUPLICATE KEY UPDATE " & _
                "`DeviceOSdetails` = '" & JSONinformationComputerSystemSQL & "', " & _
                "`DeviceHardwareDetails` = '" & JSONinformationHardwareSQL & "'" & _
                ";"
            objResultDeviceDetails.WriteLine "ALTER TABLE `device_details` AUTO_INCREMENT = 1; /* Ensure no end gaps are present in the auto incrementing sequence for Device */"
    End Select
    objResultDeviceDetails.Close
End Function
Function ReadWMI_DeviceVolumes(objWMIService, strComputer, strResultFileType, strFieldSeparator, objFSO, strCurDir, strResultFileNameDeviceVolumes, ForAppending)
    If (objFSO.FileExists(strCurDir & "\" & strResultFileNameDeviceVolumes & strResultFileType)) Then
        bolFileDeviceVolumeHeaderToAdd = False
    Else
        bolFileDeviceVolumeHeaderToAdd = True
    End If
    Set objResultDeviceVolumes = objFSO.OpenTextFile(strCurDir & "\" & strResultFileNameDeviceVolumes & strResultFileType, ForAppending, True)
    ReadWMI__Win32_LogicalDisk objWMIService, strComputer, strResultFileType, strFieldSeparator, objResultDeviceVolumes, bolFileDeviceVolumeHeaderToAdd
    If (LCase(strResultFileType) = ".sql") Then
        ApplySoftwareNormalizationForLogicalDisks objResultDeviceVolumes
    End If
    objResultDeviceVolumes.Close
End Function
Function RecursiveFileSearchToFileOutput(strFolderToSearch, strFileNameToSearch, strResultFileType, objResultFile, strFieldsGlued, strFieldSeparator, strVolumeSerialNumber, intFilesCheckedForMatchUntilFound, strTableResults)
    Dim oSubFolder, strCurrentFile
    If (InStr(1, strFileNameToSearch, "*", vbTextCompare) > 0) Then
        aryFileNamePieces = Split(strFileNameToSearch, "*")
        strMethodToFind = "Aproximate"
    Else
        strMethodToFind = "Exact"
    End IF
    aryFields = Split(strFieldsGlued, "||")
    If (FolderHasSubFolders(strFolderToSearch)) Then
        For Each oSubFolder In strFolderToSearch.SubFolders
            RecursiveFileSearchToFileOutput oSubFolder, strFileNameToSearch, strResultFileType, objResultFile, strFieldsGlued, strFieldSeparator, strVolumeSerialNumber, intFilesCheckedForMatchUntilFound, strTableResults
        Next
        For Each strCurrentFile In strFolderToSearch.Files
            strFileToAnalyzeExact = ""
            strCurrentFileCleaned = Replace(Replace(strCurrentFile, strFolderToSearch, ""), "\", "")
            intFilesCheckedForMatchUntilFound = intFilesCheckedForMatchUntilFound + 1
            If (InStr(1, strFileNameToSearch, "*", vbTextCompare) > 0) Then
                If ((InStr(1, strCurrentFileCleaned, aryFileNamePieces(0), vbTextCompare) > 0) And (InStr(1, strCurrentFileCleaned, aryFileNamePieces(1), vbTextCompare) > 0)) Then
                    strFileToAnalyzeExact = strCurrentFileCleaned
                End If
            Else
                If (strCurrentFileCleaned = strFileNameToSearch) Then
                    strFileToAnalyzeExact = strFileNameToSearch
                End If
            End If
            If (strFileToAnalyzeExact <> "") Then
                intFilesCheckedForMatchUntilFound = intFilesCheckedForMatchUntilFound - 1
                aryValues = Array(_
                    CurrentDateTimeToSqlFormat(), _
                    strVolumeSerialNumber, _
                    strFileNameToSearch, _
                    strMethodToFind, _
                    strFolderToSearch, _
                    strCurrentFileCleaned, _
                    ConvertDateTimeToSqlFormat(CDate(strCurrentFile.DateCreated)), _
                    ConvertDateTimeToSqlFormat(CDate(strCurrentFile.DateLastModified)), _
                    strVersionPrefix & objFSO.GetFileVersion(strCurrentFile), _
                    strCurrentFile.Size, _
                    intFilesCheckedForMatchUntilFound _
                )
                Select Case LCase(strResultFileType)
                    Case ".csv"
                        objResultFile.WriteLine AdjustEmptyValueWithinArrayAndGlueIt(aryValues, "-", strFieldSeparator)
                    Case ".sql"
                        objResultFile.WriteLine "INSERT INTO `" & strTableResults & "` (" & _
                            BuildInsertOrUpdateSQLstructure(aryFields, aryValues, "InsertFields", 0, 0) & _
                            ") VALUES(" & _
                            BuildInsertOrUpdateSQLstructure(aryFields, aryValues, "InsertValues", 0, 0) & _
                            ");"
                End Select
                intFilesCheckedForMatchUntilFound = 0
            End If
        Next
    End If
End Function
'-----------------------------------------------------------------------------------------------------------------------
