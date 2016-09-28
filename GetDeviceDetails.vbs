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
Const strResultFileName = "DeviceDetails" 
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set WshShell = WScript.CreateObject("WScript.Shell") 
Set dtmConvertedDate = CreateObject("WbemScripting.SWbemDateTime")
'-----------------------------------------------------------------------------------------------------------------------
MsgBox "This script will read from Windows Management Instrumentation (WMI) current Device Details and export it in a file with a pre-configured name!" & vbNewLine & vbNewLine & "please wait until script is completed...", vbOKOnly + vbInformation, "Start feedback"
InputResultType = MsgBox("This is a script intended to read current Device Details from Windows Management Instrumentation (WMI)!" & vbNewLine & vbNewLine & "Do you want to store obtained results into CSV format file?" & vbNewLine & vbNewLine & "if you choose No a SQL file will be used instead" & vbNewLine & "otherwise choosing Cancel will end current script without any processing and result.", vbYesNoCancel + vbQuestion, "Choose processing result type")
If (InputResultType = vbCancel) Then
    MsgBox "This is a script intended to read current Device Details from Windows Management Instrumentation (WMI)!" & vbNewLine & vbNewLine & "You have chosen to terminate execution without any processing and no result, should you arrive at this point by mistake just re-execute it and pay greater attention to previous options dialogue, otherwise thanks for your attention!", vbOKOnly + vbExclamation, "Script end"
Else
    StartTime = Timer()
    strCurDir = WshShell.CurrentDirectory
    Set SrvListFile = objFSO.OpenTextFile(strCurDir & "\WindowsComputerList.txt", ForReading) 
        ReadWMI__Win32_ComputerSystem()
    SrvListFile.Close
    EndTime = Timer()
    MsgBox "This script has completed read current Device Details from Windows Management Instrumentation (WMI), (in just " & FormatNumber(EndTime - StartTime, 0) & " seconds), please consult generated file [" & strCurDir & "\" & strResultFileName & strResultFileType & "]." & vbNewLine & vbNewLine & "Thank you for using this script, hope to see you back soon!", vbOKOnly + vbInformation, "Script end"
End If
'-----------------------------------------------------------------------------------------------------------------------
Function ReadWMI__Win32_ComputerSystem()
    Select Case InputResultType
        Case vbYes
            strResultFileType = ".csv"
            If (objFSO.FileExists(strCurDir & "\" & strResultFileName & strResultFileType)) Then
                bolFileHeaderToAdd = False
            Else
                bolFileHeaderToAdd = True
            End If
        Case vbNo 
            strResultFileType = ".sql"
    End Select
    Set ReportFile = objFSO.OpenTextFile(strCurDir & "\" & strResultFileName & strResultFileType, ForAppending, True) 
    aryInformationToExpose = Array(_
        "Computer Name", _
        "Boot Device", _
        "Build Number", _
        "Build Type", _
        "Caption", _
        "Code Set", _
        "Country Code", _
        "Current Time Zone", _
        "Encryption Level", _
        "Foreground Application Boost", _
        "Install Date", _
        "Locale", _
        "Manufacturer", _
        "Organization", _
        "OS Architecture", _
        "OS Language", _
        "OS Product Suite", _
        "OS Type", _
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
    Do Until SrvListFile.AtEndOfStream 
        strComputer = LCase(SrvListFile.ReadLine) 
        Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
        Set objComputerSystem = objWMIService.ExecQuery("Select * from Win32_ComputerSystem")
        Set oss = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
        Set colTimeZone = objWMIService.ExecQuery("Select * from Win32_TimeZone")
        For Each objComputer in objComputerSystem
            For Each osObj in oss
                For Each objTZ in colTimeZone
                    dtmConvertedDate.Value = osObj.InstallDate
                    aryValuesToExpose = Array(_
                        objComputer.Name, _
                        osObj.BootDevice, _
                        osObj.BuildNumber, _
                        osObj.BuildType, _
                        osObj.Caption, _
                        osObj.CodeSet, _
                        osObj.CountryCode, _
                        osObj.CurrentTimeZone & " {" & objTZ.Description & "}", _
                        osObj.EncryptionLevel, _
                        osObj.ForegroundApplicationBoost, _
                        dtmConvertedDate.GetVarDate, _
                        osObj.Locale & " (" & LanguageElementsToIdentify("LCID Hexadecimal", osObj.Locale, "Language - Country/Region") & ")", _
                        osObj.Manufacturer, _
                        osObj.Organization, _
                        osObj.OSArchitecture, _
                        osObj.OSLanguage & " (" & LanguageElementsToIdentify("LCID Decimal", osObj.OSLanguage, "Language - Country/Region") & ")", _
                        osObj.OSProductSuite, _
                        osObj.OSType & " (" & OSTypeDescription(osObj.OSType) & ")", _
                        osObj.Primary, _
                        osObj.RegisteredUser, _
                        osObj.SerialNumber, _
                        osObj.SystemDrive, _
                        osObj.SystemDirectory, _
                        Round((osObj.TotalVirtualMemorySize / 1024), 0), _
                        Round((osObj.TotalVisibleMemorySize / 1024), 0), _
                        osObj.Version, _
                        osObj.WindowsDirectory _
                    )
                    Select Case LCase(strResultFileType)
                        Case ".csv"
                            If (bolFileHeaderToAdd) Then
                                ReportFile.writeline Join(aryInformationToExpose, strFieldSeparator)
                            End If
                            ReportFile.writeline Join(aryValuesToExpose, strFieldSeparator)
                        Case ".sql"
                            strFieldSeparatorMySQL = ", "
                            JSONinformationDeviceOSdetails = ""
                            intCounter = 0
                            For Each CurrenInformationToExpose in aryInformationToExpose
                                If (intCounter = 0) Then
                                    JSONinformationDeviceOSdetails = "{ "
                                Else
                                    If (intCounter > 1) Then
                                        JSONinformationDeviceOSdetails = JSONinformationDeviceOSdetails & ", "
                                    End If
                                    JSONinformationDeviceOSdetails = JSONinformationDeviceOSdetails & _
                                        """" & CurrenInformationToExpose & """: " & _
                                        """" & aryValuesToExpose(intCounter) & """"
                                End If
                                intCounter = intCounter + 1
                            Next
                            JSONinformationDeviceOSdetails = JSONinformationDeviceOSdetails & " }"
                            JSONinformationDeviceOSdetails = Replace(JSONinformationDeviceOSdetails, "\", "\\\\")
                            ReportFile.writeline "INSERT INTO `device_details` (`DeviceName`, `DeviceOSdetails`) VALUES(" & _
                                "'" & objComputer.Name & "'" & strFieldSeparatorMySQL & "'" & JSONinformationDeviceOSdetails & _
                                "') ON DUPLICATE KEY UPDATE `DeviceOSdetails` = '" & JSONinformationDeviceOSdetails & _
                                "';"
                    End Select
                Next
            Next
        Next
    Loop 
    If LCase(strResultFileType) = ".sql" Then
        ReportFile.writeline "ALTER TABLE `device_details` AUTO_INCREMENT = 1;"
    End If
    ReportFile.Close
End Function 
Function LanguageElementsToIdentify(GivenElement, GivenValue, FeedbackElement)
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
                            LanguageElementsToIdentify = CurrentLanguageCode(0)
                        End If
                    Case "LCID Hexadecimal"
                        If (CStr(GivenValue) = CurrentLanguageCode(2)) Then
                            LanguageElementsToIdentify = CurrentLanguageCode(1)
                        End If
                End Select
            Case "LCID Decimal"
                Select Case FeedbackElement
                    Case "Language - Country/Region"
                        If (CStr(GivenValue) = CurrentLanguageCode(0)) Then
                            LanguageElementsToIdentify = CurrentLanguageCode(2)
                        End If
                    Case "LCID Hexadecimal"
                        If (CStr(GivenValue) = CurrentLanguageCode(0)) Then
                            LanguageElementsToIdentify = CurrentLanguageCode(1)
                        End If
                End Select
            Case "LCID Hexadecimal"
                Select Case FeedbackElement
                    Case "Language - Country/Region"
                        If (CStr(GivenValue) = CurrentLanguageCode(1)) Then
                            LanguageElementsToIdentify = CurrentLanguageCode(2)
                        End If
                    Case "LCID Decimal"
                        If (CStr(GivenValue) = CurrentLanguageCode(1)) Then
                            LanguageElementsToIdentify = CurrentLanguageCode(0)
                        End If
                End Select
        End Select
    Next
End Function
Function OSTypeDescription(InputOSTypeCode)
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
            OSTypeDescription = CurrentOSTypeInfo(1)
        End If
    Next
End Function
'-----------------------------------------------------------------------------------------------------------------------
