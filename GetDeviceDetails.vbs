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
                            ReportFile.writeline "INSERT INTO `device_details` (`DeviceName`, `DeviceOSdetails`) VALUES(" & _
                                "'" & objComputer.Name & "'" & strFieldSeparatorMySQL & "'" & JSONinformationDeviceOSdetails & _
                                "') ON KEY DUPLICATE UPDATE `DeviceOSdetails` = '" & JSONinformationDeviceOSdetails & _
                                "';"
                    End Select
                Next
            Next
        Next
    Loop 
    ReportFile.Close
End Function 
Function LanguageElementsToIdentify(GivenElement, GivenValue, FeedbackElement)
    aryLanguageCodes = Array(Array("10241", "2801", "Arabic - Syria"), Array("10249", "2809", "English - Belize"), Array("1025", "0401", "Arabic - Saudi Arabia"), Array("10250", "280a", "Spanish - Peru"), Array("10252", "280c", "French - Senegal"), Array("1026", "0402", "Bulgarian"), Array("1027", "0403", "Catalan"), Array("1028", "0404", "Chinese - Taiwan"), Array("1029", "0405", "Czech"), Array("1030", "0406", "Danish"), Array("1031", "0407", "German - Germany"), Array("1032", "0408", "Greek"), Array("1033", "0409", "English - United States"), Array("1034", "040a", "Spanish - Spain Array(Traditional Sort)"), Array("1035", "040b", "Finnish"), Array("1036", "040c", "French - France"), Array("1037", "040d", "Hebrew"), Array("1038", "040e", "Hungarian"), Array("1039", "040f", "Icelandic"), Array("1040", "0410", "Italian - Italy"), Array("1041", "0411", "Japanese"), Array("1042", "0412", "Korean"), Array("1043", "0413", "Dutch - Netherlands"), Array("1044", "0414", "Norwegian Array(Bokmål)"), Array("1045", "0415", "Polish"), Array("1046", "0416", "Portuguese - Brazil"), Array("1047", "0417", "Rhaeto-Romanic"), Array("1048", "0418", "Romanian"), Array("1049", "0419", "Russian"), Array("1050", "041a", "Croatian"), Array("1051", "041b", "Slovak"), Array("1052", "041c", "Albanian - Albania"), Array("1053", "041d", "Swedish"), Array("1054", "041e", "Thai"), Array("1055", "041f", "Turkish"), Array("1056", "0420", "Urdu"), Array("1057", "0421", "Indonesian"), Array("1058", "0422", "Ukrainian"), Array("1059", "0423", "Belarusian"), Array("1060", "0424", "Slovenian"), Array("1061", "0425", "Estonian"), Array("1062", "0426", "Latvian"), Array("1063", "0427", "Lithuanian"), Array("1064", "0428", "Tajik"), Array("1065", "0429", "Farsi"), Array("1066", "042a", "Vietnamese"), Array("1067", "042b", "Armenian - Armenia"), Array("1068", "042c", "Azeri Array(Latin)"), Array("1069", "042d", "Basque"), Array("1070", "042e", "Sorbian"), Array("1071", "042f", "FYRO Macedonian"), Array("1072", "0430", "Sutu"), Array("1073", "0431", "Tsonga"), Array("1074", "0432", "Tswana"), Array("1075", "0433", "Venda"), Array("1076", "0434", "Xhosa"), Array("1077", "0435", "Zulu"), Array("1078", "0436", "Afrikaans - South Africa"), Array("1079", "0437", "Georgian"), Array("1080", "0438", "Faroese"), Array("1081", "0439", "Hindi"), Array("1082", "043a", "Maltese"), Array("1083", "043b", "Sami Array(Lappish)"), Array("1084", "043c", "Scottish Gaelic"), Array("1085", "043d", "Yiddish"), Array("1086", "043e", "Malay - Malaysia"), Array("1087", "043f", "Kazakh"), Array("1088", "0440", "Kyrgyz Array(Cyrillic)"), Array("1089", "0441", "Swahili"), Array("1090", "0442", "Turkmen"), Array("1091", "0443", "Uzbek Array(Latin)"), Array("1092", "0444", "Tatar"), Array("1093", "0445", "Bengali Array(India)"), Array("1094", "0446", "Punjabi"), Array("1095", "0447", "Gujarati"), Array("1096", "0448", "Oriya"), Array("1097", "0449", "Tamil"), Array("1098", "044a", "Telugu"), Array("1099", "044b", "Kannada"), Array("1100", "044c", "Malayalam"), Array("1101", "044d", "Assamese"), Array("1102", "044e", "Marathi"), Array("1103", "044f", "Sanskrit"), Array("1104", "0450", "Mongolian Array(Cyrillic)"), Array("1105", "0451", "Tibetan - People's Republic of China"), Array("1106", "0452", "Welsh"), Array("1107", "0453", "Khmer"), Array("1108", "0454", "Lao"), Array("1109", "0455", "Burmese"), Array("1110", "0456", "Galician"), Array("1111", "0457", "Konkani"), Array("1112", "0458", "Manipuri"), Array("1113", "0459", "Sindhi - India"), Array("1114", "045a", "Syriac"), Array("1115", "045b", "Sinhalese - Sri Lanka"), Array("1116", "045c", "Cherokee - United States"), Array("1117", "045d", "Inuktitut"), Array("1118", "045e", "Amharic - Ethiopia"), Array("1119", "045f", "Tamazight Array(Arabic)"), Array("1120", "0460", "Kashmiri Array(Arabic)"), Array("1121", "0461", "Nepali"), Array("1122", "0462", "Frisian - Netherlands"), Array("1123", "0463", "Pashto"), Array("1124", "0464", "Filipino"), Array("1125", "0465", "Divehi"), Array("1126", "0466", "Edo"), Array("11265", "2c01", "Arabic - Jordan"), Array("1127", "0467", "Fulfulde - Nigeria"), Array("11273", "2c09", "English - Trinidad"), Array("11274", "2c0a", "Spanish - Argentina"), Array("11276", "2c0c", "French - Cameroon"), Array("1128", "0468", "Hausa - Nigeria"), Array("1129", "0469", "Ibibio - Nigeria"), Array("1130", "046a", "Yoruba"), Array("1131", "046B", "Quecha - Bolivia"), Array("1132", "046c", "Sepedi"), Array("1133", "046d", "Bashkir"), Array("1134", "046e", "Luxembourgish"), Array("1135", "046f", "Greenlandic"), Array("1136", "0470", "Igbo - Nigeria"), Array("1137", "0471", "Kanuri - Nigeria"), Array("1138", "0472", "Oromo"), Array("1139", "0473", "Tigrigna - Ethiopia"), Array("1140", "0474", "Guarani - Paraguay"), Array("1141", "0475", "Hawaiian - United States"), Array("1142", "0476", "Latin"), Array("1143", "0477", "Somali"), Array("1144", "0478", "Yi"), Array("1145", "0479", "Papiamentu"), Array("1146", "0471", "Mapudungun"), Array("1148", "047c", "Mohawk"), Array("1150", "047e", "Breton"), Array("1152", "0480", "Uighur - China"), Array("1153", "0481", "Maori - New Zealand"), Array("1154", "0482", "Occitan"), Array("1155", "0483", "Corsican"), Array("1156", "0484", "Alsatian"), Array("1157", "0485", "Yakut"), Array("1158", "0486", "K'iche"), Array("1159", "0487", "Kinyarwanda"), Array("1160", "0488", "Wolof"), Array("1164", "048c", "Dari"), Array("12289", "3001", "Arabic - Lebanon"), Array("12297", "3009", "English - Zimbabwe"), Array("12298", "300a", "Spanish - Ecuador"), Array("12300", "300c", "French - Cote d'Ivoire"), Array("1279", "04ff", "HID Array(Human Interface Device)"), Array("13313", "3401", "Arabic - Kuwait"), Array("13321", "3409", "English - Philippines"), Array("13322", "340a", "Spanish - Chile"), Array("13324", "340c", "French - Mali"), Array("14337", "3801", "Arabic - U.A.E."), Array("14345", "3809", "English - Indonesia"), Array("14346", "380a", "Spanish - Uruguay"), Array("14348", "380c", "French - Morocco"), Array("15361", "3c01", "Arabic - Bahrain"), Array("15369", "3c09", "English - Hong Kong SAR"), Array("15370", "3c0a", "Spanish - Paraguay"), Array("15372", "3c0c", "French - Haiti"), Array("16385", "4001", "Arabic - Qatar"), Array("16393", "4009", "English - India"), Array("16394", "400a", "Spanish - Bolivia"), Array("17417", "4409", "English - Malaysia"), Array("17418", "440a", "Spanish - El Salvador"), Array("18441", "4809", "English - Singapore"), Array("18442", "480a", "Spanish - Honduras"), Array("19466", "4c0a", "Spanish - Nicaragua"), Array("2049", "0801", "Arabic - Iraq"), Array("20490", "500a", "Spanish - Puerto Rico"), Array("2052", "0804", "Chinese - People's Republic of China"), Array("2055", "0807", "German - Switzerland"), Array("2057", "0809", "English - United Kingdom"), Array("2058", "080a", "Spanish - Mexico"), Array("2060", "080c", "French - Belgium"), Array("2064", "0810", "Italian - Switzerland"), Array("2067", "0813", "Dutch - Belgium"), Array("2068", "0814", "Norwegian Array(Nynorsk)"), Array("2070", "0816", "Portuguese - Portugal"), Array("2072", "0818", "Romanian - Moldava"), Array("2073", "0819", "Russian - Moldava"), Array("2074", "081a", "Serbian Array(Latin)"), Array("2077", "081d", "Swedish - Finland"), Array("2080", "0820", "Urdu - India"), Array("2092", "082c", "Azeri Array(Cyrillic)"), Array("2108", "083c", "Irish"), Array("2110", "083e", "Malay - Brunei Darussalam"), Array("2115", "0843", "Uzbek Array(Cyrillic)"), Array("2117", "0845", "Bengali Array(Bangladesh)"), Array("2118", "0846", "Punjabi Array(Pakistan)"), Array("2128", "0850", "Mongolian Array(Mongolian)"), Array("2129", "0851", "Tibetan - Bhutan"), Array("2137", "0859", "Sindhi - Pakistan"), Array("2143", "085f", "Tamazight Array(Latin)"), Array("2144", "0860", "Kashmiri"), Array("2145", "0861", "Nepali - India"), Array("21514", "540a", "Spanish - United States"), Array("2155", "086B", "Quecha - Ecuador"), Array("2163", "0873", "Tigrigna - Eritrea"), Array("22538", "580a", "Spanish - Latin America"), Array("3073", "0c01", "Arabic - Egypt"), Array("3076", "0c04", "Chinese - Hong Kong SAR"), Array("3079", "0c07", "German - Austria"), Array("3081", "0c09", "English - Australia"), Array("3082", "0c0a", "Spanish - Spain Array(Modern Sort)"), Array("3084", "0c0c", "French - Canada"), Array("3098", "0c1a", "Serbian Array(Cyrillic)"), Array("3179", "0C6B", "Quecha - Peru"), Array("4097", "1001", "Arabic - Libya"), Array("4100", "1004", "Chinese - Singapore"), Array("4103", "1007", "German - Luxembourg"), Array("4105", "1009", "English - Canada"), Array("4106", "100a", "Spanish - Guatemala"), Array("4108", "100c", "French - Switzerland"), Array("4122", "101a", "Croatian Array(Bosnia/Herzegovina)"), Array("5121", "1401", "Arabic - Algeria"), Array("5124", "1404", "Chinese - Macao SAR"), Array("5127", "1407", "German - Liechtenstein"), Array("5129", "1409", "English - New Zealand"), Array("5130", "140a", "Spanish - Costa Rica"), Array("5132", "140c", "French - Luxembourg"), Array("5146", "141A", "Bosnian Array(Bosnia/Herzegovina)"), Array("58380", "e40c", "French - North Africa"), Array("6145", "1801", "Arabic - Morocco"), Array("6153", "1809", "English - Ireland"), Array("6154", "180a", "Spanish - Panama"), Array("6156", "180c", "French - Monaco"), Array("7169", "1c01", "Arabic - Tunisia"), Array("7177", "1c09", "English - South Africa"), Array("7178", "1c0a", "Spanish - Dominican Republic"), Array("7180", "1c0c", "French - West Indies"), Array("8193", "2001", "Arabic - Oman"), Array("8201", "2009", "English - Jamaica"), Array("8202", "200a", "Spanish - Venezuela"), Array("8204", "200c", "French - Reunion"), Array("9217", "2401", "Arabic - Yemen"), Array("9225", "2409", "English - Caribbean"), Array("9226", "240a", "Spanish - Colombia"), Array("9228", "240c", "French - Democratic Rep. of Congo"))
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
    aryOSTypeInfos = Array(Array(0, "Unknown"), Array(1, "Other"), Array(10, "MVS"), Array(11, "OS400"), Array(12, "OS/2"), Array(13, "JavaVM"), Array(14, "MSDOS"), Array(15, "WIN3x"), Array(16, "WIN95"), Array(17, "WIN98"), Array(18, "WINNT"), Array(19, "WINCE"), Array(2, "MACOS"), Array(20, "NCR3000"), Array(21, "NetWare"), Array(22, "OSF"), Array(23, "DC/OS"), Array(24, "Reliant UNIX"), Array(25, "SCO UnixWare"), Array(26, "SCO OpenServer"), Array(27, "Sequent"), Array(28, "IRIX"), Array(29, "Solaris"), Array(3, "ATTUNIX"), Array(30, "SunOS"), Array(31, "U6000"), Array(32, "ASERIES"), Array(33, "TandemNSK"), Array(34, "TandemNT"), Array(35, "BS2000"), Array(36, "LINUX"), Array(37, "Lynx"), Array(38, "XENIX"), Array(39, "VM/ESA"), Array(4, "DGUX"), Array(40, "Interactive UNIX"), Array(41, "BSDUNIX"), Array(42, "FreeBSD"), Array(43, "NetBSD"), Array(44, "GNU Hurd"), Array(45, "OS9"), Array(46, "MACH Kernel"), Array(47, "Inferno"), Array(48, "QNX"), Array(49, "EPOC"), Array(5, "DECNT"), Array(50, "IxWorks"), Array(51, "VxWorks"), Array(52, "MiNT"), Array(53, "BeOS"), Array(54, "HP MPE"), Array(55, "NextStep"), Array(56, "PalmPilot"), Array(57, "Rhapsody"), Array(58, "Windows 2000"), Array(59, "Dedicated"), Array(6, "Digital Unix"), Array(60, "OS/390"), Array(61, "VSE"), Array(62, "TPF"), Array(7, "OpenVMS"), Array(8, "HPUX"), Array(9, "AIX"))
    For Each CurrentOSTypeInfo In aryOSTypeInfos
        If (InputOSTypeCode = CurrentOSTypeInfo(0)) Then
            OSTypeDescription = CurrentOSTypeInfo(1)
        End If
    Next
End Function
'-----------------------------------------------------------------------------------------------------------------------
