# Capture-Windows-Software
## About
Visual Basic Script to capture all the software installed/portable in Windows as well as OS, hardware and logical disk details and save results in either CSV format or SQL (to be used w/ MySQL, see below details)

# Features
* Reads Device Details from WMI
* Reads Logical Disk Details from WMI
* Reads all installed applications/drivers/runtimes/libraries/modules (both 32-bit and 64-bit) from Windows Registry
* Extracts versions from inscope portable applications (from whitelisted)
* Scans configured targeted pathes for versions of blacklisted EXE/DLL to assess vulnerabilities or End Of Life versions

## Supported environments and Testing
Following Microsoft Windows operating systems are targeted as supported environments:

Home (Client) | Business (Server) |  Device Details | Logical Disk Details | Installed things | Portable Whitelist Versions | Security Assesment
------------- | ----------------- | --------------- | -------------------- | ---------------- | --------------------------- | ------------------
:sun_with_face: 10 | :partly_sunny: 2016 | :white_check_mark: | :white_check_mark: | :white_check_mark: | :white_check_mark: | :white_check_mark:
:partly_sunny: 8.1 | :partly_sunny: 2012 R2 | :white_check_mark: | :white_check_mark: | :white_check_mark: | :white_check_mark: | :white_check_mark:
:partly_sunny: 8 | :partly_sunny: 2012 | :white_check_mark: | :white_check_mark: | :white_check_mark: | :white_check_mark: | :white_check_mark:
:first_quarter_moon: 7 | :first_quarter_moon: 2008 R2 | :wavy_dash: | :white_check_mark: | :white_check_mark: | :white_check_mark: | :white_check_mark:
:partly_sunny: Vista | :partly_sunny: 2008 | :wavy_dash: | :white_check_mark: | :white_check_mark: | :white_check_mark: | :white_check_mark:
:first_quarter_moon: XP | :partly_sunny: 2003 | :wavy_dash: | :white_check_mark: | :white_check_mark: |:white_check_mark: | :white_check_mark:
:partly_sunny: ME | :first_quarter_moon: 2000 |  :wavy_dash: |  :wavy_dash: | :white_check_mark: |:white_check_mark: | :white_check_mark:

where above used emoticons stands for:
* :sun_with_face: frequent testing
* :first_quarter_moon: seldom testing
* :partly_sunny: never tested _(as did not have an environment to test upon)_
* :white_check_mark: full support
* :wavy_dash: supported with most features (few values will be #N/A)

## SQL results compatibility for MySQL versions
* :white_check_mark: MySQL Server 5.7.x
* :no_entry: MySQL Server 3.23.x through 5.6.x _(no support for generated columns used to automatically determine certain things)_

## MySQL EER schema for results storage and traceability
![Capture-Windows-Software - MySQL EER schema](https://github.com/danielgp/capture-windows-software/blob/master/MySQL/CaptureWindowsSoftware-EER_Diagram.svg)

## Used refferences

Reference name | URL
-------------- | ---
Emoji Cheat Sheet | http://www.webpagefx.com/tools/emoji-cheat-sheet/
GetFileVersion Method | https://msdn.microsoft.com/en-us/library/b4e05k97(v=vs.84).aspx
Mastering Markdown | https://guides.github.com/features/mastering-markdown/
Operating System Version | https://msdn.microsoft.com/en-us/library/windows/desktop/ms724832(v=vs.85).aspx
MySQL Community Downloads | http://dev.mysql.com/downloads/
Uninstall Registry Key | https://msdn.microsoft.com/en-us/library/windows/desktop/aa372105(v=vs.85).aspx
VBScript Features | https://msdn.microsoft.com/en-us/library/273zc69c.aspx
VBScript Reference | https://technet.microsoft.com/en-us/library/ee198844.aspx
VBScript Version Information | https://msdn.microsoft.com/en-us/library/4y5y7bh5.aspx
Win32_BaseBoard class | https://msdn.microsoft.com/en-us/library/aa394072(v=vs.85).aspx
Win32_BIOS class | https://msdn.microsoft.com/en-us/library/aa394077(v=vs.85).aspx
Win32_ComputerSystem class | https://msdn.microsoft.com/en-us/library/aa394102(v=vs.85).aspx
Win32_DiskDrive class | https://msdn.microsoft.com/en-us/library/aa394132(v=vs.85).aspx
Win32_LogicalDisk class | https://msdn.microsoft.com/en-us/library/aa394173(v=vs.85).aspx
Win32_OperatingSystem | https://msdn.microsoft.com/en-us/library/aa394239(v=vs.85).aspx
Win32_PhysicalMemoryArray class | https://msdn.microsoft.com/en-us/library/aa394348(v=vs.85).aspx
Win32_Processor class | https://msdn.microsoft.com/en-us/library/aa394373(v=vs.85).aspx
Win32_VideoController class | https://msdn.microsoft.com/en-us/library/aa394512(v=vs.85).aspx
