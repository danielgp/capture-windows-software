# Capture-Windows-Software
## About
Visual Basic Script to capture all the software installed/portable in Windows
 as well as OS, hardware and logical disk details and save results 
 in either CSV format or SQL (to be used w/ MySQL, see below details)

## Supported environments and Testing
Following Operating systems are targeted as supported environments:

Home | Business
---- | --------
:sun_with_face: Microsoft Windows 10 | :partly_sunny: Microsoft Windows Server 2016
:partly_sunny: Microsoft Windows 8.1 | :partly_sunny: Microsoft Windows Server 2012 R2
:partly_sunny: Microsoft Windows 8 | :partly_sunny: Microsoft Windows Server 2012
:first_quarter_moon: Microsoft Windows 7 | :first_quarter_moon: Microsoft Windows Server 2008 R2
:partly_sunny: Microsoft Windows Vista | :partly_sunny: Microsoft Windows Server 2008
:partly_sunny: Microsoft Windows XP | :partly_sunny: Microsoft Windows Server 2003

where above used emoticons stands for:
* :sun_with_face: frequent testing
* :first_quarter_moon: seldom testing
* :new_moon_with_face: rare testing
* :partly_sunny: never tested *(as did not have an environment to test upon)*

## SQL results compatibility for MySQL versions
* :white_check_mark: MySQL Server 5.7.x
* :no_entry: MySQL Server 5.6.x through 3.23.x

## MySQL EER schema for results storage and traceability
![Capture-Windows-Software - MySQL EER schema](https://github.com/danielgp/capture-windows-software/blob/master/MySQL/CaptureWindowsSoftware-EER_Diagram.svg)
