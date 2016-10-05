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
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set WshShell = WScript.CreateObject("WScript.Shell") 
strCurDir = WshShell.CurrentDirectory
Set PortableSoftwareList = objFSO.OpenTextFile(strCurDir & "\PortableSoftwareList.txt", ForReading) 
Do Until PortableSoftwareList.AtEndOfStream 
    strFileToAnalyze = PortableSoftwareList.ReadLine
    If (InStr(1, strFileToAnalyze, "*", vbTextCompare)) Then
        strFilePieces = Split(strFileToAnalyze, "\")
        PieceCounter = 0
        PiecesCounted = UBound(strFilePieces)
        For Each strFilePiece In strFilePieces 
            If (PieceCounter < PiecesCounted) Then
                If (PieceCounter = 0) Then
                    strFilePath = strFilePiece
                Else
                    strFilePath = strFilePath & "\" & strFilePiece
                End If
            End If
            PieceCounter = PieceCounter + 1
            strFileNameToSearch = strFilePiece
        Next
        If (objFSO.FolderExists(strFilePath)) Then
            strFileNamePieces = Split(strFileNameToSearch, "*")
            If (UBound(strFileNamePieces) = 1) Then ' only 1 single * is supported
                Set strFolderToSearch = objFSO.GetFolder(strFilePath)
                For Each strCurrentFile In strFolderToSearch.Files
                    If ((InStr(1, strCurrentFile, strFileNamePieces(0), vbTextCompare) > 0) And (InStr(1, strCurrentFile, strFileNamePieces(1), vbTextCompare) > 0)) Then
                        strFileToAnalyzeExact = strCurrentFile
                    End If
                Next
            End If
        End If
    Else
        strFileToAnalyzeExact = strFileToAnalyze
    End If
    If (objFSO.FileExists(strFileToAnalyzeExact)) Then
        MsgBox strFileToAnalyzeExact & vbNewLine & objFSO.GetFileVersion(strFileToAnalyzeExact)
    End If
Loop 