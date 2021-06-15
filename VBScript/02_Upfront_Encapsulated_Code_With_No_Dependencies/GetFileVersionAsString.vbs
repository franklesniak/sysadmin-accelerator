Function GetFileVersionAsString(ByRef strFileVersion, ByVal strFilePath)
    'region FunctionMetadata ####################################################
    ' Safely obtains the file version of a binary file. This is the "file version" displayed in
    ' the properties of the file, details tab, when viewed from Windows Explorer.
    '
    ' Function takes two positional arguments:
    '   The first argument (strFileVersion) will be the string representation of the
    '       file's version (whose path is strFilePath).
    '   The second argument (strFilePath) is the path to the file for which we want to know the
    '       file version.
    '
    ' The function returns 0 if the file's version was retrieved successfully. A negative
    ' number is returned if the file's product version was not retrieved successfully.
    '
    ' Example:
    ' strFilePath = "C:\Windows\System32\hal.dll"
    ' intReturnCode = GetFileVersionAsString(strFileVersion, strFilePath)
    ' If intReturnCode = 0 Then
    '   ' The product version of hal.dll was retrieved successfully and is stored in
    '   ' strFileVersion in string format.
    ' End If
    '
    ' Note: the technique used in this function requires Windows Scripting Host 2.0 or newer,
    ' which was included in Windows releases beginning with Windows 2000 and Windows ME. It was
    ' available as a separate download for Windows 95, 98, and NT 4.0.
    '
    ' Version: 1.0.20210614.0
    'endregion FunctionMetadata ####################################################

    'region License ####################################################
    ' Copyright 2021 Frank Lesniak
    '
    ' Permission is hereby granted, free of charge, to any person obtaining a copy of this
    ' software and associated documentation files (the "Software"), to deal in the Software
    ' without restriction, including without limitation the rights to use, copy, modify, merge,
    ' publish, distribute, sublicense, and/or sell copies of the Software, and to permit
    ' persons to whom the Software is furnished to do so, subject to the following conditions:
    '
    ' The above copyright notice and this permission notice shall be included in all copies or
    ' substantial portions of the Software.
    '
    ' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
    ' INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
    ' PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
    ' FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
    ' OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
    ' DEALINGS IN THE SOFTWARE.
    'endregion License ####################################################

    'region DownloadLocationNotice ####################################################
    ' The most up-to-date version of this script can be found on the author's GitHub repository
    ' at https://github.com/franklesniak/sysadmin-accelerator
    'endregion DownloadLocationNotice ####################################################

    'region Acknowledgements ####################################################
    ' Andrew Clinick, for his article "If It Moves, Script It" (available in the MSDN library
    ' published 2003 Jan), which tipped me off that FileSystemObject is available starting in
    ' Windows Scripting Host 2.0.
    '
    ' Jerry Lee Ford, Jr., for providing a history of VBScript and Windows Scripting Host in
    ' his book, "Microsoft WSH and VBScript Programming for the Absolute Beginner".
    '
    ' Gunter Born, for providing a history of Windows Scripting Host in his book "Microsoft
    ' Windows Script Host 2.0 Developer's Guide" that corrected some points.
    '
    ' The Microsoft "Scripting Guys", who published sample code to obtain a file's regular
    ' "file version" using FileSystemObject:
    ' https://devblogs.microsoft.com/scripting/how-can-i-determine-the-version-number-of-a-file/
    'endregion Acknowledgements ####################################################

    'region DependsOn ####################################################
    ' TestObjectForData()
    'endregion DependsOn ####################################################

    Dim intFunctionReturn
    Dim objFileSystemObject
    Dim boolResult
    Dim strWorkingFileVersion

    Err.Clear

    intFunctionReturn = 0

    On Error Resume Next
    Set objFileSystemObject = CreateObject("Scripting.FileSystemObject")
    If Err Then
        On Error Goto 0
        Err.Clear
        intFunctionReturn = -1
    Else
        boolResult = objFileSystemObject.FileExists(strFilePath)
        If Err Then
            On Error Goto 0
            Err.Clear
            intFunctionReturn = -2
        Else
            On Error Goto 0
            If boolResult = False Then
                intFunctionReturn = -3
            Else
                On Error Resume Next
                strWorkingFileVersion = objFileSystemObject.GetFileVersion(strFilePath)
                If Err Then
                    On Error Goto 0
                    Err.Clear
                    intFunctionReturn = -4
                Else
                    On Error Goto 0
                    If TestObjectForData(strWorkingFileVersion) = False Then
                        intFunctionReturn = -5
                    End If
                End If
            End If
        End If
    End If

    If intFunctionReturn = 0 Then
        strFileVersion = strWorkingFileVersion
    End If
    
    GetFileVersionAsString = intFunctionReturn
End Function
