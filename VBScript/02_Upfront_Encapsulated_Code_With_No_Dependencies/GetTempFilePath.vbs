Function GetTempFilePath(ByRef strTempFilePath)
    'region FunctionMetadata ####################################################
    ' Safely obtains the path to the temporary files folder
    '
    ' Function takes one positional argument (strTempFilePath) that is populated upon
    ' success with the path to the temprary files folder. The path is appended with a trailing
    ' backslash.
    '
    ' The function returns 0 if the temporary folder path was retrieved successfully. A
    ' negative number is returned if the temporary folder path was not retrieved successfully.
    '
    ' Example:
    ' intReturnCode = GetTempFilePath(strTempFilePath)
    ' If intReturnCode = 0 Then
    '   ' Temporary folder path was retrieved successfully and stored in strTempFilePath.
    ' End If
    '
    ' Note: the technique used in this function requires Windows Scripting Host 2.0 or newer,
    ' which was included in Windows releases beginning with Windows 2000 and Windows ME. It was
    ' available as a separate download for Windows 95, 98, and NT 4.0.
    '
    ' Note: if the processor architecture of the VBScript process does not match the operating
    ' system's processor architecture (e.g., 32-bit Intel IA32/x86 VBScript process running on
    ' 64-bit AMD64/Intel x86-64 Windows), then the path to the Windows System folder may be
    ' automatically substituted for the Windows-on-Windows (WOW) equivalent path.
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
    'endregion Acknowledgements ####################################################

    'region DependsOn ####################################################
    ' TestObjectForData()
    ' GetTempFolderPath()
    'endregion DependsOn ####################################################

    Dim intFunctionReturn
    Dim intReturnMultiplier
    Dim intReturnCode
    Dim strTempFolderPath
    Dim objFileSystemObject
    Dim strTempFile
    Dim strWorkingTempFilePath

    Err.Clear

    intFunctionReturn = 0
    intReturnMultiplier = 1

    intReturnCode = GetTempFolderPath(strTempFolderPath)
    If intReturnCode <> 0 Then
        intFunctionReturn = intReturnCode
    Else
        intReturnMultiplier = intReturnMultiplier * 8
        On Error Resume Next
        Set objFileSystemObject = CreateObject("Scripting.FileSystemObject")
        If Err Then
            On Error Goto 0
            Err.Clear
            intFunctionReturn = -1 * intReturnMultiplier
        Else
            strTempFile = objFileSystemObject.GetTempName
            If Err Then
                On Error Goto 0
                Err.Clear
                intFunctionReturn = -2 * intReturnMultiplier
            Else
                On Error Goto 0
                If TestObjectForData(strTempFile) = False Then
                    intFunctionReturn = -3 * intReturnMultiplier
                Else
                    strWorkingTempFilePath = strTempFolderPath & strTempFile
                End If
            End If
        End If
    End If

    If intFunctionReturn = 0 Then
        strTempFilePath = strWorkingTempFilePath
    End If
    
    GetTempFilePath = intFunctionReturn
End Function
