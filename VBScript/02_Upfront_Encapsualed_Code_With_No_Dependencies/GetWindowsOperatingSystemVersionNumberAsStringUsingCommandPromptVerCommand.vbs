Function GetWindowsOperatingSystemVersionNumberAsStringUsingCommandPromptVerCommand(ByRef strOperatingSystemVersion)
    'region FunctionMetadata ####################################################
    ' Safely obtains the operating system version number from the Command Prompt using the
    ' "ver" command
    '
    ' Function takes one positional arguments (strOperatingSystemVersion), which will be
    '       populated with the operating system version in string format upon success
    '
    ' The function returns 0 or a positive number if the operating system version number was
    ' retrieved successfully. A negative number is returned if the operating system version
    ' number was not retrieved successfully.
    '
    ' Example:
    '   intReturnCode = GetWindowsOperatingSystemVersionNumberAsStringUsingCommandPromptVerCommand(strOperatingSystemVersion)
    '   If intReturnCode = 0 Then
    '       ' strOperatingSystemVersion is populated with the operating system version number
    '       ' in string format.
    '   Else
    '       ' The operating system version number could not be retrieved via the Command
    '       ' Prompt's "ver" command.
    '   End If
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
    ' at https://github.com/franklesniak/VBScript_Resources
    'endregion DownloadLocationNotice ####################################################

    'region DependsOn ####################################################
    ' TestObjectForData()
    ' GetCommandPromptPath()
    ' GetTempFilePath()
    'endregion DependsOn ####################################################

    Dim intFunctionReturn
    Dim intReturnMultiplier
    Dim intReturnCode
    Dim strCommandPromptPath
    Dim strTempFilePath
    Dim objWSHShell
    Dim objFileSystemObject
    Dim strWorkingOperatingSystemVersion
    Dim objTextStreamTempFile
    Dim boolFoundLine
    Dim strLine
    Dim arrLine
    Dim intCounter
    Dim arrLinePortion
    Dim arrLinePortion2
    Dim boolFoundVersionNumber

    Const forReading = 1

    Err.Clear

    intFunctionReturn = 0
    intReturnMultiplier = 1

    intReturnCode = GetCommandPromptPath(strCommandPromptPath)
    If intReturnCode < 0 Then
        intFunctionReturn = intFunctionReturn + (intReturnMultiplier * intReturnCode)
    Else
        intReturnMultiplier = intReturnMultiplier * 8
        intReturnMultiplier = intReturnMultiplier * 8
        intReturnMultiplier = intReturnMultiplier * 8
        intReturnCode = GetTempFilePath(strTempFilePath)
        If intReturnCode <> 0 Then
            intFunctionReturn = intFunctionReturn + (intReturnMultiplier * intReturnCode)
        Else
            intReturnMultiplier = intReturnMultiplier * 8
            intReturnMultiplier = intReturnMultiplier * 8
            On Error Resume Next
            Set objWSHShell = CreateObject("WScript.Shell")
            If Err Then
                On Error Goto 0
                Err.Clear
                intFunctionReturn = intFunctionReturn + (intReturnMultiplier * -1)
            Else
                Set objFileSystemObject = CreateObject("Scripting.FileSystemObject")
                If Err Then
                    On Error Goto 0
                    Err.Clear
                    intFunctionReturn = intFunctionReturn + (intReturnMultiplier * -2)
                Else
                    intReturnCode = objWSHShell.Run("""" & strCommandPromptPath & """ /c ""ver > """ & strTempFilePath & """""", 0, True)
                    If Err Then
                        On Error Goto 0
                        Err.Clear
                        intFunctionReturn = intFunctionReturn + (intReturnMultiplier * -3)
                    Else
                        On Error Goto 0
                    End If
                End If
            End If
        End If
    End If

    If intFunctionReturn = 0 Then
        ' No errors have occurred
        ' The file at strTempFilePath is populated with the output of the "ver" command
        ' objFileSystemObject was created successfully
        On Error Resume Next
        Set objTextStreamTempFile = objFileSystemObject.OpenTextFile(strTempFilePath, forReading, False)
        If Err Then
            Err.Clear
            WScript.Sleep(100*Rnd())
            Set objTextStreamTempFile = objFileSystemObject.OpenTextFile(strTempFilePath, forReading, False)
            If Err Then
                Err.Clear
                WScript.Sleep(200*Rnd())
                Set objTextStreamTempFile = objFileSystemObject.OpenTextFile(strTempFilePath, forReading, False)
                If Err Then
                    Err.Clear
                    WScript.Sleep(800*Rnd())
                    Set objTextStreamTempFile = objFileSystemObject.OpenTextFile(strTempFilePath, forReading, False)
                    If Err Then
                        On Error Goto 0
                        Err.Clear
                        intFunctionReturn = intFunctionReturn + (intReturnMultiplier * -4)
                    Else
                        On Error Goto 0
                    End If
                Else
                    On Error Goto 0
                End If
            Else
                On Error Goto 0
            End If
        Else
            On Error Goto 0
        End If
    End If

    If intFunctionReturn = 0 Then
        ' No errors have occurred
        ' The file at strTempFilePath is populated with the output of the "ver" command
        ' objFileSystemObject was created successfully
        ' The temp file is open for reading using objTextStreamTempFile
        boolFoundLine = False
        On Error Resume Next
        Do Until ((objTextStreamTempFile.AtEndOfStream) Or (boolFoundLine = True))
            strLine = objTextStreamTempFile.ReadLine
            If TestObjectForData(strLine) = True Then
                arrLine = Split(strLine, " ")
                If UBound(arrLine) > 0 Then
                    boolFoundLine = True
                End If
            End If
        Loop
        If Err Then
            On Error Goto 0
            Err.Clear
            intFunctionReturn = intFunctionReturn + (intReturnMultiplier * -5)
        Else
            On Error Goto 0
            If boolFoundLine = False Then
                intFunctionReturn = intFunctionReturn + (intReturnMultiplier * -6)
            Else
                'arrLine is already a split of strLine by space
                boolFoundVersionNumber = False
                For intCounter = 0 To UBound(arrLine)
                    arrLinePortion = Split(arrLine(intCounter), ".")
                    If UBound(arrLinePortion) > 0 Then
                        'arrLine(intCounter) contains what appears to be a version number
                        boolFoundVersionNumber = True
                        arrLinePortion2 = Split(arrLine(intCounter), "]")
                        strWorkingOperatingSystemVersion = arrLinePortion2(0)
                        Exit For
                    End If
                Next
                If boolFoundVersionNumber = False Then
                    intFunctionReturn = intFunctionReturn + (intReturnMultiplier * -7)
                Else
                    If TestObjectForData(strWorkingOperatingSystemVersion) = False Then
                        intFunctionReturn = intFunctionReturn + (intReturnMultiplier * -8)
                    End If
                End If
            End If
        End If
        On Error Resume Next
        objTextStreamTempFile.Close
        If Err Then
            On Error Goto 0
            Err.Clear
            ' do not return error
        Else
            Set objTextStreamTempFile = Nothing
            If Err Then
                On Error Goto 0
                Err.Clear
                ' do not return error
            Else
                objFileSystemObject.DeleteFile strTempFilePath, True
                If Err Then
                    On Error Goto 0
                    Err.Clear
                    ' do not return error
                Else
                    On Error Goto 0
                End If
            End If
        End If
    End If

    If intFunctionReturn = 0 Then
        strOperatingSystemVersion = strWorkingOperatingSystemVersion
    End If
    
    GetWindowsOperatingSystemVersionNumberAsStringUsingCommandPromptVerCommand = intFunctionReturn
End Function
