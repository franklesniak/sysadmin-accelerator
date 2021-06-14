Function GetCommandPromptPath(ByRef strCommandPromptPath)
    'region FunctionMetadata ####################################################
    ' Safely determines the path to the Windows Command Prompt (command interpreter)
    '
    ' This function takes one argument (strCommandPromptPath) that is populated upon success
    ' with the path to the Command Prompt executable.
    '
    ' The function returns 0 or a positive number if the path to the Command Prompt was
    ' retrieved successfully; it returns a negative number if the path to the Command Prompt
    ' was not retrived successfully.
    '
    ' Example:
    ' intReturnCode = GetCommandPromptPath(strCommandPromptPath)
    ' If intReturnCode = 0 Then
    '   ' Path to command prompt executable was retrieved successfully and stored in
    '   ' strCommandPromptPath.
    ' End If
    '
    ' Version: 1.0.20210613.0
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

    'region Acknowledgements ####################################################
    ' Microsoft, for including in the MSDN Library Jan 2003 information on the nuiances in
    ' accessing environment variables on pre-Windows 2000 and Windows ME-and-prior operating
    ' systems (namely that VBScript in Windows 9x can only access per-process environment
    ' variables)
    ' (link unavailable, check Internet Archive for source)
    'endregion Acknowledgements ####################################################

    'region DependsOn ####################################################
    ' TestObjectForData()
    ' GetWindowsPath()
    ' GetWindowsSystemPath()
    'endregion DependsOn ####################################################

    Dim intFunctionReturn
    Dim intReturnMultiplier
    Dim objWSHShell
    Dim objEnvironment
    Dim strWorkingCommandPromptPath
    Dim intReturnCode
    Dim strWindowsSystemPath
    Dim objFileSystemObject
    Dim boolResult
    Dim strWindowsPath

    Err.Clear
    
    intFunctionReturn = 0
    intReturnMultiplier = 1

    ' Try shell environment variable approach
    On Error Resume Next
    Set objWSHShell = WScript.CreateObject("WScript.Shell")
    If Err Then
        Err.Clear
        On Error Goto 0
        intFunctionReturn = intFunctionReturn + (-1 * intReturnMultiplier)
    Else
        Set objEnvironment = objWSHShell.Environment("Process")
        If Err Then
            Err.Clear
            On Error Goto 0
            intFunctionReturn = intFunctionReturn + (-2 * intReturnMultiplier)
        Else
            strWorkingCommandPromptPath = objEnvironment("COMSPEC")
            If Err Then
                Err.Clear
                On Error Goto 0
                intFunctionReturn = intFunctionReturn + (-3 * intReturnMultiplier)
            Else
                On Error Goto 0
                If TestObjectForData(strWorkingCommandPromptPath) = False Then
                    intFunctionReturn = intFunctionReturn + (-4 * intReturnMultiplier)
                End If
            End If
        End If
    End If

    If intFunctionReturn < 0 Then
        intReturnMultiplier = intReturnMultiplier * 8
        intReturnCode = GetWindowsSystemPath(strWindowsSystemPath)
        If intReturnCode <> 0 Then
            intFunctionReturn = intFunctionReturn + (intReturnCode * intReturnMultiplier)
        Else
            intReturnMultiplier = intReturnMultiplier * 8
            On Error Resume Next
            Set objFileSystemObject = CreateObject("Scripting.FileSystemObject")
            If Err Then
                On Error Goto 0
                Err.Clear
                intFunctionReturn = intFunctionReturn + (-1 * intReturnMultiplier)
            Else
                ' Try cmd.exe
                boolResult = objFileSystemObject.FileExists(strWindowsSystemPath & "cmd.exe")
                If Err Then
                    On Error Goto 0
                    Err.Clear
                    intFunctionReturn = intFunctionReturn + (-2 * intReturnMultiplier)
                Else
                    On Error Goto 0
                    If boolResult = True Then
                        strWorkingCommandPromptPath = strWindowsSystemPath & "cmd.exe"
                        intFunctionReturn = 1
                    Else
                        ' Try command.com
                        On Error Resume Next
                        boolResult = objFileSystemObject.FileExists(strWindowsSystemPath & "command.com")
                        If Err Then
                            On Error Goto 0
                            Err.Clear
                            intFunctionReturn = intFunctionReturn + (-3 * intReturnMultiplier)
                        Else
                            On Error Goto 0
                            If boolResult = True Then
                                strWorkingCommandPromptPath = strWindowsSystemPath & "command.com"
                                intFunctionReturn = 1
                            Else
                                intFunctionReturn = intFunctionReturn + (-4 * intReturnMultiplier)
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If

    If intFunctionReturn < 0 Then
        intReturnMultiplier = intReturnMultiplier * 8
        intReturnCode = GetWindowsPath(strWindowsPath)
        If intReturnCode <> 0 Then
            intFunctionReturn = intFunctionReturn + (intReturnCode * intReturnMultiplier)
        Else
            intReturnMultiplier = intReturnMultiplier * 8
            On Error Resume Next
            Set objFileSystemObject = CreateObject("Scripting.FileSystemObject")
            If Err Then
                On Error Goto 0
                Err.Clear
                intFunctionReturn = intFunctionReturn + (-1 * intReturnMultiplier)
            Else
                ' Try cmd.exe
                boolResult = objFileSystemObject.FileExists(strWindowsPath & "cmd.exe")
                If Err Then
                    On Error Goto 0
                    Err.Clear
                    intFunctionReturn = intFunctionReturn + (-2 * intReturnMultiplier)
                Else
                    On Error Goto 0
                    If boolResult = True Then
                        strWorkingCommandPromptPath = strWindowsPath & "cmd.exe"
                        intFunctionReturn = 2
                    Else
                        ' Try command.com
                        On Error Resume Next
                        boolResult = objFileSystemObject.FileExists(strWindowsPath & "command.com")
                        If Err Then
                            On Error Goto 0
                            Err.Clear
                            intFunctionReturn = intFunctionReturn + (-3 * intReturnMultiplier)
                        Else
                            On Error Goto 0
                            If boolResult = True Then
                                strWorkingCommandPromptPath = strWindowsPath & "command.com"
                                intFunctionReturn = 2
                            Else
                                intFunctionReturn = intFunctionReturn + (-4 * intReturnMultiplier)
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    If intFunctionReturn >= 0 Then
        strCommandPromptPath = strWorkingCommandPromptPath
    End If
    
    GetCommandPromptPath = intFunctionReturn
End Function
