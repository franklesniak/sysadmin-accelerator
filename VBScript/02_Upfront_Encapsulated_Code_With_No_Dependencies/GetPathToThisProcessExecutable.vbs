Function GetPathToThisProcessExecutable(ByRef strPathToThisExecutable)
    'region FunctionMetadata ####################################################
    ' Safely determines the path to this process's executable
    '
    ' This function takes one argument (strPathToThisExecutable), which is populated upon
    ' success with the path to this process's wscript.exe/cscript.exe executable.
    '
    ' The function returns 0 if the path to this process's architecture was determined
    ' successfully. It returns a negative number if it was not determined successfully.
    '
    ' Example:
    '   intReturnCode = GetPathToThisProcessExecutable(strPathToThisExecutable)
    '   If intReturnCode = 0 Then
    '       ' Process completed successfully
    '       ' strPathToThisExecutable contains a path like "C:\Windows\System32\wscript.exe"
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
    ' at https://github.com/franklesniak/sysadmin-accelerator
    'endregion DownloadLocationNotice ####################################################

    'region DependsOn ####################################################
    ' TestObjectForData()
    'endregion DependsOn ####################################################

    Dim intFunctionReturn
    Dim strWorkingPath

    Err.Clear

    intFunctionReturn = 0
    
    On Error Resume Next
    strWorkingPath = WScript.FullName
    If Err Then
        On Error Goto 0
        Err.Clear
        intFunctionReturn = -1
    Else
        On Error Goto 0
        If TestObjectForData(strWorkingPath) = False Then
            intFunctionReturn = -2
        End If
    End If

    If intFunctionReturn = 0 Then
        strPathToThisExecutable = strWorkingPath
    End If

    GetPathToThisProcessExecutable = intFunctionreturn
End Function
