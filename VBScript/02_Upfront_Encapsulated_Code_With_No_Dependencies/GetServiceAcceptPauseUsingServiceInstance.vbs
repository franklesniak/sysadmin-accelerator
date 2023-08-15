Function GetServiceAcceptPauseUsingServiceInstance(ByRef boolAcceptPause, ByVal objServiceInstance)
    'region FunctionMetadata #######################################################
    ' Assuming that objServiceInstance represents an instance of the WMI service class
    ' (of type Win32_Service), this function determines whether the service accepts a
    ' pause command
    '
    ' The function takes two positional arguments:
    '  - The first argument (boolAcceptPause) is populated upon success with a
    '    boolean indicating whether the service accepts a pause command, as reported by
    '    WMI
    '  - The second argument (objServiceInstance) is an instance of the WMI class
    '    Win32_Service
    '
    ' The function returns a 0 if the function could determine whether the service
    ' accepts a pause command successfully. It returns a negative integer if an error
    ' occurred determining whether the function accepts a pause.
    '
    ' Example:
    '   intReturnCode = GetServiceInstances(arrServiceInstances)
    '   If intReturnCode > 0 Then
    '       ' At least one Win32_Service instance was retrieved successfully
    '       For Each objServiceInstance In arrServiceInstances
    '           intReturnCode = GetServiceAcceptPauseUsingServiceInstance(boolAcceptPause, objServiceInstance)
    '           If intReturnCode = 0 Then
    '               ' The function successfully determined whether or not the service
    '               ' accepts a pause command. The result is stored in boolAcceptPause
    '           End If
    '       Next
    '   End If
    '
    ' Version: 1.0.20230814.0
    'endregion FunctionMetadata #######################################################

    'region License ################################################################
    ' Copyright 2023 Frank Lesniak
    '
    ' Permission is hereby granted, free of charge, to any person obtaining a copy of
    ' this software and associated documentation files (the "Software"), to deal in the
    ' Software without restriction, including without limitation the rights to use,
    ' copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the
    ' Software, and to permit persons to whom the Software is furnished to do so,
    ' subject to the following conditions:
    '
    ' The above copyright notice and this permission notice shall be included in all
    ' copies or substantial portions of the Software.
    '
    ' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    ' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
    ' FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR
    ' COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN
    ' AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
    ' WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
    'endregion License ################################################################

    'region DownloadLocationNotice #################################################
    ' The most up-to-date version of this script can be found on the author's GitHub
    ' repository at https://github.com/franklesniak/sysadmin-accelerator
    'endregion DownloadLocationNotice #################################################

    'region Acknowledgements #######################################################
    ' None!
    'endregion Acknowledgements #######################################################

    'region DependsOn ##############################################################
    ' TestObjectForData()
    ' TestObjectIsBoolean()
    'endregion DependsOn ##############################################################

    Dim intFunctionReturn
    Dim intReturnMultiplier
    Dim boolInterimResult
    Dim boolResultToReturn

    Err.Clear

    intFunctionReturn = 0
    intReturnMultiplier = 128
    boolInterimResult = Empty
    boolResultToReturn = Empty

    If TestObjectForData(objServiceInstance) <> True Then
        intFunctionReturn = intFunctionReturn + (-1 * intReturnMultiplier)
    Else
        On Error Resume Next
        boolInterimResult = objServiceInstance.AcceptPause
        If Err Then
            Err.Clear
            On Error GoTo 0
            intFunctionReturn = intFunctionReturn + (-2 * intReturnMultiplier)
        Else
            On Error GoTo 0
            If TestObjectIsBoolean(boolInterimResult) <> True Then
                intFunctionReturn = intFunctionReturn + (-3 * intReturnMultiplier)
            Else
                boolResultToReturn = boolInterimResult
            End If
        End If
    End If

    If intFunctionReturn >= 0 Then
        boolAcceptPause = boolResultToReturn
    End If
    
    GetServiceAcceptPauseUsingServiceInstance = intFunctionReturn
End Function
