Function GetServiceWaitHintUsingServiceInstance(ByRef intWaitHint, ByVal objServiceInstance)
    'region FunctionMetadata #######################################################
    ' Assuming that objServiceInstance represents an instance of the WMI service class
    ' (of type Win32_Service), this function retrieves the wait hint of the service.
    ' The wait hint is used to provide an estimate (in milliseconds) of the of the
    ' total time expected for a pending operation to complete. It provides a hint to
    ' the Service Control Manager (along with the CheckPoint property not incrementing)
    ' that the service has failed to change state.
    '
    ' The function takes two positional arguments:
    '  - The first argument (intWaitHint) is populated upon success with an integer
    '    representing the service's wait hint, i.e., the estimated number of
    '    milliseconds for the current operation to complete.
    '  - The second argument (objServiceInstance) is an instance of the WMI class
    '    Win32_Service
    '
    ' The function returns a 0 if the function could determine the value of the
    ' service's wait hint. It returns a negative integer if an error occurred
    ' determining the service's wait hint.
    '
    ' Example:
    '   intReturnCode = GetServiceInstances(arrServiceInstances)
    '   If intReturnCode > 0 Then
    '       ' At least one Win32_Service instance was retrieved successfully
    '       For Each objServiceInstance In arrServiceInstances
    '           intReturnCode = GetServiceWaitHintUsingServiceInstance(intWaitHint, objServiceInstance)
    '           If intReturnCode = 0 Then
    '               ' The function successfully determined the service's wait hint. The
    '               ' result is stored in intWaitHint
    '           End If
    '       Next
    '   End If
    '
    ' Version: 1.0.20230815.0
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
    ' TestObjectIsAnyTypeOfInteger()
    'endregion DependsOn ##############################################################

    Dim intFunctionReturn
    Dim intReturnMultiplier
    Dim intInterimResult
    Dim intResultToReturn

    Err.Clear

    intFunctionReturn = 0
    intReturnMultiplier = 128
    intInterimResult = Empty
    intResultToReturn = Empty

    If TestObjectForData(objServiceInstance) <> True Then
        intFunctionReturn = intFunctionReturn + (-1 * intReturnMultiplier)
    Else
        On Error Resume Next
        intInterimResult = objServiceInstance.WaitHint
        If Err Then
            Err.Clear
            On Error GoTo 0
            intFunctionReturn = intFunctionReturn + (-2 * intReturnMultiplier)
        Else
            On Error GoTo 0
            If TestObjectIsAnyTypeOfInteger(intInterimResult) <> True Then
                intFunctionReturn = intFunctionReturn + (-3 * intReturnMultiplier)
            Else
                intResultToReturn = intInterimResult
            End If
        End If
    End If

    If intFunctionReturn >= 0 Then
        intWaitHint = intResultToReturn
    End If
    
    GetServiceWaitHintUsingServiceInstance = intFunctionReturn
End Function
