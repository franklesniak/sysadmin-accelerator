Function TestTPMInstancesForEnabledTPMUsingTPMInstances(ByRef boolEnabledTPMPresent, ByVal arrTPMInstances)
    'region FunctionMetadata ####################################################
    ' Assuming that arrTPMInstances represents an array / collection of the available TPM
    ' instances, this function tests each of them to see if there is at least one enabled TPM
    ' present.
    '
    ' The function takes two positional arguments:
    '  - The first argument (boolEnabledTPMPresent) is populated upon success with a boolean
    '    True or False. True means that at least one enabled TPM was present, while False means
    '    that no enabled TPM was present.
    '  - The second argument (arrTPMInstances) is an array/collection of objects of class
    '    Win32_Tpm
    '
    ' The function returns a 0 if the test was performed successfully. It returns a negative
    ' integer if an error occurred performing the test.
    '
    ' Example:
    '   intReturnCode = GetTPMInstances(arrTPMInstances)
    '   If intReturnCode > 0 Then
    '       ' At least one TPM was retrieved successfully
    '       intReturnCode = TestTPMInstancesForEnabledTPMUsingTPMInstances(boolEnabledTPMPresent, arrTPMInstances)
    '       If intReturnCode = 0 Then
    '           ' The test was performed successfully
    '           ' boolEnabledTPMPresent is True if at least one TPM was present that is Enabled
    '           ' boolEnabledTPMPresent is False if no TPMs were found enabled
    '       End If
    '   End If
    '
    ' Version: 1.0.20230518.0
    'endregion FunctionMetadata ####################################################

    'region License ####################################################
    ' Copyright 2023 Frank Lesniak
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
    ' Microsoft and Sven Aelterman, for providing ZTICheckForTPM/ZTICheckForTPM_v2, which
    ' inspired this approach.
    'endregion Acknowledgements ####################################################

    'region DependsOn ####################################################
    ' TestObjectForData()
    ' TestObjectIsAnyTypeOfInteger()
    'endregion DependsOn ####################################################

    Dim intFunctionReturn
    Dim intReturnMultiplier
    Dim intTemp
    Dim boolResult
    Dim objTPMInstance
    Dim boolSingleResult
    Dim intReturnCode
    Dim boolTest

    Err.Clear

    intFunctionReturn = 0
    intReturnMultiplier = 128

    If TestObjectForData(arrTPMInstances) <> True Then
        intFunctionReturn = intFunctionReturn + (-1 * intReturnMultiplier)
    Else
        On Error Resume Next
        intTemp = arrTPMInstances.Count
        If Err Then
            On Error Goto 0
            Err.Clear
            intFunctionReturn = intFunctionReturn + (-2 * intReturnMultiplier)
        Else
            On Error Goto 0
            If TestObjectIsAnyTypeOfInteger(intTemp) = False Then
                intFunctionReturn = intFunctionReturn + (-3 * intReturnMultiplier)
            Else
                If intTemp < 0 Then
                    intFunctionReturn = intFunctionReturn + (-4 * intReturnMultiplier)
                ElseIf intTemp = 0 Then
                    intFunctionReturn = intFunctionReturn + (-5 * intReturnMultiplier)
                Else
                    boolResult = False
                    On Error Resume Next
                    For Each objTPMInstance in arrTPMInstances
                        If Err Then
                            Err.Clear
                        Else
                            intReturnCode = objTPMInstance.IsEnabled(boolSingleResult)
                            If Err Then
                                Err.Clear
                                ' Assume it's not enabled and keep going
                            Else
                                If intReturnCode <> 0 Then
                                    ' Assume it's not enabled and keep going
                                Else
                                    boolTest = (boolSingleResult = True)
                                    If Err Then
                                        Err.Clear
                                        ' Assume it's not enabled and keep going
                                    Else
                                        If boolTest = True Then
                                            boolResult = True
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next
                    On Error Goto 0
                    If Err Then
                        Err.Clear
                    End If
                End If
            End If
        End If

        intReturnMultiplier = intReturnMultiplier * 8
        If intFunctionReturn < 0 Then
            ' Perhaps a single TPM instance was passed in, rather than an array/collection
            ' Try it
            On Error Resume Next
            intReturnCode = arrTPMInstances.IsEnabled(boolSingleResult)
            If Err Then
                On Error Goto 0
                Err.Clear
                intFunctionReturn = intFunctionReturn + (-1 * intReturnMultiplier)
            Else
                On Error Goto 0
                ' We seem to have retrieved the IsEnabled status
                ' Reset the function return code
                intFunctionReturn = 0
                boolResult = False
                If intReturnCode <> 0 Then
                    ' Assume it's not enabled
                Else
                    On Error Resume Next
                    boolTest = (boolSingleResult = True)
                    If Err Then
                        On Error Goto 0
                        Err.Clear
                        ' Assume it's not enabled and keep going
                    Else
                        On Error Goto 0
                        If boolTest = True Then
                            boolResult = True
                        End If
                    End If
                End If
            End If
        End If
    End If

    If intFunctionReturn = 0 Then
        boolEnabledTPMPresent = boolResult
    End If
    
    TestTPMInstancesForEnabledTPMUsingTPMInstances = intFunctionReturn
End Function
