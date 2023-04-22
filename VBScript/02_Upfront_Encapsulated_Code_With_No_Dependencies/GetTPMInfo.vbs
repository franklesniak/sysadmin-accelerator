Function GetTPMInfo(ByRef intTriStateTPMPresent, ByRef intTriStateEnabledTPMPresent, ByRef intTriStateActivatedTPMPresent, ByRef intTriStateReadyTPMPresent)
    'region FunctionMetadata #######################################################
    ' This function retrieves the most commonly needed TPM information, namely:
    '  - Whether a TPM was present
    '  - If a TPM was present, and if this script was run with elevated privileges:
    '     - Whether at least one TPM was enabled
    '     - Whether at least one TPM was activated
    '     - Whether at least one TPM was ready
    '
    ' The function takes four positional arguments:
    '  - The first argument (intTriStateTPMPresent) is an integer that will be
    '    populated with one of three values:
    '        -1 = a TPM was present
    '         0 = a TPM was not present
    '         1 = it could not be determined whether a TPM was present. This is usually
    '             caused by restrictions that prevent access to WMI
    '  - The second argument (intTriStateEnabledTPMPresent) is an integer that will be
    '    populated with one of three values:
    '        -1 = at least one enabled TPM was present
    '         0 = no enabled TPM was present
    '         1 = it could not be determined whether an enabled TPM was present. This
    '             is usually caused by the script being run without elevated
    '             permissions
    '  - The third argument (intTriStateActivatedTPMPresent) is an integer that will be
    '    populated with one of three values:
    '        -1 = at least one activated TPM was present
    '         0 = no activated TPM was present
    '         1 = it could not be determined whether an activated TPM was present. This
    '             is usually caused by the script being run without elevated
    '             permissions
    '  - The fourth argument (intTriStateReadyTPMPresent) is an integer that will be
    '    populated with one of three values:
    '        -1 = at least one ready TPM was present
    '         0 = no ready TPM was present
    '         1 = it could not be determined whether a ready TPM was present. This is
    '             usually caused by the script being run without elevated permissions
    '
    ' The function returns a 0 if all TPM info was retrieved successfully. It returns a
    ' positive integer if it was determined whether a TPM was present, but the TPM(s)'
    ' enabled status, activation status, or readiness could not be determined. The
    ' function returns a negative integer if an error occurred and no information could
    ' be determined.
    '
    ' Example:
    '   intReturnCode = GetTPMInfo(intTriStateTPMPresent, intTriStateEnabledTPMPresent, intTriStateActivatedTPMPresent, intTriStateReadyTPMPresent)
    '   If intReturnCode >= 0 Then
    '       ' intTriStateTPMPresent is populated as expected
    '   End If
    '   If intReturnCode = 0 Then
    '       ' intTriStateEnabledTPMPresent, intTriStateActivatedTPMPresent, and
    '       ' intTriStateReadyTPMPresent are populated as expected
    '   End If
    '
    ' Version: 1.1.20230422.0
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
    ' Microsoft and Sven Aelterman, for providing ZTICheckForTPM/ZTICheckForTPM_v2,
    ' which inspired this approach.
    'endregion Acknowledgements #######################################################

    'region DependsOn ##############################################################
    ' ConnectLocalWMINamespace()
    ' TestTPMPresentUsingCIMv2WMINamespace()
    ' TestObjectForData()
    ' GetTPMInstances()
    ' TestTPMPresentUsingTPMWMINamespace()
    ' TestTPMInstancesForEnabledTPMUsingTPMInstances()
    ' TestTPMInstancesForActivatedTPMUsingTPMInstances()
    ' TestTPMInstancesForReadyTPMUsingTPMInstances()
    'endregion DependsOn ##############################################################

    Dim intFunctionReturn
    Dim intReturnMultiplier
    Dim boolPreliminaryResult
    Dim boolSecondResult
    Dim intPreliminaryTriStateTPMPresent
    Dim intPreliminaryTriStateEnabledTPMPresent
    Dim intPreliminaryTriStateActivatedTPMPresent
    Dim intPreliminaryTriStateReadyTPMPresent
    Dim intReturnCode
    Dim objCIMv2WMINamespace
    Dim boolTrySecondMethod
    Dim arrTPMWMINamespaceInstances
    Dim intTriStateTPMNamespaceInstancesFailed
    Dim boolTemp

    intFunctionReturn = 0
    intReturnMultiplier = 1
    boolPreliminaryResult = Null
    boolSecondResult = Null
    intPreliminaryTriStateTPMPresent = 1
    intPreliminaryTriStateEnabledTPMPresent = 1
    intPreliminaryTriStateActivatedTPMPresent = 1
    intPreliminaryTriStateReadyTPMPresent = 1
    intTriStateTPMNamespaceInstancesFailed = 1

    intReturnCode = ConnectLocalWMINamespace(objCIMv2WMINamespace, Null, Null)
    If intReturnCode <> 0 Then
        intFunctionReturn = intFunctionReturn + (intReturnCode * intReturnMultiplier)
    Else
        intReturnCode = TestTPMPresentUsingCIMv2WMINamespace(boolPreliminaryResult, objCIMv2WMINamespace)
        If intReturnCode <> 0 Then
            intFunctionReturn = intFunctionReturn + (intReturnCode * intReturnMultiplier)
        End If
    End If

    boolTrySecondMethod = False
    If intFunctionReturn < 0 Then
        boolTrySecondMethod = True
    End If
    If TestObjectForData(boolPreliminaryResult) <> True Then
        boolTrySecondMethod = True
    Else
        If boolPreliminaryResult = False Then
            boolTrySecondMethod = True
        End If
    End If

    If boolTrySecondMethod = True Then
        intReturnMultiplier = intReturnMultiplier * 16
        intReturnCode = GetTPMInstances(arrTPMWMINamespaceInstances)
        If intReturnCode < 0 Then
            intTriStateTPMNamespaceInstancesFailed = -1
            intFunctionReturn = intFunctionReturn + (intReturnCode * intReturnMultiplier)
        Else
            If intReturnCode = 0 Then
                intTriStateTPMNamespaceInstancesFailed = -1
            Else
                intTriStateTPMNamespaceInstancesFailed = 0
            End If
            intReturnMultiplier = 1
            intReturnCode = TestTPMPresentUsingTPMWMINamespace(boolSecondResult, arrTPMWMINamespaceInstances)
            If intReturnCode <> 0 Then
                intFunctionReturn = intFunctionReturn + (intReturnCode * intReturnMultiplier)
            Else
                intTriStateTPMNamespaceInstancesFailed = 0
            End If
        End If
    End If

    If TestObjectForData(boolPreliminaryResult) = True Then
        ' boolPreliminaryResult is True or False
        If boolPreliminaryResult = True Then
            ' boolPreliminaryResult is True
            intPreliminaryTriStateTPMPresent = -1
            intFunctionReturn = 0
        Else
            ' boolPreliminaryResult is False
            If TestObjectForData(boolSecondResult) = True Then
                ' boolPreliminaryResult is False
                ' boolSecondResult is True or False
                If boolSecondResult = True Then
                    ' boolPreliminaryResult is False
                    ' boolSecondResult is True
                    intPreliminaryTriStateTPMPresent = -1
                    intFunctionReturn = 0
                Else
                    ' boolPreliminaryResult is False
                    ' boolSecondResult is False
                    intPreliminaryTriStateTPMPresent = 0
                    intFunctionReturn = 0
                End If
            Else
                ' boolPreliminaryResult is False
                ' boolSecondResult is undefined because the second method failed
                intPreliminaryTriStateTPMPresent = 0
                intFunctionReturn = 0
            End If
        End If
    Else
        ' boolPreliminaryResult is undefined because the first method failed
        If TestObjectForData(boolSecondResult) = True Then
            ' boolPreliminaryResult is undefined because the first method failed
            ' boolSecondResult is True or False
            If boolSecondResult = True Then
                intPreliminaryTriStateTPMPresent = -1
            Else
                intPreliminaryTriStateTPMPresent = 0
            End If
            intFunctionReturn = 0
        Else
            ' boolPreliminaryResult is undefined because the first method failed
            ' boolSecondResult is undefined because the second method failed
            ' Don't adjust intPreliminaryTriStateTPMPresent or intFunctionReturn
        End If
    End If

    If intFunctionReturn <> 0 Then
        intTriStateTPMPresent = 1 ' Mark as unknown
        intTriStateEnabledTPMPresent = 1 ' Mark as unknown
        intTriStateActivatedTPMPresent = 1 ' Mark as unknown
        intTriStateReadyTPMPresent = 1 ' Mark as unknown
    Else
        ' No error occurred yet
        intTriStateTPMPresent = intPreliminaryTriStateTPMPresent

        If intTriStateTPMPresent = 0 Then
            ' There is no TPM present. Therefore, it is not possible for there to be a
            ' TPM that is enabled, activated, or ready.
            intTriStateEnabledTPMPresent = 0
            intTriStateActivatedTPMPresent = 0
            intTriStateReadyTPMPresent = 0
        Else
            ' There is a TPM present. Proceed with testing for enabled, activated, and
            ' ready TPMs.
            If intTriStateTPMNamespaceInstancesFailed <> -1 Then
                ' Use of the TPM WMI namespace has not already failed
                If intTriStateTPMNamespaceInstancesFailed = 1 Or TestObjectForData(arrTPMWMINamespaceInstances) = False Then
                    ' TPM namespace instances are not yet retrieved
                    intReturnMultiplier = intReturnMultiplier * 16
                    intReturnCode = GetTPMInstances(arrTPMWMINamespaceInstances)
                    If intReturnCode < 0 Then
                        intTriStateTPMNamespaceInstancesFailed = -1
                        intFunctionReturn = intFunctionReturn + (intReturnCode * intReturnMultiplier)
                    ElseIf intReturnCode = 0 Then
                        intTriStateTPMNamespaceInstancesFailed = -1
                    Else
                        intTriStateTPMNamespaceInstancesFailed = 0
                    End If
                    intReturnMultiplier = 1
                End If

                ' TPM namespace instances should either be retrieved, or they should be
                ' failed
                If intTriStateTPMNamespaceInstancesFailed = 0 Then
                    ' TPM namespace instances retrieved (no failure) and stored in
                    ' arrTPMWMINamespaceInstances

                    intReturnMultiplier = intReturnMultiplier * 2048 * 8

                    intReturnCode = TestTPMInstancesForEnabledTPMUsingTPMInstances(boolTemp, arrTPMWMINamespaceInstances)
                    If intReturnCode <> 0 Then
                        intFunctionReturn = intFunctionReturn + (-1 * intReturnMultiplier)
                    Else
                        If boolTemp = True Then
                            intPreliminaryTriStateEnabledTPMPresent = -1
                        Else
                            intPreliminaryTriStateEnabledTPMPresent = 0
                        End If
                    End If

                    intReturnCode = TestTPMInstancesForActivatedTPMUsingTPMInstances(boolTemp, arrTPMWMINamespaceInstances)
                    If intReturnCode <> 0 Then
                        intFunctionReturn = intFunctionReturn + (-2 * intReturnMultiplier)
                    Else
                        If boolTemp = True Then
                            intPreliminaryTriStateActivatedTPMPresent = -1
                        Else
                            intPreliminaryTriStateActivatedTPMPresent = 0
                        End If
                    End If

                    intReturnCode = TestTPMInstancesForReadyTPMUsingTPMInstances(boolTemp, arrTPMWMINamespaceInstances)
                    If intReturnCode <> 0 Then
                        intFunctionReturn = intFunctionReturn + (-3 * intReturnMultiplier)
                    Else
                        If boolTemp = True Then
                            intPreliminaryTriStateReadyTPMPresent = -1
                        Else
                            intPreliminaryTriStateReadyTPMPresent = 0
                        End If
                    End If

                    If intFunctionReturn <> 0 Then
                        ' An error occurred in this section
                        intFunctionReturn = 1
                    End If
                Else
                    ' The retrieval of TPM namespace instances failed
                    intFunctionReturn = 1
                End If
            Else
                ' The retrieval of TPM namespace instances has already failed
                intFunctionReturn = 1
            End If
        End If

        intTriStateEnabledTPMPresent = intPreliminaryTriStateEnabledTPMPresent
        intTriStateActivatedTPMPresent = intPreliminaryTriStateActivatedTPMPresent
        intTriStateReadyTPMPresent = intPreliminaryTriStateReadyTPMPresent
    End If
    
    GetTPMInfo = intFunctionReturn
End Function
