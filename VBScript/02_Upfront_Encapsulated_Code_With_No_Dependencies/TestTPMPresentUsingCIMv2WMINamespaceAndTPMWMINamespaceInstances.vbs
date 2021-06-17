Function TestTPMPresentUsingCIMv2WMINamespaceAndTPMWMINamespaceInstances(ByRef boolTPMPresent, ByVal objCIMv2WMINamespace, ByVal arrTPMWMINamespaceInstances)
    'region FunctionMetadata ####################################################
    ' Assuming that objCIMv2WMINamespace represents a connection to the root\CIMv2 WMI
    ' namespace, and assuming that arrTPMWMINamespaceInstances represents an array / collection
    ' of the available TPM instances, this function uses one or both methods to determine if at
    ' least one TPM is present.
    '
    ' The function takes three positional arguments:
    '  - The first argument (boolTPMPresent) is populated upon success with a boolean True or
    '    False. True means that at least one TPM device was present, while False means that no
    '    TPM device was present
    '  - The second argument (objCIMv2WMINamespace) is a connection to the root\CIMv2 WMI
    '    namespace
    '  - The third argument (arrTPMWMINamespaceInstances) is an array/collection of objects of
    '    class Win32_Tpm
    '
    ' The function returns a 0 if the test was performed successfully. It returns a negative
    ' integer if an error occurred performing the test.
    '
    ' Example:
    '   intReturnCode = ConnectLocalWMINamespace(objCIMv2WMINamespace, Null, Null)
    '   If intReturnCode = 0 Then
    '       intReturnCode = GetTPMInstances(arrTPMWMINamespaceInstances)
    '       If intReturnCode > 0 Then
    '           ' At least one TPM was retrieved successfully
    '           intReturnCode = TestTPMPresentUsingCIMv2WMINamespaceAndTPMWMINamespaceInstances(boolTPMPresent, objCIMv2WMINamespace, arrTPMWMINamespaceInstances)
    '           If intReturnCode = 0 Then
    '               ' The test was performed successfully
    '               ' boolTPMPresent is True if at least one TPM was present
    '               ' boolTPMPresent is False if no TPMs were present
    '           End If
    '       End If
    '   End If
    '
    ' Version: 1.0.20210616.0
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
    ' Microsoft and Sven Aelterman, for providing ZTICheckForTPM/ZTICheckForTPM_v2, which
    ' inspired this approach.
    'endregion Acknowledgements ####################################################

    'region DependsOn ####################################################
    ' TestObjectForData()
    ' TestObjectIsAnyTypeOfInteger()
    'endregion DependsOn ####################################################

    Dim intFunctionReturn
    Dim intReturnMultiplier
    Dim arrTPMDevices
    Dim intTemp
    Dim boolTest
    Dim boolPreliminaryResult
    Dim boolSecondResult
    Dim boolTrySecondMethod

    Err.Clear

    intFunctionReturn = 0
    intReturnMultiplier = 2048
    boolPreliminaryResult = Null
    boolSecondResult = Null

    If TestObjectForData(objCIMv2WMINamespace) <> True Then
        intFunctionReturn = intFunctionReturn + (-1 * intReturnMultiplier)
    Else
        On Error Resume Next
        Set arrTPMDevices = objCIMv2WMINamespace.ExecQuery("SELECT Caption FROM Win32_PnPEntity WHERE ClassGuid = ""{d94ee5d8-d189-4994-83d2-f68d7d41b0e6}""")
        If Err Then
            On Error Goto 0
            Err.Clear
            intFunctionReturn = intFunctionReturn + (-2 * intReturnMultiplier)
        Else
            intTemp = arrTPMDevices.Count
            If Err Then
                On Error Goto 0
                Err.Clear
                intFunctionReturn = intFunctionReturn + (-3 * intReturnMultiplier)
            Else
                On Error Goto 0
                If TestObjectIsAnyTypeOfInteger(intTemp) = False Then
                    intFunctionReturn = intFunctionReturn + (-4 * intReturnMultiplier)
                Else
                    On Error Resume Next
                    boolTest = (intTemp > 0)
                    If Err Then
                        On Error Goto 0
                        Err.Clear
                        intFunctionReturn = intFunctionReturn + (-5 * intReturnMultiplier)
                    Else
                        On Error Goto 0
                        If boolTest = True Then
                            boolPreliminaryResult = True
                        Else
                            boolPreliminaryResult = False
                        End If
                    End If
                End If
            End If
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
        intReturnMultiplier = intReturnMultiplier * 8
        If TestObjectForData(arrTPMWMINamespaceInstances) <> True Then
            intFunctionReturn = intFunctionReturn + (-1 * intReturnMultiplier)
        Else
            On Error Resume Next
            intTemp = arrTPMWMINamespaceInstances.Count
            If Err Then
                On Error Goto 0
                Err.Clear
                intFunctionReturn = intFunctionReturn + (-2 * intReturnMultiplier)
            Else
                On Error Goto 0
                If TestObjectIsAnyTypeOfInteger(intTemp) = False Then
                    intFunctionReturn = intFunctionReturn + (-3 * intReturnMultiplier)
                Else
                    On Error Resume Next
                    boolTest = (intTemp > 0)
                    If Err Then
                        On Error Goto 0
                        Err.Clear
                        intFunctionReturn = intFunctionReturn + (-4 * intReturnMultiplier)
                    Else
                        On Error Goto 0
                        If boolTest = True Then
                            boolSecondResult = True
                        Else
                            boolSecondResult = False
                        End If
                    End If
                End If
            End If
        End If
    End If

    If TestObjectForData(boolPreliminaryResult) = True Then
        ' boolPreliminaryResult is True or False
        If boolPreliminaryResult = True Then
            ' boolPreliminaryResult is True
            boolTPMPresent = True
            intFunctionReturn = 0
        Else
            ' boolPreliminaryResult is False
            If TestObjectForData(boolSecondResult) = True Then
                ' boolPreliminaryResult is False
                ' boolSecondResult is True or False
                If boolSecondResult = True Then
                    ' boolPreliminaryResult is False
                    ' boolSecondResult is True
                    boolTPMPresent = True
                    intFunctionReturn = 0
                Else
                    ' boolPreliminaryResult is False
                    ' boolSecondResult is False
                    boolTPMPresent = False
                    intFunctionReturn = 0
                End If
            Else
                ' boolPreliminaryResult is False
                ' boolSecondResult is undefined because the second method failed
                boolTPMPresent = False
                intFunctionReturn = 0
            End If
        End If
    Else
        ' boolPreliminaryResult is undefined because the first method failed
        If TestObjectForData(boolSecondResult) = True Then
            ' boolPreliminaryResult is undefined because the first method failed
            ' boolSecondResult is True or False
            boolTPMPresent = boolSecondResult
            intFunctionReturn = 0
        Else
            ' boolPreliminaryResult is undefined because the first method failed
            ' boolSecondResult is undefined because the second method failed
            ' Don't adjust boolTPMPresent or intFunctionReturn
        End If
    End If
    
    TestTPMPresentUsingCIMv2WMINamespaceAndTPMWMINamespaceInstances = intFunctionReturn
End Function
