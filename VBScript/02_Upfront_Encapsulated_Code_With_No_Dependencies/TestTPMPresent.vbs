Function TestTPMPresent(ByRef boolTPMPresent)
    'region FunctionMetadata ####################################################
    ' This function determines if at least one TPM is present.
    '
    ' The function takes one positional argument (boolTPMPresent), which is populated upon
    ' success with a boolean True or False. True means that at least one TPM device was
    ' present, while False means that no TPM device was present
    '
    ' The function returns a 0 if the test was performed successfully. It returns a negative
    ' integer if an error occurred performing the test.
    '
    ' Example:
    '   intReturnCode = TestTPMPresent(boolTPMPresent)
    '   If intReturnCode = 0 Then
    '       ' The test was performed successfully
    '       ' boolTPMPresent is True if at least one TPM was present
    '       ' boolTPMPresent is False if no TPMs were present
    '   End If
    '
    ' Version: 1.0.20210617.0
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
    ' ConnectLocalWMINamespace()
    ' TestTPMPresentUsingCIMv2WMINamespace()
    ' TestObjectForData()
    ' GetTPMInstances()
    ' TestTPMPresentUsingTPMWMINamespace()
    'endregion DependsOn ####################################################

    Dim intFunctionReturn
    Dim intReturnMultiplier
    Dim intReturnCode
    Dim objCIMv2WMINamespace
    Dim boolPreliminaryResult
    Dim boolTrySecondMethod
    Dim arrTPMWMINamespaceInstances
    Dim boolSecondResult

    intFunctionReturn = 0
    intReturnMultiplier = 1
    boolPreliminaryResult = Null
    boolSecondResult = Null

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
            intFunctionReturn = intFunctionReturn + (intReturnCode * intReturnMultiplier)
        Else
            intReturnMultiplier = 1
            intReturnCode = TestTPMPresentUsingTPMWMINamespace(boolSecondResult, arrTPMWMINamespaceInstances)
            If intReturnCode <> 0 Then
                intFunctionReturn = intFunctionReturn + (intReturnCode * intReturnMultiplier)
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
    
    TestTPMPresent = intFunctionReturn
End Function
