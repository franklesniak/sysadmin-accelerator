Function GetTPMInstancesUsingWMINamespace(ByRef arrTPMInstances, ByVal objWMINamespace)
    'region FunctionMetadata ####################################################
    ' Assuming that objWMINamespace represents a successful connection to the a TPM WMI
    ' namespace, this function retrieves the available TPM instances as objects of the class
    ' Win32_Tpm. Typically, there is only one TPM available on a computer.
    '
    ' The function takes two positional arguments:
    '  - The first argument (arrTPMInstances) is populated upon success with a collection of
    '    TPM instances of the type Win32_Tpm
    '  - The second argument (objWMINamespace) is a WMI Namespace connection argument that must
    '    already be connected to the WMI namespace root\cimv2\Security\MicrosoftTpm
    '
    ' The function returns the number of TPM instances retrieved. A successful function call
    ' returns a positive integer. The function will return 0 if no error occurred, but no TPMs
    ' were present. Finally, the function returns a negative integer if an error occurs.
    '
    ' Example:
    '   intReturnCode = ConnectLocalTPMWMINamespace(objSWbemServicesTPMWMINamespace)
    '   If intReturnCode = 0 Then
    '       ' Successfully connected to the local computer's TPM WMI Namespace
    '       intReturnCode = GetTPMInstancesUsingWMINamespace(arrTPMInstances, objSWbemServicesTPMWMINamespace)
    '       If intReturnCode > 0 Then
    '           ' At least one TPM was retrieved successfully
    '       End If
    '   End If
    '
    ' Version: 1.0.20210615.0
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
    Dim arrWorkingTPMInstances
    Dim intTemp

    Err.Clear

    intFunctionReturn = 0
    intReturnMultiplier = 16

    If TestObjectForData(objWMINamespace) <> True Then
        intFunctionReturn = intFunctionReturn + (-1 * intReturnMultiplier)
    Else
        On Error Resume Next
        Set arrWorkingTPMInstances = objWMINamespace.InstancesOf("Win32_Tpm")
        If Err Then
            On Error Goto 0
            Err.Clear
            intFunctionReturn = intFunctionReturn + (-2 * intReturnMultiplier)
        Else
            intTemp = arrWorkingTPMInstances.Count
            If Err Then
                On Error Goto 0
                Err.Clear
                intFunctionReturn = intFunctionReturn + (-3 * intReturnMultiplier)
            Else
                On Error Goto 0
                If TestObjectIsAnyTypeOfInteger(intTemp) = False Then
                    intFunctionReturn = intFunctionReturn + (-4 * intReturnMultiplier)
                Else
                    If intTemp < 0 Then
                        intFunctionReturn = intFunctionReturn + (-5 * intReturnMultiplier)
                    ElseIf intTemp > 0 Then
                        intFunctionReturn = intTemp
                    End If
                End If
            End If
        End If
    End If

    If intFunctionReturn > 0 Then
        On Error Resume Next
        Set arrTPMInstances = objWMINamespace.InstancesOf("Win32_Tpm")
        If Err Then
            On Error Goto 0
            Err.Clear
            intFunctionReturn = (-6 * intReturnMultiplier)
        Else
            On Error Goto 0
        End If
    End If
    
    GetTPMInstancesUsingWMINamespace = intFunctionReturn
End Function
