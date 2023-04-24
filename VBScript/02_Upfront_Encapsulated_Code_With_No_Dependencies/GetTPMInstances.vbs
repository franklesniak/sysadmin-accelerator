Function GetTPMInstances(ByRef arrTPMInstances)
    'region FunctionMetadata #######################################################
    ' This function retrieves the available TPM instances as objects of the class
    ' Win32_Tpm. Typically, there is only one TPM available on a computer.
    '
    ' The function takes one positional argument (arrTPMInstances), which is populated
    ' upon success with a collection of TPM instances of the type Win32_Tpm
    '
    ' The function returns the number of TPM instances retrieved. A successful function
    ' call returns a positive integer. The function will return 0 if no error occurred,
    ' but no TPMs were present. Finally, the function returns a negative integer if an
    ' error occurs.
    '
    ' Example:
    '   intReturnCode = GetTPMInstances(arrTPMInstances)
    '   If intReturnCode > 0 Then
    '       ' At least one TPM was retrieved successfully
    '   End If
    '
    ' Version: 1.0.20210615.1
    'endregion FunctionMetadata #######################################################

    'region License ################################################################
    ' Copyright 2021 Frank Lesniak
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
    ' ConnectLocalTPMWMINamespace()
    ' GetTPMInstancesUsingWMINamespace()
    'endregion DependsOn ##############################################################

    Dim intFunctionReturn
    Dim intReturnMultiplier
    Dim intReturnCode
    Dim objSWbemServicesTPMWMINamespace

    intFunctionReturn = 0
    intReturnMultiplier = 1

    intReturnCode = ConnectLocalTPMWMINamespace(objSWbemServicesTPMWMINamespace)
    If intReturnCode < 0 Then
        intFunctionReturn = intFunctionReturn + (intReturnCode * intReturnMultiplier)
    Else
        intReturnCode = GetTPMInstancesUsingWMINamespace(arrTPMInstances, objSWbemServicesTPMWMINamespace)
        intFunctionReturn = intFunctionReturn + (intReturnCode * intReturnMultiplier)
    End If
    
    GetTPMInstances = intFunctionReturn
End Function
