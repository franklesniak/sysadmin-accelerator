Function GetComputerSystemProductInstances(ByRef arrComputerSystemProductInstances)
    'region FunctionMetadata #######################################################
    ' This function retrieves the Win32_ComputerSystemProduct instances and stores them
    ' in arrComputerSystemProductInstances.
    '
    ' The function takes one positional argument (arrComputerSystemProductInstances),
    ' which is populated upon success with the computer system product instances
    ' returned from WMI of type Win32_ComputerSystemProduct
    '
    ' The function returns 0 if Win32_ComputerSystemProduct instances were retrieved
    ' successfully, and there was one computer system product instance (as expected).
    ' If no Win32_ComputerSystemProduct objects could be retrieved, then the function
    ' returns a negative number. If there are unexpectedly multiple instances of
    ' Win32_ComputerSystemProduct, then the function returns a positive number equal to
    ' the number of WMI instances retrieved minus one.
    '
    ' Example:
    '   intReturnCode = GetComputerSystemProductInstances(arrComputerSystemProductInstances)
    '   If intReturnCode = 0 Then
    '       ' The Win32_ComputerSystemProduct instance was retrieved successfully and
    '       ' is available at arrComputerSystemProductInstances.ItemIndex(0)
    '   ElseIf intReturnCode > 0 Then
    '       ' More than one Win32_ComputerSystemProduct instance was retrieved, which
    '       ' is unexpected
    '   Else
    '       ' An error occurred and no Win32_ComputerSystemProduct instances were
    '       ' retrieved
    '   End If
    '
    ' Version: 1.0.20230422.0
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
    ' The most up-to-date version of this script can be found on the author's GitHub repository
    ' at https://github.com/franklesniak/sysadmin-accelerator
    'endregion DownloadLocationNotice #################################################

    'region Acknowledgements #######################################################
    ' None!
    'endregion Acknowledgements #######################################################

    'region DependsOn ##############################################################
    ' ConnectLocalWMINamespace()
    ' GetComputerSystemProductInstancesUsingWMINamespace()
    'endregion DependsOn ##############################################################

    Dim intFunctionReturn
    Dim intReturnMultiplier
    Dim intReturnCode
    Dim objSWbemServicesWMINamespace

    intFunctionReturn = 0
    intReturnMultiplier = 8

    intReturnCode = ConnectLocalWMINamespace(objSWbemServicesWMINamespace, Null, Null)
    If intReturnCode <> 0 Then
        intFunctionReturn = intFunctionReturn + (intReturnCode * intReturnMultiplier)
    Else
        intReturnCode = GetComputerSystemProductInstancesUsingWMINamespace(arrComputerSystemProductInstances, objSWbemServicesWMINamespace)
        intReturnMultiplier = 1
        intFunctionReturn = intFunctionReturn + (intReturnCode * intReturnMultiplier)
    End If
    
    GetComputerSystemProductInstances = intFunctionReturn
End Function
