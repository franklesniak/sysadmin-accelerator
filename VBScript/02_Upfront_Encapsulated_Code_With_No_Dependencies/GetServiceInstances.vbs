Function GetServiceInstances(ByRef arrServiceInstances)
    'region FunctionMetadata #######################################################
    ' This function retrieves Win32_Service instances and stores them in
    ' arrServiceInstances.
    '
    ' The function takes one positional argument (arrServiceInstances), which is
    ' populated upon success with the Service instances returned from WMI of type
    ' Win32_Service
    '
    ' If Win32_Service instances were retrieved successfully, the function returns an
    ' integer equal to the number of Service instances. If no Win32_Service objects
    ' could be retrieved, then the function returns a negative number.
    '
    ' Example:
    '   intReturnCode = GetServiceInstances(arrServiceInstances)
    '  If intReturnCode > 0 Then
    '       ' One or more Win32_Service instance were retrieved successfully. The first
    '       ' instance is available at arrServiceInstances.ItemIndex(0)
    '   Else
    '       ' An error occurred and no Win32_Service instances were retrieved
    '   End If
    '
    ' Version: 1.0.20230731.0
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

    'region DependsOn ####################################################
    ' ConnectLocalWMINamespace()
    ' GetServiceInstancesUsingWMINamespace()
    'endregion DependsOn ####################################################

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
        intReturnCode = GetServiceInstancesUsingWMINamespace(arrServiceInstances, objSWbemServicesWMINamespace)
        intReturnMultiplier = 1
        intFunctionReturn = intFunctionReturn + (intReturnCode * intReturnMultiplier)
    End If
    
    GetServiceInstances = intFunctionReturn
End Function
