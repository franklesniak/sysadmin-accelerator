Function GetSystemEnclosureInstances(ByRef arrSystemEnclosureInstances)
    'region FunctionMetadata ####################################################
    ' This function retrieves the Win32_SystemEnclosure instances and stores them in
    ' arrSystemEnclosureInstances.
    '
    ' The function takes one positional argument (arrSystemEnclosureInstances), which is
    ' populated upon success with the system enclosure instances returned from WMI of type
    ' Win32_SystemEnclosure
    '
    ' If Win32_SystemEnclosure instances were retrieved successfully, then the function returns
    ' a positive integer equal to the number of Win32_SystemEnclosure instances retrieved.
    ' Usually there is one Win32_SystemEnclosure, and this function therefore returns 1.
    ' However, in some circumstances (e.g., a docking station attached to a laptop/tablet),
    ' more than one Win32_SystemEnclosure instance is present. In these circumstances, the
    ' function would return 2, 3, etc. If an error occurs, the function returns a negative
    ' integer. If no error occurred but no instances could be retrieved, then the function
    ' returns zero.
    '
    ' Example:
    '   intReturnCode = GetSystemEnclosureInstances(arrSystemEnclosureInstances)
    '   If intReturnCode > 0 Then
    '       ' One or more Win32_SystemEnclosure instances were retrieved. The first
    '       ' instance is available at arrSystemEnclosureInstances.ItemIndex(0) and the
    '       ' number of instances is available at arrSystemEnclosureInstances.Count. In
    '       ' other words, the upper array boundary/index is
    '       ' (arrSystemEnclosureInstances.Count - 1).
    '   Else
    '       ' No Win32_SystemEnclosure instances were retrieved
    '   End If
    '
    ' Version: 1.0.20210624.0
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
    ' None!
    'endregion Acknowledgements ####################################################

    'region DependsOn ####################################################
    ' ConnectLocalWMINamespace()
    ' GetSystemEnclosureInstancesUsingWMINamespace()
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
        intReturnCode = GetSystemEnclosureInstancesUsingWMINamespace(arrSystemEnclosureInstances, objSWbemServicesWMINamespace)
        intReturnMultiplier = 1
        intFunctionReturn = intFunctionReturn + (intReturnCode * intReturnMultiplier)
    End If
    
    GetSystemEnclosureInstances = intFunctionReturn
End Function
