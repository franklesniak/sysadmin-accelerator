Function GetSystemEnclosureInstancesUsingWMINamespace(ByRef arrSystemEnclosureInstances, ByVal objWMINamespace)
    'region FunctionMetadata ####################################################
    ' Assuming that objWMINamespace represents a successful connection to the root\CIMv2
    ' WMI namespace, this function retrieves the Win32_SystemEnclosure instances and stores
    ' them in arrSystemEnclosureInstances.
    '
    ' The function takes two positional arguments:
    '  - The first argument (arrSystemEnclosureInstances) is populated upon success with the
    '    system enclosure instances returned from WMI of type Win32_SystemEnclosure
    '  - The second argument (objWMINamespace) is a WMI Namespace connection argument that must
    '    already be connected to the WMI namespace root\CIMv2
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
    '   intReturnCode = ConnectLocalWMINamespace(objSWbemServicesWMINamespace, Null, Null)
    '   If intReturnCode = 0 Then
    '       ' Successfully connected to the local computer's root\CIMv2 WMI Namespace
    '       intReturnCode = GetSystemEnclosureInstancesUsingWMINamespace(arrSystemEnclosureInstances, objSWbemServicesWMINamespace)
    '       If intReturnCode > 0 Then
    '           ' One or more Win32_SystemEnclosure instances were retrieved. The first
    '           ' instance is available at arrSystemEnclosureInstances.ItemIndex(0) and the
    '           ' number of instances is available at arrSystemEnclosureInstances.Count. In
    '           ' other words, the upper array boundary/index is
    '           ' (arrSystemEnclosureInstances.Count - 1).
    '       Else
    '           ' No Win32_SystemEnclosure instances were retrieved
    '       End If
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
    ' TestObjectForData()
    ' TestObjectIsAnyTypeOfInteger()
    'endregion DependsOn ####################################################

    Dim intFunctionReturn
    Dim intReturnMultiplier
    Dim arrWorkingSystemEnclosureInstances
    Dim intTemp

    Err.Clear

    intFunctionReturn = 0
    intReturnMultiplier = 1

    If TestObjectForData(objWMINamespace) <> True Then
        intFunctionReturn = intFunctionReturn + (-2 * intReturnMultiplier)
    Else
        On Error Resume Next
        Set arrWorkingSystemEnclosureInstances = objWMINamespace.InstancesOf("Win32_SystemEnclosure")
        If Err Then
            On Error Goto 0
            Err.Clear
            intFunctionReturn = intFunctionReturn + (-3 * intReturnMultiplier)
        Else
            intTemp = arrWorkingSystemEnclosureInstances.Count
            If Err Then
                On Error Goto 0
                Err.Clear
                intFunctionReturn = intFunctionReturn + (-4 * intReturnMultiplier)
            Else
                On Error Goto 0
                If TestObjectIsAnyTypeOfInteger(intTemp) = False Then
                    intFunctionReturn = intFunctionReturn + (-5 * intReturnMultiplier)
                Else
                    If intTemp < 0 Then
                        intFunctionReturn = intFunctionReturn + (-6 * intReturnMultiplier)
                    Else
                        ' intTemp >= 0
                        intFunctionReturn = intTemp
                        ' 0 would be returned if there are no Win32_SystemEnclosure instances
                    End If
                End If
            End If
        End If
    End If

    If intFunctionReturn > 0 Then
        On Error Resume Next
        Set arrSystemEnclosureInstances = objWMINamespace.InstancesOf("Win32_SystemEnclosure")
        If Err Then
            On Error Goto 0
            Err.Clear
            intFunctionReturn = (-7 * intReturnMultiplier)
        Else
            On Error Goto 0
        End If
    End If
    
    GetSystemEnclosureInstancesUsingWMINamespace = intFunctionReturn
End Function
