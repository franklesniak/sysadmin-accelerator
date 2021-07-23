Function GetOperatingSystemInstancesUsingWMINamespace(ByRef arrOperatingSystemInstances, ByVal objWMINamespace)
    'region FunctionMetadata ####################################################
    ' Assuming that objWMINamespace represents a successful connection to the root\CIMv2
    ' WMI namespace, this function retrieves the Win32_OperatingSystem instances and stores
    ' them in arrOperatingSystemInstances.
    '
    ' The function takes two positional arguments:
    '  - The first argument (arrOperatingSystemInstances) is populated upon success with the
    '    computer system instances returned from WMI of type Win32_OperatingSystem
    '  - The second argument (objWMINamespace) is a WMI Namespace connection argument that must
    '    already be connected to the WMI namespace root\CIMv2
    '
    ' The function returns 0 if Win32_OperatingSystem instances were retrieved successfully,
    ' and there was one computer system instance (as expected). If no Win32_OperatingSystem
    ' objects could be retrieved, then the function returns a negative number. If there are
    ' unexpectedly multiple instances of Win32_OperatingSystem, then the function returns a
    ' positive number equal to the number of WMI instances retrieved minus one.
    '
    ' Example:
    '   intReturnCode = ConnectLocalWMINamespace(objSWbemServicesWMINamespace, Null, Null)
    '   If intReturnCode = 0 Then
    '       ' Successfully connected to the local computer's root\CIMv2 WMI Namespace
    '       intReturnCode = GetOperatingSystemInstancesUsingWMINamespace(arrOperatingSystemInstances, objSWbemServicesWMINamespace)
    '       If intReturnCode = 0 Then
    '           ' The Win32_OperatingSystem instance was retrieved successfully and is
    '           ' available at arrOperatingSystemInstances.ItemIndex(0)
    '       ElseIf intReturnCode > 0 Then
    '           ' More than one Win32_OperatingSystem instance was retrieved, which is
    '           ' unexpected.
    '       Else
    '           ' An error occurred and no Win32_OperatingSystem instances were retrieved
    '       End If
    '   End If
    '
    ' Version: 1.0.20210722.1
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
    Dim arrWorkingOperatingSystemInstances
    Dim intTemp

    Err.Clear

    intFunctionReturn = 0
    intReturnMultiplier = 1

    If TestObjectForData(objWMINamespace) <> True Then
        intFunctionReturn = intFunctionReturn + (-2 * intReturnMultiplier)
    Else
        On Error Resume Next
        Set arrWorkingOperatingSystemInstances = objWMINamespace.InstancesOf("Win32_OperatingSystem")
        If Err Then
            On Error Goto 0
            Err.Clear
            intFunctionReturn = intFunctionReturn + (-3 * intReturnMultiplier)
        Else
            intTemp = arrWorkingOperatingSystemInstances.Count
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
                        intFunctionReturn = intTemp - 1
                        ' -1 would be returned if there are no Win32_OperatingSystem instances
                    End If
                End If
            End If
        End If
    End If

    If intFunctionReturn >= 0 Then
        On Error Resume Next
        Set arrOperatingSystemInstances = objWMINamespace.InstancesOf("Win32_OperatingSystem")
        If Err Then
            On Error Goto 0
            Err.Clear
            intFunctionReturn = (-7 * intReturnMultiplier)
        Else
            On Error Goto 0
        End If
    End If
    
    GetOperatingSystemInstancesUsingWMINamespace = intFunctionReturn
End Function
