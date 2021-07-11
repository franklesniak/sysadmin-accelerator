Function GetTimeZoneInstances(ByRef arrTimeZoneInstances)
    'region FunctionMetadata ####################################################
    ' This function retrieves the Win32_TimeZone instances and stores them in
    ' arrTimeZoneInstances.
    '
    ' The function takes one positional argument (arrTimeZoneInstances), which is populated
    ' upon success with the time zone instances returned from WMI of type Win32_TimeZone
    '
    ' The function returns 0 if Win32_TimeZone instances were retrieved successfully, and
    ' there was one time zone instance (as expected). If no Win32_TimeZone objects could be
    ' retrieved, then the function returns a negative number. If there are unexpectedly
    ' multiple instances of Win32_TimeZone, then the function returns a positive number equal
    ' to the number of WMI instances retrieved minus one.
    '
    ' Example:
    '   intReturnCode = GetTimeZoneInstances(arrTimeZoneInstances)
    '   If intReturnCode = 0 Then
    '       ' The Win32_TimeZone instance was retrieved successfully and is available
    '       ' at arrTimeZoneInstances.ItemIndex(0)
    '   ElseIf intReturnCode > 0 Then
    '       ' More than one Win32_TimeZone instance was retrieved, which is
    '       ' unexpected.
    '   Else
    '       ' An error occurred and no Win32_TimeZone instances were retrieved
    '   End If
    '
    ' Version: 1.0.20210708.0
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
    ' GetTimeZoneInstancesUsingWMINamespace()
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
        intReturnCode = GetTimeZoneInstancesUsingWMINamespace(arrTimeZoneInstances, objSWbemServicesWMINamespace)
        intReturnMultiplier = 1
        intFunctionReturn = intFunctionReturn + (intReturnCode * intReturnMultiplier)
    End If
    
    GetTimeZoneInstances = intFunctionReturn
End Function
