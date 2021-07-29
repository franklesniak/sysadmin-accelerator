Function GetComputerUptimeInSeconds(ByRef intSecondsSinceLastBoot)
    'region FunctionMetadata ####################################################
    ' This function obtains the number of seconds since the computer was last booted. In other
    ' words, it obtains the computer's uptime.
    '
    ' The function takes one positional argument (intSecondsSinceLastBoot), which is populated
    ' upon success with an integer indicating the number of seconds since the computer was last
    ' booted
    '
    ' The function returns a 0 if the number of seconds since the computer's last boot was
    ' obtained successfully (as an integer). It returns a negative integer if an error occurred
    ' retrieving it. Finally, it returns a positive integer if the number of seconds since the
    ' last boot was obtained, but multiple operating system instances were present that
    ' contained data for the last boot date string. When this happens, only the first
    ' Win32_OperatingSystem instance containing data for the last boot date string is used to
    ' determine the number of seconds of uptime.
    '
    ' Example:
    '   intReturnCode = GetComputerUptimeInSeconds(intSecondsSinceLastBoot)
    '   If intReturnCode >= 0 Then
    '       ' The number of seconds since last boot (system uptime) was retrieved successfully
    '       ' and is stored in intSecondsSinceLastBoot
    '   End If
    '
    ' Version: 1.0.20210729.0
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
    ' GetComputerUptimeInSecondsUsingCurrentDateTimeComputerSystemInstancesOperatingSystemInstancesAndTimeZoneInstances()
    'endregion DependsOn ####################################################

    Dim intFunctionReturn
    Dim intReturnMultiplier
    Dim intReturnCode

    Dim objSWbemServicesWMINamespace
    Dim arrComputerSystemInstances
    Dim arrOperatingSystemInstances
    Dim arrTimeZoneInstances
    Dim intResultToReturn

    intFunctionReturn = 0
    intReturnMultiplier = 4194304 * 4
    intResultToReturn = Null

    intReturnCode = ConnectLocalWMINamespace(objSWbemServicesWMINamespace, Null, Null)
    If intReturnCode <> 0 Then
        intFunctionReturn = intReturnCode * intReturnMultiplier
    Else
        ' Successfully connected to the local computer's root\CIMv2 WMI Namespace
        intReturnCode = GetComputerSystemInstancesUsingWMINamespace(arrComputerSystemInstances, objSWbemServicesWMINamespace)
        intReturnCode = GetOperatingSystemInstancesUsingWMINamespace(arrOperatingSystemInstances, objSWbemServicesWMINamespace)
        intReturnCode = GetTimeZoneInstancesUsingWMINamespace(arrTimeZoneInstances, objSWbemServicesWMINamespace)
        intFunctionReturn = GetComputerUptimeInSecondsUsingCurrentDateTimeComputerSystemInstancesOperatingSystemInstancesAndTimeZoneInstances(intResultToReturn, Null, arrComputerSystemInstances, arrOperatingSystemInstances, arrTimeZoneInstances)
    End If

    If intFunctionReturn >= 0 Then
        intSecondsSinceLastBoot = intResultToReturn
    End If
    
    GetComputerUptimeInSeconds = intFunctionReturn
End Function
