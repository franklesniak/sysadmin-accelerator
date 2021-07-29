Function GetComputerUptimeInSecondsUsingCurrentDateTimeComputerSystemInstancesOperatingSystemInstancesAndTimeZoneInstances(ByRef intSecondsSinceLastBoot, ByVal datetimeNow, ByVal arrComputerSystemInstances, ByVal arrOperatingSystemInstances, ByVal arrTimeZoneInstances)
    'region FunctionMetadata ####################################################
    ' Assuming that arrComputerSystemInstances represents an array / collection of the
    ' available computer system instances (of type Win32_ComputerSystem),
    ' arrOperatingSystemInstances represents an array / collection of the available operating
    ' system instances (of type Win32_OperatingSystem), and arrTimeZoneInstances represents an
    ' array / collection of the available time zone instances (of type Win32_TimeZone), this
    ' function obtains the number of seconds since the computer was last booted. In other
    ' words, it obtains the computer's uptime.
    '
    ' The function takes five positional arguments:
    '  - The first argument (intSecondsSinceLastBoot) is populated upon success with an integer
    '    indicating the number of seconds since the computer was last booted
    '  - The second argument (datetimeNow) is a VBScript-native datetime object (VT_DATE) that
    '    represents the current day and time. Normally it would be set to the equivalent of
    '    Now(). If no data is supplied for this argument (e.g., Null or a variable set to
    '    Nothing is passed, the function defaults to using the current datetime.
    '  - The third argument (arrComputerSystemInstances) is an array/collection of objects of
    '    class Win32_ComputerSystem
    '  - The fourth argument (arrOperatingSystemInstances) is an array/collection of objects of
    '    class Win32_OperatingSystem
    '  - The fifth argument (arrTimeZoneInstances) is an array/collection of objects of class
    '    Win32_TimeZone
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
    '   intReturnCode = ConnectLocalWMINamespace(objSWbemServicesWMINamespace, Null, Null)
    '   If intReturnCode = 0 Then
    '       ' Successfully connected to the local computer's root\CIMv2 WMI Namespace
    '       intReturnCode = GetComputerSystemInstancesUsingWMINamespace(arrComputerSystemInstances, objSWbemServicesWMINamespace)
    '       If intReturnCode >= 0 Then
    '           ' At least one Win32_ComputerSystem instance was retrieved successfully
    '           intReturnCode = GetOperatingSystemInstancesUsingWMINamespace(arrOperatingSystemInstances, objSWbemServicesWMINamespace)
    '           If intReturnCode >= 0 Then
    '           ' At least one Win32_OperatingSystem instance was retrieved successfully
    '               intReturnCode = GetTimeZoneInstancesUsingWMINamespace(arrTimeZoneInstances, objSWbemServicesWMINamespace)
    '               If intReturnCode >= 0 Then
    '                   ' At least one Win32_TimeZone instance was retrieved successfully
    '                   datetimeNow = Now()
    '                   intReturnCode = GetComputerUptimeInSecondsUsingCurrentDateTimeComputerSystemInstancesOperatingSystemInstancesAndTimeZoneInstances(intSecondsSinceLastBoot, datetimeNow, arrComputerSystemInstances, arrOperatingSystemInstances, arrTimeZoneInstances)
    '                   If intReturnCode >= 0 Then
    '                       ' The number of seconds since last boot (system uptime) was
    '                       ' retrieved successfully and is stored in intSecondsSinceLastBoot
    '                   End If
    '               End If
    '           End If
    '       End If
    '   End If
    '
    ' Version: 1.0.20210728.0
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
    ' TestObjectIsDateTimeContainingData()
    ' GetComputerLastBootAsNativeDatetimeObjectUsingComputerSystemOperatingSystemAndTimeZoneInstances()
    'endregion DependsOn ####################################################

    Dim intFunctionReturn
    Dim intReturnMultiplier
    Dim intReturnCode

    Dim intResultToReturn
    Dim datetimeLastBootDate
    Dim dateTimeWorkingNow

    Err.Clear

    intFunctionReturn = 0
    intReturnMultiplier = 4194304
    intResultToReturn = Null

    intFunctionReturn = GetComputerLastBootAsNativeDatetimeObjectUsingComputerSystemOperatingSystemAndTimeZoneInstances(datetimeLastBootDate, arrComputerSystemInstances, arrOperatingSystemInstances, arrTimeZoneInstances)
    If intFunctionReturn >= 0 Then
        ' One or more Win32_OperatingSystem instances had a valid last boot date
        If TestObjectIsDateTimeContainingData(datetimeNow) <> True Then
            On Error Resume Next
            dateTimeWorkingNow = Now()
            If Err Then
                On Error Goto 0
                Err.Clear
                intFunctionReturn = -1 * intReturnMultiplier
            Else
                On Error Goto 0
            End If
        Else
            dateTimeWorkingNow = datetimeNow
        End If
        If intFunctionReturn >= 0 Then
            On Error Resume Next
            intResultToReturn = DateDiff("s", datetimeLastBootDate, dateTimeWorkingNow)
            If Err Then
                On Error Goto 0
                Err.Clear
                intFunctionReturn = -2 * intReturnMultiplier
            Else
                On Error Goto 0
            End if
        End If
    End If

    If intFunctionReturn >= 0 Then
        intSecondsSinceLastBoot = intResultToReturn
    End If
    
    GetComputerUptimeInSecondsUsingCurrentDateTimeComputerSystemInstancesOperatingSystemInstancesAndTimeZoneInstances = intFunctionReturn
End Function
