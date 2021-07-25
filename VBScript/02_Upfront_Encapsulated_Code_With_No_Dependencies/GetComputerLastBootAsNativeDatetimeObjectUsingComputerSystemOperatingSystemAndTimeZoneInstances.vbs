Function GetComputerLastBootAsNativeDatetimeObjectUsingComputerSystemOperatingSystemAndTimeZoneInstances(ByRef datetimeLastBootDate, ByVal arrComputerSystemInstances, ByVal arrOperatingSystemInstances, ByVal arrTimeZoneInstances)
    'region FunctionMetadata ####################################################
    ' Assuming that arrComputerSystemInstances represents an array / collection of the
    ' available computer system instances (of type Win32_ComputerSystem),
    ' arrOperatingSystemInstances represents an array / collection of the available operating
    ' system instances (of type Win32_OperatingSystem), and arrTimeZoneInstances represents an
    ' array / collection of the available time zone instances (of type Win32_TimeZone), this
    ' function obtains the computer's last boot date and time as a local time zone-adjusted
    ' VBScript-native datetime object (i.e., VT_DATE object).
    '
    ' The function takes four positional arguments:
    '  - The first argument (datetimeLastBootDate) is populated upon success with a VBScript-
    '    native datetime object (i.e., a VT_DATE object) containing the computer's last boot
    '    date and time. The date and time are adjusted to this computer's local time zone. The
    '    computer's last boot date and time is equivalent to the Win32_OperatingSystem object
    '    property LastBootUpTime, but converted from CIM_TIMEDATE to VT_DATE
    '  - The second argument (arrComputerSystemInstances) is an array/collection of objects of
    '    class Win32_ComputerSystem
    '  - The third argument (arrOperatingSystemInstances) is an array/collection of objects of
    '    class Win32_OperatingSystem
    '  - The fourth argument (arrTimeZoneInstances) is an array/collection of objects of class
    '    Win32_TimeZone
    '
    ' The function returns a 0 if the computer's last boot date and time was obtained
    ' successfully as a VBScript-native datetime (i.e., VT_DATE) object. It returns a negative
    ' integer if an error occurred retrieving it. Finally, it returns a positive integer if the
    ' last boot date and time was obtained, but multiple operating system instances were
    ' present that contained data for the last boot date string. When this happens, only the
    ' first Win32_OperatingSystem instance containing data for the last boot date string is
    ' used.
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
    '                   intReturnCode = GetComputerLastBootAsNativeDatetimeObjectUsingComputerSystemOperatingSystemAndTimeZoneInstances(datetimeLastBootDate, arrComputerSystemInstances, arrOperatingSystemInstances, arrTimeZoneInstances)
    '                   If intReturnCode >= 0 Then
    '                       ' The last boot date/time was retrieved successfully in VBScript-
    '                       ' native VT_DATE format and is stored in datetimeLastBootDate
    '                   End If
    '               End If
    '           End If
    '       End If
    '   End If
    '
    ' Version: 1.0.20210723.0
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
    ' GetComputerLastBootCIMDATETIMEStringUsingOperatingSystemInstances()
    ' ConvertStringCIMDATETIMEToVTDATELocalTimeUsingComputerSystemAndTimeZoneInstances()
    'endregion DependsOn ####################################################

    Dim intFunctionReturn
    Dim intReturnMultiplier
    Dim intReturnCode

    Dim datetimeResultToReturn
    Dim strComputerLastBootDate

    Err.Clear

    intFunctionReturn = 0
    intReturnMultiplier = 128 * 8
    datetimeResultToReturn = Null

    intFunctionReturn = GetComputerLastBootCIMDATETIMEStringUsingOperatingSystemInstances(strComputerLastBootDate, arrOperatingSystemInstances)
    If intFunctionReturn >= 0 Then
        ' One or more Win32_OperatingSystem instances had a valid last boot date
        intReturnCode = ConvertStringCIMDATETIMEToVTDATELocalTimeUsingComputerSystemAndTimeZoneInstances(datetimeResultToReturn, strComputerLastBootDate, arrComputerSystemInstances, arrTimeZoneInstances)
        If intReturnCode >= 0 Then
            ' Conversion completed successfully
        Else
            ' Conversion failed
            intFunctionReturn = intFunctionReturn + (intReturnCode * intReturnMultiplier)
        End If
    End If

    If intFunctionReturn >= 0 Then
        datetimeLastBootDate = datetimeResultToReturn
    End If
    
    GetComputerLastBootAsNativeDatetimeObjectUsingComputerSystemOperatingSystemAndTimeZoneInstances = intFunctionReturn
End Function
