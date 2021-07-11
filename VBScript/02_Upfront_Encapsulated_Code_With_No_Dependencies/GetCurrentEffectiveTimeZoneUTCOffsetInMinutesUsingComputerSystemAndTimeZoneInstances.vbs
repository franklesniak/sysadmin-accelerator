Function GetCurrentEffectiveTimeZoneUTCOffsetInMinutesUsingComputerSystemAndTimeZoneInstances(ByRef intCurrentEffectiveTimeZoneUTCOffsetInMinutes, ByVal arrComputerSystemInstances, ByVal arrTimeZoneInstances)
    'region FunctionMetadata ####################################################
    ' Assuming arrComputerSystemInstances is populated with instances of the WMI
    ' Win32_ComputerSystem class and/or arrTimeZoneInstances is populated with instances of the
    ' WMI Win32_TimeZone class, this function obtains the computer's current effective time
    ' zone UTC offset (in minutes). For example, for a computer in Central (US) Standard Time
    ' (CST), the time zone UTC offset would be -360 because CST is GMT-6.
    '
    ' The function takes three positional arguments:
    '   The first argument (intCurrentEffectiveTimeZoneUTCOffsetInMinutes) is populated upon
    '       success with a string containing the computer's current effective time zone UTC
    '       offset (in minutes) as reported by WMI.
    '   The second argument (arrComputerSystemInstances) is an array/collection of objects of
    '       class Win32_ComputerSystem
    '   The third argument (arrTimeZoneInstances) is an array/collection of objects of class
    '       Win32_TimeZone
    '
    ' The function returns a 0 or a positive number if the current effective time zone UTC
    ' offset (in minutes) was obtained successfully. A zero indicates that the preferred
    ' Win32_ComputerSystem approach worked. A one indicates that the less-preferred
    ' Win32_TimeZone approach worked successfully. The function returns a negative integer if
    ' an error occurred retrieving the time zone offset.
    '
    ' Example:
    '   intReturnCode = GetComputerSystemInstances(arrComputerSystemInstances)
    '   If intReturnCode >= 0 Then
    '       ' At least one Win32_ComputerSystem instance was retrieved successfully
    '       intReturnCode = GetTimeZoneInstances(arrTimeZoneInstances)
    '       If intReturnCode >= 0 Then
    '           ' At least one Win32_TimeZone instance was retrieved successfully
    '           intReturnCode = GetCurrentEffectiveTimeZoneUTCOffsetInMinutesUsingComputerSystemAndTimeZoneInstances(intCurrentEffectiveTimeZoneUTCOffsetInMinutes, arrComputerSystemInstances, arrTimeZoneInstances)
    '           If intReturnCode >= 0 Then
    '               ' The computer's current effective time zone UTC offset (in minutes) was retrieved
    '               ' successfully and is stored in intCurrentEffectiveTimeZoneUTCOffsetInMinutes
    '           End If
    '       End If
    '   End If
    '
    ' Version: 1.0.20210710.0
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
    ' GetCurrentEffectiveTimeZoneUTCOffsetInMinutesUsingComputerSystemInstances()
    ' GetCurrentEffectiveTimeZoneUTCOffsetInMinutesUsingTimeZoneInstances()
    'endregion DependsOn ####################################################

    Dim intFunctionReturn
    Dim intResult
    Dim intReturnCode

    intFunctionReturn = 0

    If TestObjectForData(arrComputerSystemInstances) <> True Then
        intFunctionReturn = -1
    Else
        intFunctionReturn = GetCurrentEffectiveTimeZoneUTCOffsetInMinutesUsingComputerSystemInstances(intResult, arrComputerSystemInstances)
        If intFunctionReturn >= 0 Then
            ' The computer's current effective time zone UTC offset (in minutes) was retrieved
            ' successfully and is stored in intResult
            intCurrentEffectiveTimeZoneUTCOffsetInMinutes = intResult
        End If
    End If

    If intFunctionReturn >= 0 Then
        intFunctionReturn = 0
    Else
        ' An error occurred trying the ComputerSystem method. Try the TimeZone method.
        If TestObjectForData(arrTimeZoneInstances) <> True Then
            intReturnCode = -1
        Else
            intReturnCode = GetCurrentEffectiveTimeZoneUTCOffsetInMinutesUsingTimeZoneInstances(intResult, arrTimeZoneInstances)
            If intReturnCode >= 0 Then
                ' The computer's current effective time zone UTC offset (in minutes) was
                ' retrieved successfully and is stored in intResult
                intCurrentEffectiveTimeZoneUTCOffsetInMinutes = intResult
            End If
        End If
        If intReturnCode < 0 Then
            intFunctionReturn = intFunctionReturn + (8 * 128 * intReturnCode)
        Else
            ' This method succeeded
            intFunctionReturn = 1
        End If
    End If
    
    GetCurrentEffectiveTimeZoneUTCOffsetInMinutesUsingComputerSystemAndTimeZoneInstances = intFunctionReturn
End Function
