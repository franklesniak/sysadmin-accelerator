Function GetCurrentEffectiveTimeZoneUTCOffsetInMinutesUsingTimeZoneInstances(ByRef intCurrentEffectiveTimeZoneUTCOffsetInMinutes, ByVal arrTimeZoneInstances)
    'region FunctionMetadata ####################################################
    ' Assuming that arrTimeZoneInstances represents an array / collection of the
    ' available time zone instances (of type Win32_TimeZone), this function obtains
    ' the computer's current effective time zone UTC offset (in minutes), if available. For
    ' example, for a computer in Central (US) Standard Time (CST), the time zone UTC offset
    ' would be -360 because CST is GMT-6.
    '
    ' The function takes two positional arguments:
    '  - The first argument (intCurrentEffectiveTimeZoneUTCOffsetInMinutes) is populated upon
    '    success with a string containing the computer's current effective time zone UTC offset
    '    (in minutes) as reported by WMI.
    '  - The second argument (arrTimeZoneInstances) is an array/collection of objects of
    '    class Win32_TimeZone
    '
    ' The function returns a 0 if the current effective time zone UTC offset (in minutes) was
    ' obtained successfully. It returns a negative integer if an error occurred retrieving the
    ' time zone offset. Finally, it returns a positive integer if the time zone offset was
    ' obtained, but multiple time zone instances were present that contained data for the
    ' time zone offset. When this happens, only the first Win32_TimeZone instance
    ' containing data for the time zone offset is used.
    '
    ' Example:
    '   intReturnCode = GetTimeZoneInstances(arrTimeZoneInstances)
    '   If intReturnCode >= 0 Then
    '       ' At least one Win32_TimeZone instance was retrieved successfully
    '       intReturnCode = GetCurrentEffectiveTimeZoneUTCOffsetInMinutesUsingTimeZoneInstances(intCurrentEffectiveTimeZoneUTCOffsetInMinutes, arrTimeZoneInstances)
    '       If intReturnCode >= 0 Then
    '           ' The computer's current effective time zone UTC offset (in minutes) was
    '           ' retrieved successfully and is stored in
    '           ' intCurrentEffectiveTimeZoneUTCOffsetInMinutes
    '       End If
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
    ' TestObjectForData()
    ' GetUTCOffsetForDateInLocalTimeZoneUsingTimeZoneInstances()
    'endregion DependsOn ####################################################

    Dim intFunctionReturn
    Dim intReturnMultiplier
    Dim intReturnCode
    Dim intResultToReturn

    Err.Clear

    intFunctionReturn = 0
    intReturnMultiplier = 128 * 8 * 2 * 8
    intResultToReturn = Null

    If TestObjectForData(arrTimeZoneInstances) <> True Then
        intFunctionReturn = intFunctionReturn + (-1 * intReturnMultiplier)
    Else
        intFunctionReturn = GetUTCOffsetForDateInLocalTimeZoneUsingTimeZoneInstances(intResultToReturn, Now, arrTimeZoneInstances)
    End If

    If intFunctionReturn >= 0 Then
        intCurrentEffectiveTimeZoneUTCOffsetInMinutes = intResultToReturn
    End If
    
    GetCurrentEffectiveTimeZoneUTCOffsetInMinutesUsingTimeZoneInstances = intFunctionReturn
End Function
