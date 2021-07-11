Function GetUTCOffsetForDateInLocalTimeZoneUsingTimeZoneInstances(ByRef intUTCOffset, ByVal datetimeToCheck, ByVal arrTimeZoneInstances)
    'region FunctionMetadata ####################################################
    ' Assuming that datetimeToCheck represents a VBScript-native datetime object, this function
    ' obtains the UTC offset for the specified date and time. The ambiguous hour surrounding a
    ' change to standard time or a change to daylight time is ignored
    '
    ' The function takes two positional arguments:
    '  - The first argument (intUTCOffset) is populated upon success with an integer that
    '    indicates the number of minutes that a time zone would be offset from UTC, given the
    '    specified date. For example, Central (US) Daylight Time would have an offset of -360
    '    (i.e., GMT-6)
    '  - The second argument (datetimeToCheck) is a VBScript-native datetime object
    '    (VT_DATETIME) that represents the date and time for which the function obtains the UTC
    '    offset
    '  - The third argument (arrTimeZoneInstances) is an array/collection of objects of class
    '    Win32_TimeZone
    '
    ' The function returns a 0 if the UTC offset was obtained successfully. It returns a
    ' negative integer if an error occurred retrieving the UTC offset
    '
    ' Example:
    '   intReturnCode = GetTimeZoneInstances(arrTimeZoneInstances)
    '   If intReturnCode >= 0 Then
    '       ' At least one Win32_TimeZone instance was retrieved successfully
    '       intReturnCode = GetUTCOffsetForDateInLocalTimeZoneUsingTimeZoneInstances(intUTCOffset, datetimeToCheck, arrTimeZoneInstances)
    '       If intReturnCode = 0 Then
    '           ' The function successfully retrieved the UTC offset for the current computer's
    '           ' time zone, on the date specified by datetimeToCheck
    '       Else
    '           ' An error occurred
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
    ' Rob Van Der Woude, for writing the isDST() function that loosely inspired this function:
    ' https://www.robvanderwoude.com/files/isdst_vbs.txt
    'endregion Acknowledgements ####################################################

    'region DependsOn ####################################################
    ' TestObjectIsDateTimeContainingData()
    ' GetStandardTimeZoneInfoUsingTimeZoneInstances()
    ' GetDaylightSavingsTimeZoneInfoUsingTimeZoneInstances()
    ' GetDayOfMonthFromDayOfWeekAndWeekOfMonth()
    'endregion DependsOn ####################################################

    Dim intFunctionReturn
    Dim intReturnMultiplier
    Dim intReturnCode
    Dim intUTCOffsetToReturn

    Dim strStandardTimeZoneName
    Dim intStandardTimeOffsetFromUTC
    Dim intStandardTimeStartYear
    Dim intStandardTimeStartMonth
    Dim intStandardTimeNthDayOfWeekInMonth
    Dim intStandardTimeDayOfWeek
    Dim intStandardTimeHour
    Dim intStandardTimeMinute
    Dim intStandardTimeSecond
    Dim intStandardTimeMillisecond

    Dim strDaylightTimeZoneName
    Dim intDaylightTimeOffsetFromUTC
    Dim intDaylightTimeStartYear
    Dim intDaylightTimeStartMonth
    Dim intDaylightTimeNthDayOfWeekInMonth
    Dim intDaylightTimeDayOfWeek
    Dim intDaylightTimeHour
    Dim intDaylightTimeMinute
    Dim intDaylightTimeSecond
    Dim intDaylightTimeMillisecond

    Dim boolStandardTimeStartDateSpecified
    Dim boolDaylightTimeStartDateSpecified

    Dim intYearFromDate
    Dim intTemp

    Dim datetimeStandardTimeTransition
    Dim datetimeDaylightTimeTransition

    intFunctionReturn = 0
    intReturnMultiplier = 128 * 8 * 2

    If TestObjectIsDateTimeContainingData(datetimeToCheck) <> True Then
        intFunctionReturn = intFunctionReturn + (-1 * intReturnMultiplier)
    Else
        intReturnCode = GetStandardTimeZoneInfoUsingTimeZoneInstances(strStandardTimeZoneName, intStandardTimeOffsetFromUTC, intStandardTimeStartYear, intStandardTimeStartMonth, intStandardTimeNthDayOfWeekInMonth, intStandardTimeDayOfWeek, intStandardTimeHour, intStandardTimeMinute, intStandardTimeSecond, intStandardTimeMillisecond, arrTimeZoneInstances)
        If intReturnCode < 0 Then
            intFunctionReturn = intReturnCode
        Else
            intReturnCode = GetDaylightSavingsTimeZoneInfoUsingTimeZoneInstances(strDaylightTimeZoneName, intDaylightTimeOffsetFromUTC, intDaylightTimeStartYear, intDaylightTimeStartMonth, intDaylightTimeNthDayOfWeekInMonth, intDaylightTimeDayOfWeek, intDaylightTimeHour, intDaylightTimeMinute, intDaylightTimeSecond, intDaylightTimeMillisecond, arrTimeZoneInstances)
            If intReturnCode < 0 Then
                intFunctionReturn = intReturnCode * 2
            Else
                If intDaylightTimeStartYear = 0 And intDaylightTimeStartMonth = 0 And intDaylightTimeNthDayOfWeekInMonth = 0 And intDaylightTimeDayOfWeek = 0 And intDaylightTimeHour = 0 And intDaylightTimeMinute = 0 And intDaylightTimeSecond = 0 And intDaylightTimeMillisecond = 0 Then
                    boolDaylightTimeStartDateSpecified = False
                    datetimeDaylightTimeTransition = Null
                Else
                    boolDaylightTimeStartDateSpecified = True
                End If

                If boolDaylightTimeStartDateSpecified = False Then
                    ' There is no daylight time in this time zone
                    intUTCOffsetToReturn = intStandardTimeOffsetFromUTC
                Else
                    If intStandardTimeStartYear = 0 And intStandardTimeStartMonth = 0 And intStandardTimeNthDayOfWeekInMonth = 0 And intStandardTimeDayOfWeek = 0 And intStandardTimeHour = 0 And intStandardTimeMinute = 0 And intStandardTimeSecond = 0 And intStandardTimeMillisecond = 0 Then
                        boolStandardTimeStartDateSpecified = False
                        datetimeStandardTimeTransition = Null
                    Else
                        boolStandardTimeStartDateSpecified = True
                    End If

                    On Error Resume Next
                    intYearFromDate = DatePart("yyyy", datetimeToCheck)
                    If Err Then
                        On Error Goto 0
                        Err.Clear
                        intFunctionReturn = intFunctionReturn + (-2 * intReturnMultiplier)
                    Else
                        On Error Goto 0
                        ' Daylight time was specified
                        intTemp = GetDayOfMonthFromDayOfWeekAndWeekOfMonth(intYearFromDate, intDaylightTimeStartMonth, intDaylightTimeNthDayOfWeekInMonth, intDaylightTimeDayOfWeek)
                        If intTemp = 0 Then
                            intFunctionReturn = intFunctionReturn + (-3 * intReturnMultiplier)
                        Else
                            ' intTemp contains the day of the month intDaylightTimeStartMonth
                            ' in which the transition to daylight savings starts
                            On Error Resume Next
                            datetimeDaylightTimeTransition = DateSerial(intYearFromDate, intDaylightTimeStartMonth, intTemp) + TimeSerial(intDaylightTimeHour, intDaylightTimeMinute, intDaylightTimeSecond)
                            If Err Then
                                On Error Goto 0
                                Err.Clear
                                intFunctionReturn = intFunctionReturn + (-4 * intReturnMultiplier)
                            Else
                                On Error Goto 0
                            End If
                        End If

                        If intFunctionReturn = 0 Then
                            'No error occurred
                            If boolStandardTimeStartDateSpecified = True Then
                                intTemp = GetDayOfMonthFromDayOfWeekAndWeekOfMonth(intYearFromDate, intStandardTimeStartMonth, intStandardTimeNthDayOfWeekInMonth, intStandardTimeDayOfWeek)
                                If intTemp = 0 Then
                                    intFunctionReturn = intFunctionReturn + (-5 * intReturnMultiplier)
                                Else
                                    ' intTemp contains the day of the month
                                    ' intDaylightTimeStartMonth in which the transition to
                                    ' daylight savings starts
                                    On Error Resume Next
                                    datetimeStandardTimeTransition = DateSerial(intYearFromDate, intStandardTimeStartMonth, intTemp) + TimeSerial(intStandardTimeHour, intStandardTimeMinute, intStandardTimeSecond)
                                    If Err Then
                                        On Error Goto 0
                                        Err.Clear
                                        intFunctionReturn = intFunctionReturn + (-6 * intReturnMultiplier)
                                    Else
                                        On Error Goto 0
                                    End If
                                End If
                            End If
                        End If

                        If intFunctionReturn = 0 Then
                            'No error occurred
                            If TestObjectIsDateTimeContainingData(datetimeStandardTimeTransition) <> True Then
                                ' Standard time undefined; only daylight time defined
                                If datetimeToCheck < datetimeDaylightTimeTransition Then
                                    ' Assume date was standard time
                                    intUTCOffsetToReturn = intStandardTimeOffsetFromUTC
                                Else
                                    intUTCOffsetToReturn = intDaylightTimeOffsetFromUTC
                                End If
                            Else
                                ' Standard time and daylight time both defined
                                If datetimeDaylightTimeTransition < datetimeStandardTimeTransition Then
                                    ' Year: Standard Time -> Daylight Time -> Standard Time
                                    If datetimeToCheck < datetimeDaylightTimeTransition Then
                                        ' Date was standard time
                                        intUTCOffsetToReturn = intStandardTimeOffsetFromUTC
                                    Else
                                        If datetimeToCheck < datetimeStandardTimeTransition Then
                                            ' Date was daylight time
                                            intUTCOffsetToReturn = intDaylightTimeOffsetFromUTC
                                        Else
                                            intUTCOffsetToReturn = intStandardTimeOffsetFromUTC
                                        End If
                                    End If
                                Else
                                    ' Year: Daylight Time -> Standard Time -> Daylight Time
                                    If datetimeToCheck < datetimeStandardTimeTransition Then
                                        ' Date was daylight time
                                        intUTCOffsetToReturn = intDaylightTimeOffsetFromUTC
                                    Else
                                        If datetimeToCheck < datetimeDaylightTimeTransition Then
                                            ' Date was standard time
                                            intUTCOffsetToReturn = intStandardTimeOffsetFromUTC
                                        Else
                                            ' Date was daylight time
                                            intUTCOffsetToReturn = intDaylightTimeOffsetFromUTC
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If

    If intFunctionReturn = 0 Then
        intUTCOffset = intUTCOffsetToReturn
    End If

    GetUTCOffsetForDateInLocalTimeZoneUsingTimeZoneInstances = intFunctionReturn
End Function
