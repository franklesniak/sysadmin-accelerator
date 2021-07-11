Function GetStandardTimeZoneInfoUsingTimeZoneInstances(ByRef strStandardTimeZoneName, ByRef intStandardTimeOffsetFromUTC, ByRef intStartYear, ByRef intStartMonth, ByRef intNthDayOfWeekInMonth, ByRef intDayOfWeek, ByRef intHour, ByRef intMinute, ByRef intSecond, ByRef intMillisecond, ByVal arrTimeZoneInstances)
    'region FunctionMetadata ####################################################
    ' Assuming that arrTimeZoneInstances represents an array / collection of the available time
    ' zone instances (of type Win32_TimeZone), this function obtains the available metadata
    ' about standard time for the current time zone and returns it via a series of arguments
    '
    ' The function takes eleven positional arguments:
    '  - The first argument (strStandardTimeZoneName) is populated upon success with a string
    '    containing the name of the time zone when in standard time (e.g., "Central Standard
    '    Time"), as reported by WMI
    '  - The second argument (intStandardTimeOffsetFromUTC) is an integer that indicates the
    '    number of minutes that standard time is offset from UTC. For example, Central (US)
    '    Standard Time would have an offset of -360 (i.e., GMT-6).
    '  - The third argument (intStartYear) is an integer that indicates the year that standard
    '    time goes into effect. In the experience of the function author, this is always set to
    '    zero - meaning that it goes into effect every year.
    '  - The fourth argument (intStartMonth) is an integer that indicates the month number when
    '    standard time starts (or resumes, 1 = January)
    '  - The fifth argument (intNthDayOfWeekInMonth) is an integer that indicates the "n-th"
    '    day of the month when standard time starts (or resumes). A 5 indicates "the last"
    '    applicable day of the month. For example, if standard time starts on the fourth Sunday
    '    of October, then this argument would be set to 4. If standard time should start (or
    '    resume) on the last Sunday of October, then this argument would be set to 5
    '  - The sixth argument (intDayOfWeek) is an integer that indicates the day of the week
    '    that standard time starts (or resumes). The possible values are:
    '        0 - Sunday
    '        1 - Monday
    '        2 - Tuesday
    '        3 - Wednesday
    '        4 - Thursday
    '        5 - Friday
    '        6 - Saturday
    '  - The seventh argument (intHour) is an integer that indicates the hour of the day when
    '    the transition to standard time occurs. The hour is represented as a 24-hour clock, so
    '    7 PM would be 19
    '  - The eighth argument (intMinute) is an integer that indicates the minute in the hour of
    '    the day when the transition to standard time occurs
    '  - The ninth argument (intSecond) is an integer that indicates the second of the minute
    '    in the hour of the day when the transition to standard time occurs
    '  - The tenth argument (intMillisecond) is an integer that indicates the millisecond of
    '    the second of the minute in the hour of the day when the transition to standard time
    '    occurs
    '  - The eleventh argument (arrTimeZoneInstances) is an array/collection of objects of
    '    class Win32_TimeZone
    '
    ' NOTE: The third through tenth arguments will all be zero if a standard time start date
    ' is not applicable in the current time zone
    '
    ' The function returns a 0 if the standard time zone information was obtained successfully.
    ' It returns a negative integer if an error occurred retrieving the standard time zone
    ' information. Finally, it returns a positive integer if the standard time zone information
    ' was obtained, but multiple time zone instances were present that contained data for the
    ' standard time zone. When this happens, only the first Win32_TimeZone instance containing
    ' data for the standard time zone is used.
    '
    ' Example:
    '   intReturnCode = GetTimeZoneInstances(arrTimeZoneInstances)
    '   If intReturnCode >= 0 Then
    '       ' At least one Win32_TimeZone instance was retrieved successfully
    '       intReturnCode = GetStandardTimeZoneInfoUsingTimeZoneInstances(strStandardTimeZoneName, intStandardTimeOffsetFromUTC, intStartYear, intStartMonth, intNthDayOfWeekInMonth, intDayOfWeek, intHour, intMinute, intSecond, intMillisecond, arrTimeZoneInstances)
    '       If intReturnCode = 0 Then
    '           ' The computer's standard time zone information was retrieved successfully and
    '           ' is stored in strStandardTimeZoneName, intStandardTimeOffsetFromUTC,
    '           ' intStartYear, intStartMonth, intNthDayOfWeekInMonth, intDayOfWeek, intHour,
    '           ' intMinute, intSecond, and intMillisecond
    '       ElseIf intReturnCode > 0 Then
    '           ' More than one Win32_TimeZone instance containing data was present, which is
    '           ' unexpected. Still, the first instance was processed successfully and standard
    '           ' time zone information is stored in strStandardTimeZoneName,
    '           ' intStandardTimeOffsetFromUTC, intStartYear, intStartMonth,
    '           ' intNthDayOfWeekInMonth, intDayOfWeek, intHour, intMinute, intSecond, and
    '           ' intMillisecond
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
    ' None!
    'endregion Acknowledgements ####################################################

    'region DependsOn ####################################################
    ' TestObjectForData()
    ' TestObjectIsAnyTypeOfInteger()
    ' TestObjectIsStringContainingData()
    'endregion DependsOn ####################################################

    Dim intFunctionReturn
    Dim intReturnMultiplier
    Dim intTemp
    Dim intCounterA

    Dim strInterimTimeZoneName
    Dim intInterimStandardTimeInitialOffsetFromUTC
    Dim intInterimTimeBias
    Dim intInterimStandardTimeOffsetFromUTC
    Dim intInterimStartYear
    Dim intInterimStartMonth
    Dim intInterimNthDayOfWeekInMonth
    Dim intInterimDayOfWeek
    Dim intInterimHour
    Dim intInterimMinute
    Dim intInterimSecond
    Dim intInterimMillisecond

    Dim strOldInterimTimeZoneName
    Dim intOldInterimStandardTimeInitialOffsetFromUTC
    Dim intOldInterimTimeBias
    Dim intOldInterimStandardTimeOffsetFromUTC
    Dim intOldInterimStartYear
    Dim intOldInterimStartMonth
    Dim intOldInterimNthDayOfWeekInMonth
    Dim intOldInterimDayOfWeek
    Dim intOldInterimHour
    Dim intOldInterimMinute
    Dim intOldInterimSecond
    Dim intOldInterimMillisecond

    Dim strTimeZoneNameToReturn
    Dim intStandardTimeOffsetFromUTCToReturn
    Dim intStartYearToReturn
    Dim intStartMonthToReturn
    Dim intNthDayOfWeekInMonthToReturn
    Dim intDayOfWeekToReturn
    Dim intHourToReturn
    Dim intMinuteToReturn
    Dim intSecondToReturn
    Dim intMillisecondToReturn

    Dim intCountOfTimeZones
    Dim boolThisTimeZoneObjectWasValid

    Err.Clear

    intFunctionReturn = 0
    intReturnMultiplier = 128

    strInterimTimeZoneName = ""
    intInterimStandardTimeInitialOffsetFromUTC = 0
    intInterimTimeBias = 0
    intInterimStandardTimeOffsetFromUTC = 0
    intInterimStartYear = 0
    intInterimStartMonth = 0
    intInterimNthDayOfWeekInMonth = 0
    intInterimDayOfWeek = 0
    intInterimHour = 0
    intInterimMinute = 0
    intInterimSecond = 0
    intInterimMillisecond = 0

    strTimeZoneNameToReturn = ""
    intStandardTimeOffsetFromUTCToReturn = Null
    intStartYearToReturn = Null
    intStartMonthToReturn = Null
    intNthDayOfWeekInMonthToReturn = Null
    intDayOfWeekToReturn = Null
    intHourToReturn = Null
    intMinuteToReturn = Null
    intSecondToReturn = Null
    intMillisecondToReturn = Null

    intCountOfTimeZones = 0

    If TestObjectForData(arrTimeZoneInstances) <> True Then
        intFunctionReturn = intFunctionReturn + (-1 * intReturnMultiplier)
    Else
        On Error Resume Next
        intTemp = arrTimeZoneInstances.Count
        If Err Then
            On Error Goto 0
            Err.Clear
            intFunctionReturn = intFunctionReturn + (-2 * intReturnMultiplier)
        Else
            On Error Goto 0
            If TestObjectIsAnyTypeOfInteger(intTemp) = False Then
                intFunctionReturn = intFunctionReturn + (-3 * intReturnMultiplier)
            Else
                If intTemp < 0 Then
                    intFunctionReturn = intFunctionReturn + (-4 * intReturnMultiplier)
                ElseIf intTemp = 0 Then
                    intFunctionReturn = intFunctionReturn + (-5 * intReturnMultiplier)
                Else
                    For intCounterA = 0 To (intTemp - 1)
                        strOldInterimTimeZoneName = strInterimTimeZoneName
                        intOldInterimStandardTimeInitialOffsetFromUTC = intInterimStandardTimeInitialOffsetFromUTC
                        intOldInterimTimeBias = intInterimTimeBias
                        intOldInterimStandardTimeOffsetFromUTC = intInterimStandardTimeOffsetFromUTC
                        intOldInterimStartYear = intInterimStartYear
                        intOldInterimStartMonth = intInterimStartMonth
                        intOldInterimNthDayOfWeekInMonth = intInterimNthDayOfWeekInMonth
                        intOldInterimDayOfWeek = intInterimDayOfWeek
                        intOldInterimHour = intInterimHour
                        intOldInterimMinute = intInterimMinute
                        intOldInterimSecond = intInterimSecond
                        intOldInterimMillisecond = intInterimMillisecond

                        boolThisTimeZoneObjectWasValid = False

                        'region strStandardTimeZoneName ####################################################
                        On Error Resume Next
                        strInterimTimeZoneName = arrTimeZoneInstances.ItemIndex(intCounterA).StandardName
                        If Err Then
                            On Error Goto 0
                            Err.Clear
                            strInterimTimeZoneName = strOldInterimTimeZoneName
                        Else
                            On Error Goto 0
                            If TestObjectIsStringContainingData(strInterimTimeZoneName) <> True Then
                                strInterimTimeZoneName = strOldInterimTimeZoneName
                            Else
                                ' Found a result with real time zone data
                                If TestObjectForData(strTimeZoneNameToReturn) = False Then
                                    strTimeZoneNameToReturn = strInterimTimeZoneName
                                End If
                                boolThisTimeZoneObjectWasValid = True
                            End If
                        End If
                        'endregion strStandardTimeZoneName ####################################################

                        'region intStandardTimeOffsetFromUTC ####################################################
                        On Error Resume Next
                        intInterimStandardTimeInitialOffsetFromUTC = arrTimeZoneInstances.ItemIndex(intCounterA).Bias
                        If Err Then
                            On Error Goto 0
                            Err.Clear
                            intInterimStandardTimeInitialOffsetFromUTC = intOldInterimStandardTimeInitialOffsetFromUTC
                        Else
                            On Error Goto 0
                            If TestObjectIsAnyTypeOfInteger(intInterimStandardTimeInitialOffsetFromUTC) <> True Then
                                intInterimStandardTimeInitialOffsetFromUTC = intOldInterimStandardTimeInitialOffsetFromUTC
                            Else
                                On Error Resume Next
                                intInterimTimeBias = arrTimeZoneInstances.ItemIndex(intCounterA).StandardBias
                                If Err Then
                                    On Error Goto 0
                                    Err.Clear
                                    intInterimTimeBias = intOldInterimTimeBias
                                Else
                                    On Error Goto 0
                                    If TestObjectIsAnyTypeOfInteger(intInterimTimeBias) <> True Then
                                        intInterimTimeBias = intOldInterimTimeBias
                                    Else                                
                                        ' Found a result with real time zone data
                                        If TestObjectForData(intStandardTimeOffsetFromUTCToReturn) = False Then
                                            intStandardTimeOffsetFromUTCToReturn = intInterimStandardTimeInitialOffsetFromUTC - intInterimTimeBias
                                        End If
                                        boolThisTimeZoneObjectWasValid = True
                                    End If
                                End If
                            End If
                        End If
                        'endregion intStandardTimeOffsetFromUTC ####################################################

                        'region intStartYear ####################################################
                        On Error Resume Next
                        intInterimStartYear = arrTimeZoneInstances.ItemIndex(intCounterA).StandardYear
                        If Err Then
                            On Error Goto 0
                            Err.Clear
                            intInterimStartYear = intOldInterimStartYear
                        Else
                            On Error Goto 0
                            If TestObjectIsAnyTypeOfInteger(intInterimStartYear) <> True Then
                                intInterimStartYear = intOldInterimStartYear
                            Else
                                ' Found a result with real time zone data
                                If TestObjectForData(intStartYearToReturn) = False Then
                                    intStartYearToReturn = intInterimStartYear
                                End If
                                boolThisTimeZoneObjectWasValid = True
                            End If
                        End If
                        'endregion intStartYear ####################################################

                        'region intStartMonth ####################################################
                        On Error Resume Next
                        intInterimStartMonth = arrTimeZoneInstances.ItemIndex(intCounterA).StandardMonth
                        If Err Then
                            On Error Goto 0
                            Err.Clear
                            intInterimStartMonth = intOldInterimStartMonth
                        Else
                            On Error Goto 0
                            If TestObjectIsAnyTypeOfInteger(intInterimStartMonth) <> True Then
                                intInterimStartMonth = intOldInterimStartMonth
                            Else
                                ' Found a result with real time zone data
                                If TestObjectForData(intStartMonthToReturn) = False Then
                                    intStartMonthToReturn = intInterimStartMonth
                                End If
                                boolThisTimeZoneObjectWasValid = True
                            End If
                        End If
                        'endregion intStartMonth ####################################################

                        'region intNthDayOfWeekInMonth ####################################################
                        On Error Resume Next
                        intInterimNthDayOfWeekInMonth = arrTimeZoneInstances.ItemIndex(intCounterA).StandardDay
                        If Err Then
                            On Error Goto 0
                            Err.Clear
                            intInterimNthDayOfWeekInMonth = intOldInterimNthDayOfWeekInMonth
                        Else
                            On Error Goto 0
                            If TestObjectIsAnyTypeOfInteger(intInterimNthDayOfWeekInMonth) <> True Then
                                intInterimNthDayOfWeekInMonth = intOldInterimNthDayOfWeekInMonth
                            Else
                                ' Found a result with real time zone data
                                If TestObjectForData(intNthDayOfWeekInMonthToReturn) = False Then
                                    intNthDayOfWeekInMonthToReturn = intInterimNthDayOfWeekInMonth
                                End If
                                boolThisTimeZoneObjectWasValid = True
                            End If
                        End If
                        'endregion intNthDayOfWeekInMonth ####################################################

                        'region intDayOfWeek ####################################################
                        On Error Resume Next
                        intInterimDayOfWeek = arrTimeZoneInstances.ItemIndex(intCounterA).StandardDayOfWeek
                        If Err Then
                            On Error Goto 0
                            Err.Clear
                            intInterimDayOfWeek = intOldInterimDayOfWeek
                        Else
                            On Error Goto 0
                            If TestObjectIsAnyTypeOfInteger(intInterimDayOfWeek) <> True Then
                                intInterimDayOfWeek = intOldInterimDayOfWeek
                            Else
                                ' Found a result with real time zone data
                                If TestObjectForData(intDayOfWeekToReturn) = False Then
                                    intDayOfWeekToReturn = intInterimDayOfWeek
                                End If
                                boolThisTimeZoneObjectWasValid = True
                            End If
                        End If
                        'endregion intDayOfWeek ####################################################

                        'region intHour ####################################################
                        On Error Resume Next
                        intInterimHour = arrTimeZoneInstances.ItemIndex(intCounterA).StandardHour
                        If Err Then
                            On Error Goto 0
                            Err.Clear
                            intInterimHour = intOldInterimHour
                        Else
                            On Error Goto 0
                            If TestObjectIsAnyTypeOfInteger(intInterimHour) <> True Then
                                intInterimHour = intOldInterimHour
                            Else
                                ' Found a result with real time zone data
                                If TestObjectForData(intHourToReturn) = False Then
                                    intHourToReturn = intInterimHour
                                End If
                                boolThisTimeZoneObjectWasValid = True
                            End If
                        End If
                        'endregion intHour ####################################################

                        'region intMinute ####################################################
                        On Error Resume Next
                        intInterimMinute = arrTimeZoneInstances.ItemIndex(intCounterA).StandardMinute
                        If Err Then
                            On Error Goto 0
                            Err.Clear
                            intInterimMinute = intOldInterimMinute
                        Else
                            On Error Goto 0
                            If TestObjectIsAnyTypeOfInteger(intInterimMinute) <> True Then
                                intInterimMinute = intOldInterimMinute
                            Else
                                ' Found a result with real time zone data
                                If TestObjectForData(intMinuteToReturn) = False Then
                                    intMinuteToReturn = intInterimMinute
                                End If
                                boolThisTimeZoneObjectWasValid = True
                            End If
                        End If
                        'endregion intMinute ####################################################

                        'region intSecond ####################################################
                        On Error Resume Next
                        intInterimSecond = arrTimeZoneInstances.ItemIndex(intCounterA).StandardSecond
                        If Err Then
                            On Error Goto 0
                            Err.Clear
                            intInterimSecond = intOldInterimSecond
                        Else
                            On Error Goto 0
                            If TestObjectIsAnyTypeOfInteger(intInterimSecond) <> True Then
                                intInterimSecond = intOldInterimSecond
                            Else
                                ' Found a result with real time zone data
                                If TestObjectForData(intSecondToReturn) = False Then
                                    intSecondToReturn = intInterimSecond
                                End If
                                boolThisTimeZoneObjectWasValid = True
                            End If
                        End If
                        'endregion intSecond ####################################################

                        'region intMillisecond ####################################################
                        On Error Resume Next
                        intInterimMillisecond = arrTimeZoneInstances.ItemIndex(intCounterA).StandardMillisecond
                        If Err Then
                            On Error Goto 0
                            Err.Clear
                            intInterimMillisecond = intOldInterimMillisecond
                        Else
                            On Error Goto 0
                            If TestObjectIsAnyTypeOfInteger(intInterimMillisecond) <> True Then
                                intInterimMillisecond = intOldInterimMillisecond
                            Else
                                ' Found a result with real time zone data
                                If TestObjectForData(intMillisecondToReturn) = False Then
                                    intMillisecondToReturn = intInterimMillisecond
                                End If
                                boolThisTimeZoneObjectWasValid = True
                            End If
                        End If
                        'endregion intMillisecond ####################################################

                        If boolThisTimeZoneObjectWasValid = True Then
                            intCountOfTimeZones = intCountOfTimeZones + 1
                        End If
                    Next
                End If
            End If
        End If
    End If

    If intFunctionReturn >= 0 Then
        ' No error has occurred yet
        If intCountOfTimeZones = 0 Then
            ' No result found
            intFunctionReturn = intFunctionReturn + (-5 * intReturnMultiplier)
        Else
            intFunctionReturn = intCountOfTimeZones - 1
        End If
    End If

    If intFunctionReturn >= 0 Then
        strStandardTimeZoneName = strTimeZoneNameToReturn
        intStandardTimeOffsetFromUTC = intStandardTimeOffsetFromUTCToReturn
        intStartYear = intStartYearToReturn
        intStartMonth = intStartMonthToReturn
        intNthDayOfWeekInMonth = intNthDayOfWeekInMonthToReturn
        intDayOfWeek = intDayOfWeekToReturn
        intHour = intHourToReturn
        intMinute = intMinuteToReturn
        intSecond = intSecondToReturn
        intMillisecond = intMillisecondToReturn
    End If
    
    GetStandardTimeZoneInfoUsingTimeZoneInstances = intFunctionReturn
End Function
