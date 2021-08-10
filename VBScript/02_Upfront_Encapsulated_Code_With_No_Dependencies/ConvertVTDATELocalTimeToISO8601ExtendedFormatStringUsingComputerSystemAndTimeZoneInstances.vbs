Function ConvertVTDATELocalTimeToISO8601ExtendedFormatStringUsingComputerSystemAndTimeZoneInstances(ByRef strISO8601Output, ByVal datetimeInput, ByVal arrComputerSystemInstances, ByVal arrTimeZoneInstances)
    'region FunctionMetadata ####################################################
    ' Assuming that arrComputerSystemInstances represents an array / collection of the
    ' available computer system instances (of type Win32_ComputerSystem), and
    ' arrTimeZoneInstances represents an array / collection of the available time zone
    ' instances (of type Win32_TimeZone), safely takes a VT_DATETIME (VBScript-native datetime
    ' variant) object and converts it to a string representation of the date and time, in
    ' compliance with ISO 8601's extended format
    '
    ' The function takes four positional arguments:
    '   The first argument (strISO8601Output) is set upon success to a string representation of
    '       the date and time specified by the second argument (datetimeInput), represented in
    '       compliance with ISO 8601
    '   The second argument (datetimeInput) is a VBScript-native datetime object (VT_DATE)
    '       containing a date and time in the local computer's time zone
    '   The third argument (arrComputerSystemInstances) is an array/collection of objects of
    '       class Win32_ComputerSystem
    '   The fourth argument (arrTimeZoneInstances) is an array/collection of objects of class
    '       Win32_TimeZone
    '
    ' The function returns 0 or a positive number if the VBScript-native datetime object
    ' (VT_DATETIME) was converted to a ISO 8601-formatted string. A return of 1 indicates a
    ' warning condition in which the local computer's time zone adjustment could not be
    ' determined. It returns a negative number if an error occurred
    '
    ' Example:
    '   datetimeDateFunctionAuthored = DateSerial(2021, 8, 8)
    '   datetimeDateFunctionAuthored = datetimeDateFunctionAuthored + TimeSerial(13, 0, 0)
    '   intReturnCode = ConnectLocalWMINamespace(objSWbemServicesWMINamespace, Null, Null)
    '   intReturnCode = GetComputerSystemInstancesUsingWMINamespace(arrComputerSystemInstances, objSWbemServicesWMINamespace)
    '   intReturnCode = GetTimeZoneInstancesUsingWMINamespace(arrTimeZoneInstances, objSWbemServicesWMINamespace)
    '   intReturnCode = ConvertVTDATELocalTimeToISO8601ExtendedFormatStringUsingComputerSystemAndTimeZoneInstances(strISO8601Output, datetimeDateFunctionAuthored, arrComputerSystemInstances, arrTimeZoneInstances)
    '   If intReturnCode >= 0 Then
    '       ' Conversion completed successfully
    '       ' On a computer in the Central (US) Daylight Time (GMT-5) time zone,
    '       ' strISO8601Output is "2021-08-05T13:00:00-05:00"
    '   End If
    '
    ' Version: 1.0.20210810.0
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

    'region DependsOn ####################################################
    ' TestObjectIsDateTimeContainingData()
    ' GetCurrentEffectiveTimeZoneUTCOffsetInMinutesUsingComputerSystemAndTimeZoneInstances()
    ' GetUTCOffsetForDateInLocalTimeZoneUsingTimeZoneInstances()
    ' TestObjectIsAnyTypeOfNumber()
    'endregion DependsOn ####################################################

    Const NUM_MINUTES_IN_HOUR = 60

    Dim intFunctionReturn
    Dim intReturnMultiplier
    Dim intReturnCode
    Dim strISO8601OutputToReturn

    Dim intTemp
    Dim strYear
    Dim strMonth
    Dim strDay
    Dim strHour
    Dim strMinute
    Dim strSecond
    Dim strTotalUTCOffset
    Dim strUTCOffsetSign
    Dim objRemainingMinutes
    Dim intWorkingUTCOffsetHours
    Dim strUTCOffsetHours
    Dim intWorkingUTCOffsetMinutes
    Dim strUTCOffsetMinutes
    Dim boolErrorOccurred
    Dim intCurrentUTCOffset
    Dim intLocalComputerTimeZoneUTCOffsetAtSpecifiedDate

    intFunctionReturn = 0
    intReturnMultiplier = 128 * 8 * 2 * 8

    If TestObjectIsDateTimeContainingData(datetimeInput) <> True Then
        intFunctionReturn = intFunctionReturn + (-1 * intReturnMultiplier)
    Else
        boolErrorOccurred = False

        ' Convert the year
        If boolErrorOccurred = False Then
            On Error Resume Next
            intTemp = Year(datetimeInput)
            If Err Then
                On Error Goto 0
                Err.Clear
                boolErrorOccurred = True
            Else
                strYear = CStr(intTemp)
                If Err Then
                    On Error Goto 0
                    Err.Clear
                    boolErrorOccurred = True
                Else
                    On Error Goto 0
                    If Len(strYear) < 1 Then
                        ' Len() was 0; error condition
                        boolErrorOccurred = True
                    Else
                        ' Len() is >= 1
                        If Left(strYear, 1) = "-" Then
                            ' BCE year
                            If Len(strYear) < 2 Then
                                boolErrorOccurred = True
                            Else
                                If Len(Right(strYear, Len(strYear) - 1)) < 4 Then
                                    strYear = Left(strYear, 1) & Right("000" & Right(strYear, Len(strYear) - 1), 4)
                                End If
                            End If
                        Else
                            ' CE year or 1 BCE
                            If Len(strYear) < 4 Then
                                strYear = Right("000" & strYear, 4)
                            End If
                        End If
                    End If
                End If
            End If
        End If

        ' Convert the month
        If boolErrorOccurred = False Then
            On Error Resume Next
            intTemp = Month(datetimeInput)
            If Err Then
                On Error Goto 0
                Err.Clear
                boolErrorOccurred = True
            Else
                On Error Goto 0
                If intTemp < 1 Or intTemp > 12 Then
                    boolErrorOccurred = True
                Else
                    On Error Resume Next
                    strMonth = CStr(intTemp)
                    If Err Then
                        On Error Goto 0
                        Err.Clear
                        boolErrorOccurred = True
                    Else
                        On Error Goto 0
                        strMonth = Right("0" & strMonth, 2)
                    End If
                End If
            End If
        End If

        ' Convert the day
        If boolErrorOccurred = False Then
            On Error Resume Next
            intTemp = Day(datetimeInput)
            If Err Then
                On Error Goto 0
                Err.Clear
                boolErrorOccurred = True
            Else
                On Error Goto 0
                If intTemp < 1 Or intTemp > 31 Then
                    boolErrorOccurred = True
                Else
                    On Error Resume Next
                    strDay = CStr(intTemp)
                    If Err Then
                        On Error Goto 0
                        Err.Clear
                        boolErrorOccurred = True
                    Else
                        On Error Goto 0
                        strDay = Right("0" & strDay, 2)
                    End If
                End If
            End If
        End If

        ' Convert the hour
        If boolErrorOccurred = False Then
            On Error Resume Next
            intTemp = Hour(datetimeInput)
            If Err Then
                On Error Goto 0
                Err.Clear
                boolErrorOccurred = True
            Else
                On Error Goto 0
                If intTemp < 0 Or intTemp > 23 Then
                    boolErrorOccurred = True
                Else
                    On Error Resume Next
                    strHour = CStr(intTemp)
                    If Err Then
                        On Error Goto 0
                        Err.Clear
                        boolErrorOccurred = True
                    Else
                        On Error Goto 0
                        strHour = Right("0" & strHour, 2)
                    End If
                End If
            End If
        End If

        ' Convert the minute
        If boolErrorOccurred = False Then
            On Error Resume Next
            intTemp = Minute(datetimeInput)
            If Err Then
                On Error Goto 0
                Err.Clear
                boolErrorOccurred = True
            Else
                On Error Goto 0
                If intTemp < 0 Or intTemp > 59 Then
                    boolErrorOccurred = True
                Else
                    On Error Resume Next
                    strMinute = CStr(intTemp)
                    If Err Then
                        On Error Goto 0
                        Err.Clear
                        boolErrorOccurred = True
                    Else
                        On Error Goto 0
                        strMinute = Right("0" & strMinute, 2)
                    End If
                End If
            End If
        End If

        ' Convert the second
        If boolErrorOccurred = False Then
            On Error Resume Next
            intTemp = Second(datetimeInput)
            If Err Then
                On Error Goto 0
                Err.Clear
                boolErrorOccurred = True
            Else
                On Error Goto 0
                If intTemp < 0 Or intTemp > 59 Then
                    boolErrorOccurred = True
                Else
                    On Error Resume Next
                    strSecond = CStr(intTemp)
                    If Err Then
                        On Error Goto 0
                        Err.Clear
                        boolErrorOccurred = True
                    Else
                        On Error Goto 0
                        strSecond = Right("0" & strSecond, 2)
                    End If
                End If
            End If
        End If

        ' Determine time zone offset
        If boolErrorOccurred = False Then
            ' Populate intLocalComputerTimeZoneUTCOffsetAtSpecifiedDate
            ' We have arrTimeZoneInstances, which means we can try
            ' GetUTCOffsetForDateInLocalTimeZoneUsingTimeZoneInstances()
            intReturnCode = GetUTCOffsetForDateInLocalTimeZoneUsingTimeZoneInstances(intTemp, datetimeInput, arrTimeZoneInstances)
            If intReturnCode = 0 Then
                intLocalComputerTimeZoneUTCOffsetAtSpecifiedDate = intTemp
            End If

            ' intLocalComputerTimeZoneUTCOffsetAtSpecifiedDate is either Null or populated
            ' correctly.
            If TestObjectIsAnyTypeOfNumber(intLocalComputerTimeZoneUTCOffsetAtSpecifiedDate) = True Then
                ' Use intLocalComputerTimeZoneUTCOffsetAtSpecifiedDate
                If intLocalComputerTimeZoneUTCOffsetAtSpecifiedDate = 0 Then
                    ' Coordinated Universal Time (UTC)
                    strTotalUTCOffset = "Z"
                Else
                    If intLocalComputerTimeZoneUTCOffsetAtSpecifiedDate < 0 Then
                        strUTCOffsetSign = "-"
                        ' Make it positive
                        intLocalComputerTimeZoneUTCOffsetAtSpecifiedDate = intLocalComputerTimeZoneUTCOffsetAtSpecifiedDate * -1
                    Else
                        strUTCOffsetSign = "+"
                    End If
                    intWorkingUTCOffsetHours = Int(intLocalComputerTimeZoneUTCOffsetAtSpecifiedDate / NUM_MINUTES_IN_HOUR)
                    objRemainingMinutes = intLocalComputerTimeZoneUTCOffsetAtSpecifiedDate - (intWorkingUTCOffsetHours * NUM_MINUTES_IN_HOUR)
                    intWorkingUTCOffsetMinutes = Int(objRemainingMinutes)
                    objRemainingMinutes = objRemainingMinutes - intWorkingUTCOffsetMinutes
                    strUTCOffsetHours = Right("0" & CStr(intWorkingUTCOffsetHours), 2)
                    strUTCOffsetMinutes = Right("0" & CStr(intWorkingUTCOffsetMinutes), 2)
                    strTotalUTCOffset = strUTCOffsetSign & strUTCOffsetHours & ":" & strUTCOffsetMinutes
                End If
            Else
                ' Populate intCurrentUTCOffset
                ' We should have at least one of arrComputerSystemInstances or
                ' arrTimeZoneInstances, which means we can try
                ' GetCurrentEffectiveTimeZoneUTCOffsetInMinutesUsingComputerSystemAndTimeZoneInstances()
                intReturnCode = GetCurrentEffectiveTimeZoneUTCOffsetInMinutesUsingComputerSystemAndTimeZoneInstances(intTemp, arrComputerSystemInstances, arrTimeZoneInstances)
                If intReturnCode >= 0 Then
                    intCurrentUTCOffset = intTemp
                End If
                If TestObjectIsAnyTypeOfNumber(intCurrentUTCOffset) = True Then
                    ' Use intCurrentUTCOffset
                    If intCurrentUTCOffset = 0 Then
                        ' Coordinated Universal Time (UTC)
                        strTotalUTCOffset = "Z"
                    Else
                        If intCurrentUTCOffset < 0 Then
                            strUTCOffsetSign = "-"
                            ' Make it positive
                            intCurrentUTCOffset = intCurrentUTCOffset * -1
                        Else
                            strUTCOffsetSign = "+"
                        End If
                        intWorkingUTCOffsetHours = Int(intCurrentUTCOffset / NUM_MINUTES_IN_HOUR)
                        objRemainingMinutes = intCurrentUTCOffset - (intWorkingUTCOffsetHours * NUM_MINUTES_IN_HOUR)
                        intWorkingUTCOffsetMinutes = Int(objRemainingMinutes)
                        objRemainingMinutes = objRemainingMinutes - intWorkingUTCOffsetMinutes
                        strUTCOffsetHours = Right("0" & CStr(intWorkingUTCOffsetHours), 2)
                        strUTCOffsetMinutes = Right("0" & CStr(intWorkingUTCOffsetMinutes), 2)
                        strTotalUTCOffset = strUTCOffsetSign & strUTCOffsetHours & ":" & strUTCOffsetMinutes
                    End If
                Else
                    ' We do not know the local computer's time zone adjustment from UTC
                    ' Don't make any adjustments
                    strTotalUTCOffset = ""
                    intFunctionReturn = 1 ' Warning condition
                End If
            End If
        End If

        ' Assemble output
        If boolErrorOccurred = False Then
            strISO8601OutputToReturn = strYear & "-" & strMonth & "-" & strDay & "T" & strHour & ":" & strMinute & ":" & strSecond & strTotalUTCOffset
        End If
    End If

    If intFunctionReturn >= 0 Then
        strISO8601Output = strISO8601OutputToReturn
    End If

    ConvertVTDATELocalTimeToISO8601ExtendedFormatStringUsingComputerSystemAndTimeZoneInstances = intFunctionReturn
End Function
