Function ConvertCIMDATETIMEToISO8601ExtendedFormatString(ByRef strISO8601Output, ByVal strCIMDATETIMEInput)
    'region FunctionMetadata ####################################################
    ' Safely takes a string in the CIM_DATETIME format (see below) and converts it to a string
    ' representation of the date and time, in compliance with ISO 8601's extended format
    '
    ' A CIM_DATETIME object is a string in the following format:
    ' yyyymmddHHMMSS.mmmmmmsUUU
    ' yyyy = Four-digit year (0000 through 9999)
    ' mm = Two-digit month (01 through 12)
    ' dd = Two-digit day of the month (01 through 31). This value must be appropriate for the
    '      month. For example, February 31 is invalid
    ' HH = Two-digit hour of the day using the 24-hour clock (00 through 23)
    ' MM = Two-digit minute in the hour (00 through 59)
    ' SS = Two-digit number of seconds in the minute (00 through 59)
    ' mmmmmm = Six-digit number of microseconds in the second (000000 through 999999). This
    '          field must always be present to preserve the fixed-length nature of the string
    ' s = Plus sign (+) or minus sign (-) to indicate a positive or negative offset from
    '     Universal Time Coordinates (UTC)
    ' UUU = Three-digit offset indicating the number of minutes that the originating time zone
    '       deviates from UTC
    '
    ' The function takes two positional arguments:
    '   The first argument (strISO8601Output) is set upon success to a string representation of
    '       the date and time specified by the second argument (strCIMDATETIMEInput), represented in
    '       compliance with ISO 8601
    '   The second argument (strCIMDATETIMEInput) contains the date time object to be converted
    '       in string CIM_DATETIME format (see above)
    '
    ' The function returns 0 if the string CIM_DATETIME-formatted object was converted
    ' successfully to a ISO 2601-formatted string. It returns a negative number if an error
    ' occurred
    '
    ' Example:
    '   strCIMDATETIMEInput = "20210421000000.000000-360"
    '   intReturnCode = ConvertCIMDATETIMEToISO8601ExtendedFormatString(strISO8601Output, strCIMDATETIMEInput)
    '   If intReturnCode >= 0 Then
    '       ' Conversion completed successfully
    '       ' strISO8601Output is "2021-04-21T00:00:00-06:00"
    '   End If
    '
    ' Version: 1.0.20210809.0
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
    ' TestObjectIsStringContainingData()
    ' TestObjectIsAnyTypeOfNumber()
    'endregion DependsOn ####################################################

    Const NUM_MINUTES_IN_HOUR = 60

    Dim intFunctionReturn
    Dim intReturnMultiplier
    Dim intReturnCode
    Dim strISO8601OutputToReturn

    'Dim strTemp
    Dim intTemp
    Dim intResult
    Dim boolResult
    Dim boolResultB
    Dim boolResultC
    Dim arrTemp
    Dim arrTempA
    Dim arrTempB
    Dim intEffectiveLength
    Dim strYear
    Dim strMonth
    Dim strDay
    Dim strHour
    Dim strMinute
    Dim strSecond
    Dim strMicrosecond
    Dim strTotalUTCOffset
    Dim strUTCOffsetSign
    Dim objRemainingMinutes
    Dim intWorkingUTCOffsetHours
    Dim strUTCOffsetHours
    Dim intWorkingUTCOffsetMinutes
    Dim strUTCOffsetMinutes
    Dim boolErrorOccurred
    Dim boolMinorErrorOccurred
    Dim intSpecifiedUTCOffset
    'Dim intLocalComputerTimeZoneUTCOffsetAtSpecifiedDate

    intFunctionReturn = 0
    intReturnMultiplier = 1

    If TestObjectIsStringContainingData(strCIMDATETIMEInput) <> True Then
        intFunctionReturn = intFunctionReturn + (-1 * intReturnMultiplier)
    Else
        boolErrorOccurred = False

        If boolErrorOccurred = False Then
            If Len(strCIMDATETIMEInput) < 8 Then
                boolErrorOccurred = True
            End If
        End If

        ' Convert the year
        ' Note: CIM_DATETIME only allows four digit years between 0000 and 9999.
        If boolErrorOccurred = False Then
            On Error Resume Next
            strYear = Left(strCIMDATETIMEInput, 4)
            If Err Then
                On Error Goto 0
                Err.Clear
                boolErrorOccurred = True
            Else
                intTemp = CInt(strYear)
                If Err Then
                    On Error Goto 0
                    Err.Clear
                    boolErrorOccurred = True
                Else
                    On Error Goto 0
                    If intTemp < 0 Or intTemp > 9999 Then
                        boolErrorOccurred = True
                    End if
                End If
            End If
        End If

        ' Convert the month
        If boolErrorOccurred = False Then
            On Error Resume Next
            strMonth = Mid(strCIMDATETIMEInput, 5, 2)
            If Err Then
                On Error Goto 0
                Err.Clear
                boolErrorOccurred = True
            Else
                intTemp = CInt(strMonth)
                If Err Then
                    On Error Goto 0
                    Err.Clear
                    boolErrorOccurred = True
                Else
                    On Error Goto 0
                    If intTemp < 1 Or intTemp > 12 Then
                        boolErrorOccurred = True
                    End If
                End If
            End If
        End If

        ' Convert the day
        If boolErrorOccurred = False Then
            On Error Resume Next
            strDay = Mid(strCIMDATETIMEInput, 7, 2)
            If Err Then
                On Error Goto 0
                Err.Clear
                boolErrorOccurred = True
            Else
                intTemp = CInt(strDay)
                If Err Then
                    On Error Goto 0
                    Err.Clear
                    boolErrorOccurred = True
                Else
                    On Error Goto 0
                    If intTemp < 1 Or intTemp > 31 Then
                        boolErrorOccurred = True
                    End If
                End If
            End If
        End If

        If boolErrorOccurred = False Then
            boolMinorErrorOccurred = False
        End If

        ' Convert the hour
        If boolErrorOccurred = False And boolMinorErrorOccurred = False Then
            On Error Resume Next
            strHour = Mid(strCIMDATETIMEInput, 9, 2)
            If Err Then
                On Error Goto 0
                Err.Clear
                boolMinorErrorOccurred = True
            Else
                intTemp = CInt(strHour)
                If Err Then
                    On Error Goto 0
                    Err.Clear
                    boolMinorErrorOccurred = True
                Else
                    On Error Goto 0
                    If intTemp < 0 Or intTemp > 23 Then
                        boolMinorErrorOccurred = True
                    End If
                End If
            End If
        End If

        ' Convert the minute
        If boolErrorOccurred = False And boolMinorErrorOccurred = False Then
            On Error Resume Next
            strMinute = Mid(strCIMDATETIMEInput, 11, 2)
            If Err Then
                On Error Goto 0
                Err.Clear
                boolMinorErrorOccurred = True
            Else
                intTemp = CInt(strMinute)
                If Err Then
                    On Error Goto 0
                    Err.Clear
                    boolMinorErrorOccurred = True
                Else
                    On Error Goto 0
                    If intTemp < 0 Or intTemp > 59 Then
                        boolMinorErrorOccurred = True
                    End If
                End If
            End If
        End If

        ' Convert the second
        If boolErrorOccurred = False And boolMinorErrorOccurred = False Then
            On Error Resume Next
            strSecond = Mid(strCIMDATETIMEInput, 13, 2)
            If Err Then
                On Error Goto 0
                Err.Clear
                boolMinorErrorOccurred = True
            Else
                intTemp = CInt(strSecond)
                If Err Then
                    On Error Goto 0
                    Err.Clear
                    boolMinorErrorOccurred = True
                Else
                    On Error Goto 0
                    If intTemp < 0 Or intTemp > 59 Then
                        boolMinorErrorOccurred = True
                    End If
                End If
            End If
        End If

        ' Convert the microsecond
        If boolErrorOccurred = False And boolMinorErrorOccurred = False Then
            On Error Resume Next
            intResult = InStr(1, strCIMDATETIMEInput, ".", 1)
            If Err Then
                On Error Goto 0
                Err.Clear
                boolMinorErrorOccurred = True
            Else
                boolResult = (intResult > 0)
                If Err Then
                    On Error Goto 0
                    Err.Clear
                    boolMinorErrorOccurred = True
                Else
                    intResult = InStr(1, strCIMDATETIMEInput, "+", 1)
                    If Err Then
                        On Error Goto 0
                        Err.Clear
                        boolMinorErrorOccurred = True
                    Else
                        boolResultB = (intResult > 0)
                        If Err Then
                            On Error Goto 0
                            Err.Clear
                            boolMinorErrorOccurred = True
                        Else
                            intResult = InStr(1, strCIMDATETIMEInput, "-", 1)
                            If Err Then
                                On Error Goto 0
                                Err.Clear
                                boolMinorErrorOccurred = True
                            Else
                                boolResultC = (intResult > 0)
                                If Err Then
                                    On Error Goto 0
                                    Err.Clear
                                    boolMinorErrorOccurred = True
                                Else
                                    On Error Goto 0
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
        If boolErrorOccurred = False And boolMinorErrorOccurred = False Then
            ' boolResult = True if strCIMDATETIMEInput contained a ".", False otherwise
            ' boolResultB = True if strCIMDATETIMEInput contained a "+", False otherwise
            ' boolResultC = True if strCIMDATETIMEInput contained a "-", False otherwise
            If boolResult = False Then
                ' No microsecond specified
                strMicrosecond = "000000"
            Else
                arrTempA = Split(strCIMDATETIMEInput, ".")
                If UBound(arrTempA) < 1 Then
                    boolMinorErrorOccurred = True
                Else
                    If Len(arrTempA(1)) <= 0 Then
                        strMicrosecond = "000000"
                    Else
                        If boolResultB = False And boolResultC = False Then
                            ' No time zone bias provided
                            intEffectiveLength = Len(arrTempA(1))
                            If intEffectiveLength < 6 Then
                                intEffectiveLength = 6
                            End If
                            strMicrosecond = Right("00000" & arrTempA(1), intEffectiveLength)
                        Else
                            If boolResultB = True And boolResultC = True Then
                                boolMinorErrorOccurred = True
                            Else
                                If boolResultB = True Then
                                    arrTempB = Split(arrTempA(1), "+")
                                    If Len(arrTempB(0)) <= 0 Then
                                        boolMinorErrorOccurred = True
                                    Else
                                        intEffectiveLength = Len(arrTempB(0))
                                        If intEffectiveLength < 6 Then
                                            intEffectiveLength = 6
                                        End If
                                        strMicrosecond = Right("00000" & arrTempB(0), intEffectiveLength)
                                    End If
                                Else
                                    arrTempB = Split(arrTempA(1), "-")
                                    If Len(arrTempB(0)) <= 0 Then
                                        boolMinorErrorOccurred = True
                                    Else
                                        intEffectiveLength = Len(arrTempB(0))
                                        If intEffectiveLength < 6 Then
                                            intEffectiveLength = 6
                                        End If
                                        strMicrosecond = Right("00000" & arrTempB(0), intEffectiveLength)
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
        If boolErrorOccurred = False And boolMinorErrorOccurred = False Then
            'On Error Resume Next
            intTemp = CLng(strMicrosecond)
            If Err Then
                On Error Goto 0
                Err.Clear
                boolMinorErrorOccurred = True
                strMicrosecond = ""
            Else
                On Error Goto 0
                If intTemp < 0 Then
                    boolMinorErrorOccurred = True
                    strMicrosecond = ""
                ElseIf intTemp = 0 Then
                    strMicrosecond = ""
                Else
                    While Right(strMicrosecond, 1) = "0"
                        strMicrosecond = Left(strMicrosecond, Len(strMicrosecond) - 1)
                    WEnd
                    strMicrosecond = "," & strMicrosecond
                End If
            End If
        End If

        ' Determine time zone offset
        If boolErrorOccurred = False And boolMinorErrorOccurred = False Then
            ' Populate intSpecifiedUTCOffset
            On Error Resume Next
            arrTemp = Split(strCIMDATETIMEInput, "+")
            If Err Then
                On Error Goto 0
                Err.Clear
                boolMinorErrorOccurred = True
            Else
                intTemp = UBound(arrTemp)
                If Err Then
                    On Error Goto 0
                    Err.Clear
                    boolMinorErrorOccurred = True
                Else
                    On Error Goto 0
                    If intTemp = 0 Then
                        ' No plus sign present
                        On Error Resume Next
                        arrTemp = Split(strCIMDATETIMEInput, "-")
                        If Err Then
                            On Error Goto 0
                            Err.Clear
                            boolMinorErrorOccurred = True
                        Else
                            intTemp = UBound(arrTemp)
                            If Err Then
                                On Error Goto 0
                                Err.Clear
                                boolMinorErrorOccurred = True
                            Else
                                On Error Goto 0
                                If intTemp = 0 Then
                                    ' No bias found
                                    intSpecifiedUTCOffset = 0
                                Else
                                    ' Negative bias present
                                    On Error Resume Next
                                    intTemp = CInt(arrTemp(1))
                                    If Err Then
                                        On Error Goto 0
                                        Err.Clear
                                        boolMinorErrorOccurred = True
                                    Else
                                        On Error Goto 0
                                        intSpecifiedUTCOffset = intTemp * -1
                                    End If
                                End If
                            End If
                        End If
                    Else
                        ' Positive bias present
                        On Error Resume Next
                        intTemp = CInt(arrTemp(1))
                        If Err Then
                            On Error Goto 0
                            Err.Clear
                            boolMinorErrorOccurred = True
                        Else
                            On Error Goto 0
                            intSpecifiedUTCOffset = intTemp
                        End If
                    End If
                End If
            End If

            ' intSpecifiedUTCOffset is either Null or populated correctly.
            If TestObjectIsAnyTypeOfNumber(intSpecifiedUTCOffset) = True Then
                ' Use intSpecifiedUTCOffset
                If intSpecifiedUTCOffset = 0 Then
                    ' Coordinated Universal Time (UTC)
                    strTotalUTCOffset = "Z"
                Else
                    If intSpecifiedUTCOffset < 0 Then
                        strUTCOffsetSign = "-"
                        ' Make it positive
                        intSpecifiedUTCOffset = intSpecifiedUTCOffset * -1
                    Else
                        strUTCOffsetSign = "+"
                    End If
                    intWorkingUTCOffsetHours = Int(intSpecifiedUTCOffset / NUM_MINUTES_IN_HOUR)
                    objRemainingMinutes = intSpecifiedUTCOffset - (intWorkingUTCOffsetHours * NUM_MINUTES_IN_HOUR)
                    intWorkingUTCOffsetMinutes = Int(objRemainingMinutes)
                    objRemainingMinutes = objRemainingMinutes - intWorkingUTCOffsetMinutes
                    strUTCOffsetHours = Right("0" & CStr(intWorkingUTCOffsetHours), 2)
                    strUTCOffsetMinutes = Right("0" & CStr(intWorkingUTCOffsetMinutes), 2)
                    strTotalUTCOffset = strUTCOffsetSign & strUTCOffsetHours & ":" & strUTCOffsetMinutes
                End If
            Else
                ' We could not determine the CIM_DATETIME UTC offset
                ' Don't make any adjustments
                strTotalUTCOffset = ""
                intFunctionReturn = 1 ' Warning condition
            End If
        End If
    End If

    ' Assemble Output
    If boolErrorOccurred = False Then
        If boolMinorErrorOccurred = True Then
            ' Could not determine time or time zone offset
            strISO8601OutputToReturn = strYear & "-" & strMonth & "-" & strDay
            intFunctionReturn = intFunctionReturn + 2
        Else
            strISO8601OutputToReturn = strYear & "-" & strMonth & "-" & strDay & "T" & strHour & ":" & strMinute & ":" & strSecond & strMicrosecond & strTotalUTCOffset
        End If
    End If

    If intFunctionReturn >= 0 Then
        strISO8601Output = strISO8601OutputToReturn
    End If

    ConvertCIMDATETIMEToISO8601ExtendedFormatString = intFunctionReturn
End Function
