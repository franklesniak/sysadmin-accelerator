Function ConvertStringCIMDATETIMEToVTDATELocalTimeUsingComputerSystemAndTimeZoneInstances(ByRef datetimeOutput, ByVal strCIMDATETIMEInput, ByVal arrComputerSystemInstances, ByVal arrTimeZoneInstances)
    'region FunctionMetadata ####################################################
    ' Assuming that arrComputerSystemInstances represents an array / collection of the
    ' available computer system instances (of type Win32_ComputerSystem), and
    ' arrTimeZoneInstances represents an array / collection of the available time zone
    ' instances (of type Win32_TimeZone), safely takes a string in the CIM_DATETIME format (see
    ' below) and converts it to a VT_DATETIME (VBScript-native datetime variant) object, set to
    ' local time
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
    ' The function takes four positional arguments:
    '   The first argument (datetimeOutput) is set upon success to a VBScript-native datetime
    '       object (VT_DATE). The object will contain the conveted date and time from the
    '       second argument (strCIMDATETIMEInput), adjusted to the local computer's time zone
    '   The second argument (strCIMDATETIMEInput) contains the date time object to be converted
    '       in string CIM_DATETIME format (see above)
    '   The third argument (arrComputerSystemInstances) is an array/collection of objects of
    '       class Win32_ComputerSystem
    '   The fourth argument (arrTimeZoneInstances) is an array/collection of objects of class
    '       Win32_TimeZone
    '
    ' The function returns 0 or a positive number if the string CIM_DATETIME-formatted object
    ' was converted successfully to a VBScript-native datetime object (VT_DATETIME). It returns
    ' a negative number if an error occurred
    '
    ' Note: On operating systems prior to Windows XP and Windows Server 2003, subsecond time
    ' resolution is lost during the conversion process
    '
    ' Example:
    '   strCIMDATETIMEInput = "20210421000000.000000-360"
    '   intReturnCode = ConnectLocalWMINamespace(objSWbemServicesWMINamespace, Null, Null)
    '   If intReturnCode = 0 Then
    '       ' Successfully connected to the local computer's root\CIMv2 WMI Namespace
    '       intReturnCode = GetComputerSystemInstancesUsingWMINamespace(arrComputerSystemInstances, objSWbemServicesWMINamespace)
    '       If intReturnCode >= 0 Then
    '           ' At least one Win32_ComputerSystem instance was retrieved successfully
    '           intReturnCode = GetTimeZoneInstancesUsingWMINamespace(arrTimeZoneInstances, objSWbemServicesWMINamespace)
    '           If intReturnCode >= 0 Then
    '               ' At least one Win32_TimeZone instance was retrieved successfully
    '               intReturnCode = ConvertStringCIMDATETIMEToVTDATELocalTimeUsingComputerSystemAndTimeZoneInstances(datetimeOutput, strCIMDATETIMEInput, arrComputerSystemInstances, arrTimeZoneInstances)
    '               If intReturnCode >= 0 Then
    '                   ' Conversion completed successfully
    '                   ' On a computer in the Central (US) Daylight Time (GMT-5) time zone,
    '                   ' set to the en-US culture/date format, datetimeOutput is 4/21/2021
    '                   ' 1:00:00 AM
    '               End If
    '           End If
    '       End If
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
    ' TestObjectIsDateTimeContainingData()
    ' GetCurrentEffectiveTimeZoneUTCOffsetInMinutesUsingComputerSystemAndTimeZoneInstances()
    ' GetUTCOffsetForDateInLocalTimeZoneUsingTimeZoneInstances()
    ' TestObjectForData()
    'endregion DependsOn ####################################################

    Dim intFunctionReturn
    Dim intReturnMultiplier
    Dim intReturnCode
    Dim datetimeOutputToReturn
    Dim datetimeOldOutputToReturn

    Dim strTemp
    Dim intYear
    Dim intMonth
    Dim intDay
    Dim intHour
    Dim intMinute
    Dim intSecond
    Dim boolErrorOccurred
    Dim boolMinorErrorOccurred
    Dim objSWbemDateTime
    Dim boolComputerSystemWorked
    Dim boolTimeZoneWorked
    Dim timedateNow
    Dim intCurrentUTCOffset
    Dim intSpecifiedUTCOffset
    Dim intLocalComputerTimeZoneUTCOffsetAtSpecifiedDate
    Dim arrTemp
    Dim intTemp

    intFunctionReturn = 0
    intReturnMultiplier = 128 * 8 * 2 * 8

    If TestObjectIsStringContainingData(strCIMDATETIMEInput) <> True Then
        intFunctionReturn = intFunctionReturn + (-1 * intReturnMultiplier)
    Else
        ' First, try the WbemScripting.SWbemDateTime method
        boolErrorOccurred = False
        On Error Resume Next
        Set objSWbemDateTime = CreateObject("WbemScripting.SWbemDateTime")
        If Err Then
            On Error Goto 0
            Err.Clear
            boolErrorOccurred = True
        Else
            objSWbemDateTime.Value = strCIMDATETIMEInput
            If Err Then
                On Error Goto 0
                Err.Clear
                boolErrorOccurred = True
            Else
                datetimeOutputToReturn = objSWbemDateTime.GetVarDate()
                If Err Then
                    On Error Goto 0
                    Err.Clear
                    boolErrorOccurred = True
                Else
                    On Error Goto 0
                    If TestObjectIsDateTimeContainingData(datetimeOutputToReturn) <> True Then
                        boolErrorOccurred = True
                    End If
                End If
            End If
        End If

        ' Next, if the first method failed, try the conversion method using Win32_TimeZone
        If boolErrorOccurred = True Then
            boolErrorOccurred = False
            
            If boolErrorOccurred = False Then
                If Len(strCIMDATETIMEInput) < 8 Then
                    boolErrorOccurred = True
                End If
            End If

            ' Convert the year
            If boolErrorOccurred = False Then
                On Error Resume Next
                strTemp = Left(strCIMDATETIMEInput, 4)
                If Err Then
                    On Error Goto 0
                    Err.Clear
                    boolErrorOccurred = True
                Else
                    intYear = CInt(strTemp)
                    If Err Then
                        On Error Goto 0
                        Err.Clear
                        boolErrorOccurred = True
                    Else
                        On Error Goto 0
                    End If
                End If
            End If

            ' Convert the month
            If boolErrorOccurred = False Then
                On Error Resume Next
                strTemp = Mid(strCIMDATETIMEInput, 5, 2)
                If Err Then
                    On Error Goto 0
                    Err.Clear
                    boolErrorOccurred = True
                Else
                    intMonth = CInt(strTemp)
                    If Err Then
                        On Error Goto 0
                        Err.Clear
                        boolErrorOccurred = True
                    Else
                        On Error Goto 0
                        If intMonth < 1 Or intMonth > 12 Then
                            boolErrorOccurred = True
                        End If
                    End If
                End If
            End If

            ' Convert the day
            If boolErrorOccurred = False Then
                On Error Resume Next
                strTemp = Mid(strCIMDATETIMEInput, 7, 2)
                If Err Then
                    On Error Goto 0
                    Err.Clear
                    boolErrorOccurred = True
                Else
                    intDay = CInt(strTemp)
                    If Err Then
                        On Error Goto 0
                        Err.Clear
                        boolErrorOccurred = True
                    Else
                        On Error Goto 0
                        If intDay < 1 Or intDay > 31 Then
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
                strTemp = Mid(strCIMDATETIMEInput, 9, 2)
                If Err Then
                    On Error Goto 0
                    Err.Clear
                    boolMinorErrorOccurred = True
                Else
                    intHour = CInt(strTemp)
                    If Err Then
                        On Error Goto 0
                        Err.Clear
                        boolMinorErrorOccurred = True
                    Else
                        On Error Goto 0
                        If intHour < 0 Or intHour > 23 Then
                            boolMinorErrorOccurred = True
                        End If
                    End If
                End If
            End If

            ' Convert the minute
            If boolErrorOccurred = False And boolMinorErrorOccurred = False Then
                On Error Resume Next
                strTemp = Mid(strCIMDATETIMEInput, 11, 2)
                If Err Then
                    On Error Goto 0
                    Err.Clear
                    boolMinorErrorOccurred = True
                Else
                    intMinute = CInt(strTemp)
                    If Err Then
                        On Error Goto 0
                        Err.Clear
                        boolMinorErrorOccurred = True
                    Else
                        On Error Goto 0
                        If intMinute < 0 Or intMinute > 59 Then
                            boolMinorErrorOccurred = True
                        End If
                    End If
                End If
            End If

            ' Convert the second
            If boolErrorOccurred = False And boolMinorErrorOccurred = False Then
                On Error Resume Next
                strTemp = Mid(strCIMDATETIMEInput, 13, 2)
                If Err Then
                    On Error Goto 0
                    Err.Clear
                    boolMinorErrorOccurred = True
                Else
                    intSecond = CInt(strTemp)
                    If Err Then
                        On Error Goto 0
                        Err.Clear
                        boolMinorErrorOccurred = True
                    Else
                        On Error Goto 0
                        If intSecond < 0 Or intSecond > 59 Then
                            boolMinorErrorOccurred = True
                        End If
                    End If
                End If
            End If

            ' Build the datetime object
            If boolErrorOccurred = False Then
                On Error Resume Next
                datetimeOutputToReturn = DateSerial(intYear, intMonth, intDay)
                If Err Then
                    On Error Goto 0
                    Err.Clear
                    boolErrorOccurred = True
                Else
                    On Error Goto 0
                    If boolMinorErrorOccurred = False Then
                        datetimeOldOutputToReturn = datetimeOutputToReturn
                        On Error Resume Next
                        datetimeOutputToReturn = datetimeOutputToReturn + TimeSerial(intHour, intMinute, intSecond)
                        If Err Then
                            On Error Goto 0
                            Err.Clear
                            boolMinorErrorOccurred = True
                            datetimeOutputToReturn = datetimeOldOutputToReturn
                        Else
                            On Error Goto 0
                        End If
                    End If
                End If
            End If

            ' Get and apply time zone offsets
            If boolErrorOccurred = False Then
                boolComputerSystemWorked = False
                boolTimeZoneWorked = False
                intCurrentUTCOffset = Null
                intSpecifiedUTCOffset = Null
                intLocalComputerTimeZoneUTCOffsetAtSpecifiedDate = Null

                ' Check arrComputerSystemInstances
                If TestObjectForData(arrComputerSystemInstances) <> True Then
                    boolComputerSystemWorked = False
                Else
                    On Error Resume Next
                    intTemp = arrComputerSystemInstances.Count
                    If Err Then
                        On Error Goto 0
                        Err.Clear
                        boolComputerSystemWorked = False
                    Else
                        On Error Goto 0
                        If TestObjectIsAnyTypeOfInteger(intTemp) <> True Then
                            boolComputerSystemWorked = False
                        Else
                            If intTemp <= 0 Then
                                boolComputerSystemWorked = False
                            Else
                                boolComputerSystemWorked = True
                            End If
                        End If
                    End If
                End If

                ' Check arrTimeZoneInstances
                If TestObjectForData(arrTimeZoneInstances) <> True Then
                    boolTimeZoneWorked = False
                Else
                    On Error Resume Next
                    intTemp = arrTimeZoneInstances.Count
                    If Err Then
                        On Error Goto 0
                        Err.Clear
                        boolTimeZoneWorked = False
                    Else
                        On Error Goto 0
                        If TestObjectIsAnyTypeOfInteger(intTemp) <> True Then
                            boolTimeZoneWorked = False
                        Else
                            If intTemp <= 0 Then
                                boolTimeZoneWorked = False
                            Else
                                boolTimeZoneWorked = True
                            End If
                        End If
                    End If
                End If

                ' Populate intCurrentUTCOffset
                If boolComputerSystemWorked = True Or boolTimeZoneWorked = True Then
                    ' We have at least one of arrComputerSystemInstances or
                    ' arrTimeZoneInstances, which means we can try
                    ' GetCurrentEffectiveTimeZoneUTCOffsetInMinutesUsingComputerSystemAndTimeZoneInstances()
                    intReturnCode = GetCurrentEffectiveTimeZoneUTCOffsetInMinutesUsingComputerSystemAndTimeZoneInstances(intTemp, arrComputerSystemInstances, arrTimeZoneInstances)
                    If intReturnCode >= 0 Then
                        intCurrentUTCOffset = intTemp
                    End If
                End If

                ' Populate intLocalComputerTimeZoneUTCOffsetAtSpecifiedDate
                If boolTimeZoneWorked = True Then
                    ' We have arrTimeZoneInstances, which means we can try
                    ' GetUTCOffsetForDateInLocalTimeZoneUsingTimeZoneInstances()
                    intReturnCode = GetUTCOffsetForDateInLocalTimeZoneUsingTimeZoneInstances(intTemp, datetimeOutputToReturn, arrTimeZoneInstances)
                    If intReturnCode = 0 Then
                        intLocalComputerTimeZoneUTCOffsetAtSpecifiedDate = intTemp
                    End If
                End If

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

                ' intCurrentUTCOffset, intSpecifiedUTCOffset, and
                ' intLocalComputerTimeZoneUTCOffsetAtSpecifiedDate are all either Null or
                ' populated correctly.
                datetimeOldOutputToReturn = datetimeOutputToReturn
                If TestObjectForData(intLocalComputerTimeZoneUTCOffsetAtSpecifiedDate) = True Then
                    ' Use intLocalComputerTimeZoneUTCOffsetAtSpecifiedDate
                    If TestObjectForData(intSpecifiedUTCOffset) = True Then
                        ' Use intLocalComputerTimeZoneUTCOffsetAtSpecifiedDate and
                        ' intSpecifiedUTCOffset
                        On Error Resume Next
                        datetimeOutputToReturn = DateAdd("n", intLocalComputerTimeZoneUTCOffsetAtSpecifiedDate - intSpecifiedUTCOffset, datetimeOutputToReturn)
                        If Err Then
                            On Error Goto 0
                            Err.Clear
                            boolMinorErrorOccurred = True
                            datetimeOutputToReturn = datetimeOldOutputToReturn
                        Else
                            On Error Goto 0
                        End If
                    Else
                        ' Use intLocalComputerTimeZoneUTCOffsetAtSpecifiedDate by itself
                        On Error Resume Next
                        datetimeOutputToReturn = DateAdd("n", intLocalComputerTimeZoneUTCOffsetAtSpecifiedDate, datetimeOutputToReturn)
                        If Err Then
                            On Error Goto 0
                            Err.Clear
                            boolMinorErrorOccurred = True
                            datetimeOutputToReturn = datetimeOldOutputToReturn
                        Else
                            On Error Goto 0
                        End If
                    End If
                Else
                    If TestObjectForData(intCurrentUTCOffset) = True Then
                        ' Use intCurrentUTCOffset
                        If TestObjectForData(intSpecifiedUTCOffset) = True Then
                            ' Use intCurrentUTCOffset and intSpecifiedUTCOffset
                            On Error Resume Next
                            datetimeOutputToReturn = DateAdd("n", intCurrentUTCOffset - intSpecifiedUTCOffset, datetimeOutputToReturn)
                            If Err Then
                                On Error Goto 0
                                Err.Clear
                                boolMinorErrorOccurred = True
                                datetimeOutputToReturn = datetimeOldOutputToReturn
                            Else
                                On Error Goto 0
                            End If
                        Else
                            ' Use intCurrentUTCOffset by itself
                            On Error Resume Next
                            datetimeOutputToReturn = DateAdd("n", intCurrentUTCOffset, datetimeOutputToReturn)
                            If Err Then
                                On Error Goto 0
                                Err.Clear
                                boolMinorErrorOccurred = True
                                datetimeOutputToReturn = datetimeOldOutputToReturn
                            Else
                                On Error Goto 0
                            End If
                        End If
                    Else
                        ' We do not know the local computer's time zone adjustment from UTC
                        ' Don't make any adjustments
                    End If
                End If
            End If
            If boolErrorOccurred = False Then
                intFunctionReturn = 1
            End If
        End If
    End If

    If intFunctionReturn >= 0 Then
        datetimeOutput = datetimeOutputToReturn
    End If

    ConvertStringCIMDATETIMEToVTDATELocalTimeUsingComputerSystemAndTimeZoneInstances = intFunctionReturn
End Function
