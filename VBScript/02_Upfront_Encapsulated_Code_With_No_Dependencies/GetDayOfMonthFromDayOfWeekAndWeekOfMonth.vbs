Function GetDayOfMonthFromDayOfWeekAndWeekOfMonth(ByVal intYear, ByVal intMonth, ByVal intWeekOfMonth, ByVal intDayOfWeek)
    ' Returns the day of the month for the specified day of the week, week of the month, month,
    ' and year
    '
    ' The function takes four positional arguments:
    '   The first argument (intYear) is an integer set to the current year (e.g., 2021)
    '   The second argument (intMonth) is an integer set to the current month number of the
    '       year. For example, 1 = January, 2 = February, 3 = March, etc.
    '   The third argument (intWeekOfMonth) indicates the week number of the month:
    '       1 = First week of the month
    '       2 = Second week of the month
    '       3 = Third week of the month
    '       4 = Fourth week of the month
    '       5 = Fifth or last week of the month
    '   The fourth argument (intDayOfWeek) indicates the day of the week:
    '       0 = Sunday
    '       1 = Monday
    '       2 = Tuesday
    '       3 = Wednesday
    '       4 = Thursday
    '       5 = Friday
    '       6 = Saturday
    '
    ' If successful, the function returns an integer between 1 and 31 indicating the day of the
    ' month. If an error occurred, the function returns 0.
    '
    ' Example:
    '   ' The following outputs 3, meaning July 3, 2021, which is the first (1) Saturday (6) of
    '   ' July (7), 2021
    '   WScript.Echo(GetDayOfMonthFromDayOfWeekAndWeekOfMonth(2021, 7, 1, 6))
    '   ' The following outputs 10, meaning July 10, 2021, which is the second (2) Saturday (6)
    '   ' of July (7), 2021
    '   WScript.Echo(GetDayOfMonthFromDayOfWeekAndWeekOfMonth(2021, 7, 2, 6))
    '   ' The following outputs 17, meaning July 17, 2021, which is the third (3) Saturday (6)
    '   ' of July (7), 2021
    '   WScript.Echo(GetDayOfMonthFromDayOfWeekAndWeekOfMonth(2021, 7, 3, 6))
    '   ' The following outputs 24, meaning July 24, 2021, which is the fourth (4) Saturday (6)
    '   ' of July (7), 2021
    '   WScript.Echo(GetDayOfMonthFromDayOfWeekAndWeekOfMonth(2021, 7, 4, 6))
    '   ' The following outputs 31, meaning July 31, 2021, which is the fifth (5) Saturday (6)
    '   ' of July (7), 2021
    '   WScript.Echo(GetDayOfMonthFromDayOfWeekAndWeekOfMonth(2021, 7, 5, 6))
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
    ' Rob van der Woude, who wrote a function LastDoW() that heavily inspiried this function:
    ' https://www.robvanderwoude.com/files/isdst_vbs.txt
    'endregion Acknowledgements ####################################################

    'region DependsOn ####################################################
    ' None!
    'endregion DependsOn ####################################################

    Dim intMonthDayCounter
    Dim intWeekCounter
    Dim intFunctionReturn
    Dim dateTemp
    Dim intWorkingDayOfTheWeek
    
    intFunctionReturn = 0

    intWeekCounter = 1

    For intMonthDayCounter = 1 To 31
        If intWeekOfMonth >= intWeekCounter Then
            On Error Resume Next
            dateTemp = DateSerial(intYear, intMonth, intMonthDayCounter)
            If Err Then
                On Error Goto 0
                Err.Clear
            Else
                intWorkingDayOfTheWeek = DatePart("w", dateTemp, vbSunday) - 1
                If Err Then
                    On Error Goto 0
                    Err.Clear
                Else
                    On Error Goto 0
                    If intWorkingDayOfTheWeek = intDayOfWeek Then
                        intWeekCounter = intWeekCounter + 1
                        intFunctionReturn = intMonthDayCounter
                    End If
                End If
            End If
        End If
    Next
    
    GetDayOfMonthFromDayOfWeekAndWeekOfMonth = intFunctionReturn
End Function
