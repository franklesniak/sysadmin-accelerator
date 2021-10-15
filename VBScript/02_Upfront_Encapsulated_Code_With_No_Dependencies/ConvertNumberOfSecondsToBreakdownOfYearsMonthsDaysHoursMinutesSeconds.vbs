Function ConvertNumberOfSecondsToBreakdownOfYearsMonthsDaysHoursMinutesSeconds(ByRef intYears, ByRef intMonths, ByRef intDays, ByRef intHours, ByRef intMinutes, ByRef objOutputSeconds, ByVal objInputTotalSeconds)
    'region FunctionMetadata ####################################################
    ' Assuming that objInputTotalSeconds is an integer that indicates a total number of
    ' seconds, this function breaks down the seconds into the integer number of years, months,
    ' days, hours, minutes, and seconds that add up to the specified number of total seconds.
    '
    ' NOTE: For simplicity, and to account for leap years, this function assumes that there are
    ' 365.2425 days in a year, and 365.2425 / 12 = 30.436875 days in a month
    '
    ' The function takes seven positional arguments:
    '  - The first argument (intYears) is populated with the rounded-down number of years that
    '    comprise the number of seconds specified in the seventh argument
    '    (objInputTotalSeconds). For simplicity, and to account for leap years, this function
    '    assumes that there are 365.2425 * 24 * 60 * 60 = 31556952 seconds in a year. For
    '    example, if the seventh argument (objInputTotalSeconds) is 2047483645, then the first
    '    argument (intYears) is 64. That's because 2047483645 / 31556952 = 64.88217, which
    '    is 64 after rounding-down to the nearest whole year.
    '  - The second argument (intMonths) is populated with the rounded-down number of months
    '    remaining after subtracting the number of years that comprise the number of seconds
    '    specified in the seventh argument (objInputTotalSeconds). For simplicity, and to
    '    account for both leap years and the variable number of days in a month, this function
    '    assumes that there are (365.2425 / 12) * 24 * 60 * 60 = 2629746 seconds in a month. For
    '    example, if the seventh argument (objInputTotalSeconds) is 2047483645, then the first
    '    argument (intYears) is 64 as noted previously, and the second argument (intMonths) is
    '    10. That's because 64 years is 64 * 31556952 = 2019644928 seconds. Then,
    '    2047483645 - 2019644928 leaves a remainder of 27838717 seconds.
    '    27838717 / 2629746 = 10.58609, which is 10 after rounding down
    '  - The third argument (intDays) is populated with the rounded-down number of days
    '    remaining after subtracting the number of years that comprise the number of seconds
    '    specified in the seventh argument (objInputTotalSeconds), then subracting the
    '    remaining number of months. For example, if the seventh argument
    '    (objInputTotalSeconds) is 2047483645, then the first argument (intYears) is 64 as
    '    noted previously, and the second argument (intMonths) is 10 as noted previously, and
    '    the third argument (intDays) is 17. That's because 64 years and 10 months is
    '    (64 * 31556952) + (10 * 2629746) = 2045942388 seconds. Then, 2047483645 - 2045942388
    '    leaves a remainder of 1541257 seconds. 1541257 / 60 / 60 / 24 = 17.83862, which is 17
    '    after rounding down
    '  - The fourth argument (intHours) is populated with the rounded-down number of hours
    '    remaining after subtracting the number of years that comprise the number of seconds
    '    specified in the seventh argument (objInputTotalSeconds), then subracting the
    '    remaining number of months, then subtracting the remaining number of days. For
    '    example, if the seventh argument (objInputTotalSeconds) is 2047483645, then the first
    '    argument (intYears) is 64 as noted previously, and the second argument (intMonths) is
    '    10 as noted previously, the third argument (intDays) is 17 as noted previously, and
    '    the fourth argument (intHours) is 20. That's because 64 years, 10 months, and 17 days
    '    is (64 * 31556952) + (10 * 2629746) + (17 * 24 * 60 * 60) = 2047411188 seconds. Then,
    '    2047483645 - 2047411188 leaves a remainder of 72457 seconds.
    '    72457 / 60 / 60 = 20.12694, which is 20 after rounding down
    '  - The fifth argument (intMinutes) is populated with the rounded-down number of minutes
    '    remaining after subtracting the number of years that comprise the number of seconds
    '    specified in the seventh argument (objInputTotalSeconds), then subracting the
    '    remaining number of months, then subtracting the remaining number of days, then
    '    subtracting the remaining number of hours. For example, if the seventh argument
    '    (objInputTotalSeconds) is 2047483645, then the first argument (intYears) is 64 as
    '    noted previously, the second argument (intMonths) is 10 as noted previously, the third
    '    argument (intDays) is 17 as noted previously, the fourth argument (intHours) is 20 as
    '    noted previously, and the firth argument (intMinutes) is 7. That's because 64 years,
    '    10 months, 17 days, and 20 hours is (64 * 31556952) + (10 * 2629746) +
    '    (17 * 24 * 60 * 60) + (20 * 60 * 60) = 2047483188 seconds. Then,
    '    2047483645 - 2047483188 leaves a remainder of 457 seconds.
    '    457 / 60 = 7.61667, which is 7 after rounding down
    '  - The sixth argument (objOutputSeconds) is populated with the number of seconds
    '    remaining after subtracting the number of years that comprise the number of seconds
    '    specified in the seventh argument (objInputTotalSeconds), then subracting the
    '    remaining number of months, then subtracting the remaining number of days, then
    '    subtracting the remaining number of hours, then subtracting the remaining number of
    '    minutes. For example, if the seventh argument (objInputTotalSeconds) is 2047483645,
    '    then the first argument (intYears) is 64 as noted previously, the second argument
    '    (intMonths) is 10 as noted previously, the third argument (intDays) is 17 as noted
    '    previously, the fourth argument (intHours) is 20 as noted previously, the firth
    '    argument (intMinutes) is 7 as noted previously, and the sixth argument
    '    (objOutputSeconds) is 37. That's because 64 years, 10 months, 17 days, 20 hours, and
    '    7 minutes is (64 * 31556952) + (10 * 2629746) + (17 * 24 * 60 * 60) + (20 * 60 * 60) +
    '    (7 * 60) = 2047483608 seconds. Then, 2047483645 - 2047483608 leaves a remainder of 37
    '    seconds
    '  - The seventh argument (objInputTotalSeconds) is the number of seconds to convert to a
    '    breakdown of years, months, days, hours, minutes, and remaining seconds
    '
    ' The function returns a 0 if the conversion of total seconds to a breakdown of years,
    ' months, days, hours, minutes, and remaining seconds was successful. The function returns
    ' a negative number if the conversion was not successful.
    '
    ' Example:
    '   intTotalSecondsToConvert = 2047483645
    '   intReturnCode = ConvertNumberOfSecondsToBreakdownOfYearsMonthsDaysHoursMinutesSeconds(intYears, intMonths, intDays, intHours, intMinutes, objOutputSeconds, intTotalSecondsToConvert)
    '   If intReturnCode = 0 Then
    '       ' Successfully converted the number of seconds to a breakdown of years, months,
    '       ' days, hours, minutes and remaining seconds
    '       WScript.Echo(intTotalSecondsToConvert & " is equivalent to " & intYears & " years, " & intMonths & " months, " & intDays & " days, " & intHours & " hours, " & intMinutes & " minutes, and " & objOutputSeconds & " seconds.")
    '       ' The script outputs:
    '       ' 2047483645 is equivalent to 64 years, 10 months, 17 days, 20 hours, 7 minutes,
    '       ' and 37 seconds
    '   End If
    '
    ' Version: 1.1.20211015.0
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
    ' TestObjectIsAnyTypeOfNumber()
    'endregion DependsOn ####################################################

    Const NUM_SECONDS_IN_YEAR = 31556952 ' = 365.2425 * 24 * 60 * 60
    Const NUM_SECONDS_IN_MONTH = 2629746 ' = (365.2425 / 12) * 24 * 60 * 60
    Const NUM_SECONDS_IN_DAY = 86400 ' = 24 * 60 * 60
    Const NUM_SECONDS_IN_HOUR = 3600 ' = 60 * 60
    Const NUM_SECONDS_IN_MINUTE = 60

    Dim intFunctionReturn
    Dim objRemainingSeconds
    Dim intWorkingYears
    Dim intWorkingMonths
    Dim intWorkingDays
    Dim intWorkingHours
    Dim intWorkingMinutes

    intFunctionReturn = 0

    If TestObjectIsAnyTypeOfNumber(objInputTotalSeconds) <> True Then
        intFunctionReturn = -1
    Else
        intWorkingYears = Int(objInputTotalSeconds / NUM_SECONDS_IN_YEAR)
        objRemainingSeconds = objInputTotalSeconds - (intWorkingYears * NUM_SECONDS_IN_YEAR)
        intWorkingMonths = Int(objRemainingSeconds / NUM_SECONDS_IN_MONTH)
        objRemainingSeconds = objRemainingSeconds - (intWorkingMonths * NUM_SECONDS_IN_MONTH)
        intWorkingDays = Int(objRemainingSeconds / NUM_SECONDS_IN_DAY)
        objRemainingSeconds = objRemainingSeconds - (intWorkingDays * NUM_SECONDS_IN_DAY)
        intWorkingHours = Int(objRemainingSeconds / NUM_SECONDS_IN_HOUR)
        objRemainingSeconds = objRemainingSeconds - (intWorkingHours * NUM_SECONDS_IN_HOUR)
        intWorkingMinutes = Int(objRemainingSeconds / NUM_SECONDS_IN_MINUTE)
        objRemainingSeconds = objRemainingSeconds - (intWorkingMinutes * NUM_SECONDS_IN_MINUTE)
    End If

    If intFunctionReturn = 0 Then
        intYears = intWorkingYears
        intMonths = intWorkingMonths
        intDays = intWorkingDays
        intHours = intWorkingHours
        intMinutes = intWorkingMinutes
        objOutputSeconds = objRemainingSeconds
    End If
    
    ConvertNumberOfSecondsToBreakdownOfYearsMonthsDaysHoursMinutesSeconds = intFunctionReturn
End Function
