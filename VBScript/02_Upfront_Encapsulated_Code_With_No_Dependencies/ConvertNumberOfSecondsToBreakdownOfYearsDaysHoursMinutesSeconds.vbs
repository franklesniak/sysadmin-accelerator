Function ConvertNumberOfSecondsToBreakdownOfYearsDaysHoursMinutesSeconds(ByRef intYears, ByRef intDays, ByRef intHours, ByRef intMinutes, ByRef objOutputSeconds, ByVal objInputTotalSeconds)
    'region FunctionMetadata ####################################################
    ' Assuming that objInputTotalSeconds is an integer that indicates a total number of
    ' seconds, this function breaks down the seconds into the integer number of years, days,
    ' hours, minutes, and seconds that add up to the specified number of total seconds.
    '
    ' NOTE: For simplicity, and to account for leap years, this function assumes that there are
    ' 365.2425 days in a year
    '
    ' The function takes six positional arguments:
    '  - The first argument (intYears) is populated with the rounded-down number of years that
    '    comprise the number of seconds specified in the sixth argument
    '    (objInputTotalSeconds). For simplicity, and to account for leap years, this function
    '    assumes that there are 365.2425 * 24 * 60 * 60 = 31556952 seconds in a year. For
    '    example, if the sixth argument (objInputTotalSeconds) is 2047483645, then the first
    '    argument (intYears) is 64. That's because 2047483645 / 31556952 = 64.88217, which
    '    is 64 after rounding-down to the nearest whole year.
    '  - The second argument (intDays) is populated with the rounded-down number of days
    '    remaining after subtracting the number of years that comprise the number of seconds
    '    specified in the sixth argument (objInputTotalSeconds). For example, if the sixth
    '    argument (objInputTotalSeconds) is 2047483645, then the first argument (intYears) is
    '    64 as noted previously and the second argument (intDays) is 322. That's because 64
    '    years is (64 * 31556952) = 2019644928 seconds. Then, 2047483645 - 2019644928 leaves a
    '    remainder of 27838717 seconds. 27838717 / 60 / 60 / 24 = 322.20737, which is 322 after
    '    rounding down
    '  - The third argument (intHours) is populated with the rounded-down number of hours
    '    remaining after subtracting the number of years that comprise the number of seconds
    '    specified in the sixth argument (objInputTotalSeconds), then subtracting the
    '    remaining number of days. For example, if the sixth argument (objInputTotalSeconds)
    '    is 2047483645, then the first argument (intYears) is 64 as noted previously, and the
    '    second argument (intDays) is 322 as noted previously, and the third argument
    '    (intHours) is 4. That's because 64 years and 322 days is (64 * 31556952) +
    '    (322 * 24 * 60 * 60) = 2047465728 seconds. Then, 2047483645 - 2047465728 leaves a
    '    remainder of 17917 seconds. 17917 / 60 / 60 = 4.97694, which is 4 after rounding
    '    down
    '  - The fourth argument (intMinutes) is populated with the rounded-down number of minutes
    '    remaining after subtracting the number of years that comprise the number of seconds
    '    specified in the sixth argument (objInputTotalSeconds), then subracting the remaining
    '    number of days, then subtracting the remaining number of hours. For example, if the
    '    sixth argument (objInputTotalSeconds) is 2047483645, then the first argument
    '    (intYears) is 64 as noted previously, the second argument (intDays) is 322 as noted
    '    previously, the third argument (intHours) is 4 as noted previously, and the fourth
    '    argument (intMinutes) is 58. That's because 64 years, 322 days, and 4 hours is
    '    (64 * 31556952) + (322 * 24 * 60 * 60) + (4 * 60 * 60) = 2047480128 seconds. Then,
    '    2047483645 - 2047480128 leaves a remainder of 3517 seconds.
    '    3517 / 60 = 58.61667, which is 58 after rounding down
    '  - The fifth argument (objOutputSeconds) is populated with the number of seconds
    '    remaining after subtracting the number of years that comprise the number of seconds
    '    specified in the sixth argument (objInputTotalSeconds), then subracting the
    '    remaining number of days, then subtracting the remaining number of hours, then
    '    subtracting the remaining number of minutes. For example, if the sixth argument
    '    (objInputTotalSeconds) is 2047483645, then the first argument (intYears) is 64 as
    '    noted previously, the second argument (intDays) is 322 as noted previously, the third
    '    argument (intHours) is 4 as noted previously, the fourth argument (intMinutes) is 58
    '    as noted previously, and the fifth argument (objOutputSeconds) is 37. That's because
    '    64 years, 322 days, 4 hours, and 58 minutes is (64 * 31556952) +
    '    (322 * 24 * 60 * 60) + (4 * 60 * 60) + (58 * 60) = 2047483608 seconds. Then,
    '    2047483645 - 2047483608 leaves a remainder of 37 seconds
    '  - The sixth argument (objInputTotalSeconds) is the number of seconds to convert to a
    '    breakdown of years, days, hours, minutes, and remaining seconds
    '
    ' The function returns a 0 if the conversion of total seconds to a breakdown of years,
    ' days, hours, minutes, and remaining seconds was successful. The function returns a
    ' negative number if the conversion was not successful.
    '
    ' Example:
    '   intTotalSecondsToConvert = 2047483645
    '   intReturnCode = ConvertNumberOfSecondsToBreakdownOfYearsDaysHoursMinutesSeconds(intYears, intDays, intHours, intMinutes, objOutputSeconds, intTotalSecondsToConvert)
    '   If intReturnCode = 0 Then
    '       ' Successfully converted the number of seconds to a breakdown of years, days,
    '       ' hours, minutes and remaining seconds
    '       WScript.Echo(intTotalSecondsToConvert & " is equivalent to " & intYears & " years, " & intDays & " days, " & intHours & " hours, " & intMinutes & " minutes, and " & objOutputSeconds & " seconds.")
    '       ' The script outputs:
    '       ' 2047483645 is equivalent to 64 years, 322 days, 4 hours, 58 minutes, and 37
    '       ' seconds
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
    Const NUM_SECONDS_IN_DAY = 86400 ' = 24 * 60 * 60
    Const NUM_SECONDS_IN_HOUR = 3600 ' = 60 * 60
    Const NUM_SECONDS_IN_MINUTE = 60

    Dim intFunctionReturn
    Dim objRemainingSeconds
    Dim intWorkingYears
    Dim intWorkingDays
    Dim intWorkingHours
    Dim intWorkingMinutes

    intFunctionReturn = 0

    If TestObjectIsAnyTypeOfNumber(objInputTotalSeconds) <> True Then
        intFunctionReturn = -1
    Else
        intWorkingYears = Int(objInputTotalSeconds / NUM_SECONDS_IN_YEAR)
        objRemainingSeconds = objInputTotalSeconds - (intWorkingYears * NUM_SECONDS_IN_YEAR)
        intWorkingDays = Int(objRemainingSeconds / NUM_SECONDS_IN_DAY)
        objRemainingSeconds = objRemainingSeconds - (intWorkingDays * NUM_SECONDS_IN_DAY)
        intWorkingHours = Int(objRemainingSeconds / NUM_SECONDS_IN_HOUR)
        objRemainingSeconds = objRemainingSeconds - (intWorkingHours * NUM_SECONDS_IN_HOUR)
        intWorkingMinutes = Int(objRemainingSeconds / NUM_SECONDS_IN_MINUTE)
        objRemainingSeconds = objRemainingSeconds - (intWorkingMinutes * NUM_SECONDS_IN_MINUTE)
    End If

    If intFunctionReturn = 0 Then
        intYears = intWorkingYears
        intDays = intWorkingDays
        intHours = intWorkingHours
        intMinutes = intWorkingMinutes
        objOutputSeconds = objRemainingSeconds
    End If
    
    ConvertNumberOfSecondsToBreakdownOfYearsDaysHoursMinutesSeconds = intFunctionReturn
End Function
