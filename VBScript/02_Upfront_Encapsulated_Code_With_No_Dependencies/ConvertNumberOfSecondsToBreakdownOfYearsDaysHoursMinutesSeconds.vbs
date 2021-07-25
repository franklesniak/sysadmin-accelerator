Function ConvertNumberOfSecondsToBreakdownOfYearsDaysHoursMinutesSeconds(ByRef intYears, ByRef intDays, ByRef intHours, ByRef intMinutes, ByRef objOutputSeconds, ByVal objInputTotalSeconds)
    'region FunctionMetadata ####################################################
    ' Assuming that objInputTotalSeconds is an integer that indicates a total number of
    ' seconds, this function breaks down the seconds into the integer number of years, days,
    ' hours, minutes, and seconds that add up to the specified number of total seconds.
    '
    ' NOTE: For simplicity, and to account for leap years, this function assumes that there are
    ' 365.25 days in a year
    '
    ' The function takes six positional arguments:
    '  - The first argument (intYears) is populated with the rounded-down number of years that
    '    comprise the number of seconds specified in the sixth argument
    '    (objInputTotalSeconds). For simplicity, and to account for leap years, this function
    '    assumes that there are 365.25 * 24 * 60 * 60 = 31557600 seconds in a year. For
    '    example, if the sixth argument (objInputTotalSeconds) is 2047483645, then the first
    '    argument (intYears) is 64. That's because 2047483645 / 31557600 = 64.88084, which
    '    is 64 after rounding-down to the nearest whole year.
    '  - The second argument (intDays) is populated with the rounded-down number of days
    '    remaining after subtracting the number of years that comprise the number of seconds
    '    specified in the sixth argument (objInputTotalSeconds). For example, if the sixth
    '    argument (objInputTotalSeconds) is 2047483645, then the first argument (intYears) is
    '    64 as noted previously and the second argument (intDays) is 321. That's because 64
    '    years is (64 * 31557600) = 2019686400 seconds. Then, 2047483645 - 2019686400 leaves a
    '    remainder of 27797245 seconds. 27797245 / 60 / 60 / 24 = 321.72737, which is 321 after
    '    rounding down
    '  - The third argument (intHours) is populated with the rounded-down number of hours
    '    remaining after subtracting the number of years that comprise the number of seconds
    '    specified in the sixth argument (objInputTotalSeconds), then subracting the
    '    remaining number of days. For example, if the sixth argument (objInputTotalSeconds)
    '    is 2047483645, then the first argument (intYears) is 64 as noted previously, and the
    '    second argument (intDays) is 321 as noted previously, and the third argument
    '    (intHours) is 17. That's because 64 years and 321 days is (64 * 31557600) +
    '    (321 * 24 * 60 * 60) = 2047420800 seconds. Then, 2047483645 - 2047420800 leaves a
    '    remainder of 62845 seconds. 62845 / 60 / 60 = 17.45694, which is 17 after rounding
    '    down
    '  - The fourth argument (intMinutes) is populated with the rounded-down number of minutes
    '    remaining after subtracting the number of years that comprise the number of seconds
    '    specified in the sixth argument (objInputTotalSeconds), then subracting the remaining
    '    number of days, then subtracting the remaining number of hours. For example, if the
    '    sixth argument (objInputTotalSeconds) is 2047483645, then the first argument
    '    (intYears) is 64 as noted previously, the second argument (intDays) is 321 as noted
    '    previously, the third argument (intHours) is 17 as noted previously, and the fourth
    '    argument (intMinutes) is 27. That's because 64 years, 17 days, and 8 hours is
    '    (64 * 31557600) + (321 * 24 * 60 * 60) + (17 * 60 * 60) = 2047482000 seconds. Then,
    '    2047483645 - 2047482000 leaves a remainder of 1645 seconds.
    '    1645 / 60 = 27.41667, which is 27 after rounding down
    '  - The fifth argument (objOutputSeconds) is populated with the number of seconds
    '    remaining after subtracting the number of years that comprise the number of seconds
    '    specified in the sixth argument (objInputTotalSeconds), then subracting the
    '    remaining number of days, then subtracting the remaining number of hours, then
    '    subtracting the remaining number of minutes. For example, if the sixth argument
    '    (objInputTotalSeconds) is 2047483645, then the first argument (intYears) is 64 as
    '    noted previously, the second argument (intDays) is 321 as noted previously, the third
    '    argument (intHours) is 17 as noted previously, the fourth argument (intMinutes) is 27
    '    as noted previously, and the fifth argument (objOutputSeconds) is 25. That's because
    '    64 years, 321 days, 17 hours, and 27 minutes is (64 * 31557600) +
    '    (321 * 24 * 60 * 60) + (17 * 60 * 60) + (27 * 60) = 2047483620 seconds. Then,
    '    2047483645 - 2047483620 leaves a remainder of 25 seconds
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
    '       ' 2047483645 is equivalent to 64 years, 321 days, 17 hours, 27 minutes, and 25
    '       ' seconds
    '   End If
    '
    ' Version: 1.0.20210724.0
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

    Const NUM_SECONDS_IN_YEAR = 31557600 ' = 365.25 * 24 * 60 * 60
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
