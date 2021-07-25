Function ConvertNumberToIntegerFloor(ByVal objNumberToConvert)
    'region FunctionMetadata ####################################################
    ' Safely takes a number and rounds it down, similarly to a "floor" function
    '
    ' Note: if objNumberToConvert is negative, this function rounds it down to the next-most
    ' negative integer. For example: -6.5 becomes -7.
    '
    ' Note: if the input is an integer, this function maintains the same integer type (byte,
    ' single, double, quad, etc.)
    '
    ' The function takes one argument (objNumberToConvert), which must contain a number
    ' (integer or floating point), which will be "rounded down" to the nearest integer
    '
    ' Upon success, the function returns the "rounded down" integer. If the input object is not
    ' a number, or if another error occured, the function returns 0.
    '
    ' Example 1:
    '   objNumberToConvert = 123.45
    '   intRoundedDown = ConvertNumberToIntegerFloor(objNumberToConvert)
    '   ' intRoundedDown is 123
    '
    ' Example 2:
    '   objNumberToConvert = 123.5
    '   intRoundedDown = ConvertNumberToIntegerFloor(objNumberToConvert)
    '   ' intRoundedDown is 123
    '   ' (this example is significant because VBScript normally "rounds to even", so an Int()
    '   ' or Fix() function, for example, would have returned 124
    '
    ' Example 3:
    '   objNumberToConvert = 122.5
    '   intRoundedDown = ConvertNumberToIntegerFloor(objNumberToConvert)
    '   ' intRoundedDown is 122
    '   ' (this example is significant because VBScript normally "rounds to even", so an Int()
    '   ' or Fix() function, for example, would have returned 122
    '
    ' Example 4:
    '   objNumberToConvert = "122.5"
    '   intRoundedDown = ConvertNumberToIntegerFloor(objNumberToConvert)
    '   ' intRoundedDown is 0 because "122.5" is a string, not a number
    '
    ' Version: 1.0.20210724.1
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
    ' TestObjectIsAnyTypeOfNumber()
    'endregion DependsOn ####################################################

    Dim intFunctionReturn
    Dim intCalculated

    intFunctionReturn = 0

    If TestObjectIsAnyTypeOfNumber(objNumberToConvert) = True Then
        ' Input was a number; continue
        intCalculated = Round(objNumberToConvert)

        If intCalculated > objNumberToConvert Then
            intCalculated = intCalculated - 1
        End If

        intFunctionReturn = intCalculated
    End If

    ConvertNumberToIntegerFloor = intFunctionReturn
End Function
