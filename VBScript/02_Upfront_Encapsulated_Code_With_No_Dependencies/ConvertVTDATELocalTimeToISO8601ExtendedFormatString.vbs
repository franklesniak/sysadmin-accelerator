Function ConvertVTDATELocalTimeToISO8601ExtendedFormatString(ByRef strISO8601Output, ByVal datetimeInput)
    'region FunctionMetadata ####################################################
    ' Safely takes a VT_DATETIME (VBScript-native datetime variant) object and converts it to a
    ' string representation of the date and time, in compliance with ISO 8601's extended format
    '
    ' The function takes two positional arguments:
    '   The first argument (strISO8601Output) is set upon success to a string representation of
    '       the date and time specified by the second argument (datetimeInput), represented in
    '       compliance with ISO 8601
    '   The second argument (datetimeInput) is a VBScript-native datetime object (VT_DATE)
    '       containing a date and time in the local computer's time zone
    '
    ' The function returns 0 or a positive number if the VBScript-native datetime object
    ' (VT_DATETIME) was converted to a ISO 8601-formatted string. A return of 1 indicates a
    ' warning condition in which the local computer's time zone adjustment could not be
    ' determined. It returns a negative number if an error occurred
    '
    ' Example:
    '   datetimeDateFunctionAuthored = DateSerial(2021, 8, 8)
    '   datetimeDateFunctionAuthored = datetimeDateFunctionAuthored + TimeSerial(13, 0, 0)
    '   intReturnCode = ConvertVTDATELocalTimeToISO8601ExtendedFormatString(strISO8601Output, datetimeDateFunctionAuthored)
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
    ' ConnectLocalWMINamespace()
    ' GetComputerSystemInstancesUsingWMINamespace()
    ' GetTimeZoneInstancesUsingWMINamespace()
    ' ConvertVTDATELocalTimeToISO8601ExtendedFormatStringUsingComputerSystemAndTimeZoneInstances()
    'endregion DependsOn ####################################################

    Dim intFunctionReturn
    Dim intReturnCode
    Dim strISO8601OutputToReturn

    Dim objSWbemServicesWMINamespace
    Dim arrComputerSystemInstances
    Dim arrTimeZoneInstances

    intReturnCode = ConnectLocalWMINamespace(objSWbemServicesWMINamespace, Null, Null)
    intReturnCode = GetComputerSystemInstancesUsingWMINamespace(arrComputerSystemInstances, objSWbemServicesWMINamespace)
    intReturnCode = GetTimeZoneInstancesUsingWMINamespace(arrTimeZoneInstances, objSWbemServicesWMINamespace)
    intFunctionReturn = ConvertVTDATELocalTimeToISO8601ExtendedFormatStringUsingComputerSystemAndTimeZoneInstances(strISO8601OutputToReturn, datetimeInput, arrComputerSystemInstances, arrTimeZoneInstances)

    If intFunctionReturn >= 0 Then
        strISO8601Output = strISO8601OutputToReturn
    End If

    ConvertVTDATELocalTimeToISO8601ExtendedFormatString = intFunctionReturn
End Function
