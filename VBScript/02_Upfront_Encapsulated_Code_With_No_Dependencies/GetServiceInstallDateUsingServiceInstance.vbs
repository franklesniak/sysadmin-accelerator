Function GetServiceInstallDateUsingServiceInstance(ByRef strServiceInstallDate, ByVal objServiceInstance)
    'region FunctionMetadata #######################################################
    ' Assuming that objServiceInstance represents an instance of the WMI service class
    ' (of type Win32_Service), this function obtains the service's installation date
    ' and time in DMTF CIM_DATETIME string format.
    '
    ' Note: the author believes that this property is always null, therefore this
    ' function would not return any data.
    '
    ' A CIM_DATETIME object is a string in the following format:
    ' yyyymmddHHMMSS.mmmmmmsUUU
    ' yyyy = Four-digit year (0000 through 9999)
    ' mm = Two-digit month (01 through 12)
    ' dd = Two-digit day of the month (01 through 31). This value must be appropriate
    '      for the month. For example, February 31 is invalid
    ' HH = Two-digit hour of the day using the 24-hour clock (00 through 23)
    ' MM = Two-digit minute in the hour (00 through 59)
    ' SS = Two-digit number of seconds in the minute (00 through 59)
    ' mmmmmm = Six-digit number of microseconds in the second (000000 through 999999).
    '          This field must always be present to preserve the fixed-length nature of
    '          the string
    ' s = Plus sign (+) or minus sign (-) to indicate a positive or negative offset
    '     from Universal Time Coordinates (UTC)
    ' UUU = Three-digit offset indicating the number of minutes that the originating
    '       time zone deviates from UTC
    '
    ' The function takes two positional arguments:
    '  - The first argument (strServiceInstallDate) is populated upon success with a
    '    string in CIM_DATETIME format (see above) containing the date and time when
    '    the service was installed
    '  - The second argument (objServiceInstance) is an array/collection of objects of
    '    class Win32_OperatingSystem
    '
    ' The function returns a 0 if the function could determine the value of the
    ' service's installation date and time. It returns a negative integer if an error
    ' occurred determining the service's installation date and time.
    '
    ' Example:
    '   intReturnCode = GetServiceInstances(arrServiceInstances)
    '   If intReturnCode > 0 Then
    '       ' At least one Win32_Service instance was retrieved successfully
    '       For Each objServiceInstance In arrServiceInstances
    '           intReturnCode = GetServiceInstallDateUsingServiceInstance(strServiceInstallDate, objServiceInstance)
    '           If intReturnCode = 0 Then
    '               ' The service installation date string was retrieved successfully
    '               ' and is stored in strServiceInstallDate
    '           End If
    '       Next
    '   End If
    '
    ' Version: 1.0.20230815.0
    'endregion FunctionMetadata #######################################################

    'region License ################################################################
    ' Copyright 2023 Frank Lesniak
    '
    ' Permission is hereby granted, free of charge, to any person obtaining a copy of
    ' this software and associated documentation files (the "Software"), to deal in the
    ' Software without restriction, including without limitation the rights to use,
    ' copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the
    ' Software, and to permit persons to whom the Software is furnished to do so,
    ' subject to the following conditions:
    '
    ' The above copyright notice and this permission notice shall be included in all
    ' copies or substantial portions of the Software.
    '
    ' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    ' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
    ' FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR
    ' COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN
    ' AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
    ' WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
    'endregion License ################################################################

    'region DownloadLocationNotice #################################################
    ' The most up-to-date version of this script can be found on the author's GitHub
    ' repository at https://github.com/franklesniak/sysadmin-accelerator
    'endregion DownloadLocationNotice #################################################

    'region Acknowledgements #######################################################
    ' None!
    'endregion Acknowledgements #######################################################

    'region DependsOn ##############################################################
    ' TestObjectForData()
    ' TestObjectIsStringContainingData()
    'endregion DependsOn ##############################################################

    Dim intFunctionReturn
    Dim intReturnMultiplier
    Dim strInterimResult
    Dim strResultToReturn

    Err.Clear

    intFunctionReturn = 0
    intReturnMultiplier = 128
    strInterimResult = ""
    strResultToReturn = ""

    If TestObjectForData(objServiceInstance) <> True Then
        intFunctionReturn = intFunctionReturn + (-1 * intReturnMultiplier)
    Else
        On Error Resume Next
        strInterimResult = objServiceInstance.InstallDate
        If Err Then
            Err.Clear
            On Error GoTo 0
            intFunctionReturn = intFunctionReturn + (-2 * intReturnMultiplier)
        Else
            On Error GoTo 0
            If TestObjectIsStringContainingData(strInterimResult) <> True Then
                intFunctionReturn = intFunctionReturn + (-3 * intReturnMultiplier)
            Else
                strResultToReturn = strInterimResult
            End If
        End If
    End If

    If intFunctionReturn >= 0 Then
        strServiceInstallDate = strResultToReturn
    End If
    
    GetServiceInstallDateUsingServiceInstance = intFunctionReturn
End Function
