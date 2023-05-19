Function GetBIOSReleaseDateCIMDATETIMEStringUsingBIOSInstances(ByRef strBIOSReleaseDate, ByVal arrBIOSInstances)
    'region FunctionMetadata ####################################################
    ' Assuming that arrBIOSInstances represents an array / collection of the available BIOS
    ' instances (of type Win32_BIOS), this function obtains the computer's systems management
    ' BIOS release date in DMTF CIM_DATETIME string format, if available and configured by the
    ' computer's manufacturer.
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
    '  - The first argument (strBIOSReleaseDate) is populated upon success with a string
    '    in CIM_DATETIME format (see above) containing the computer's systems management BIOS
    '    release date. The systems management BIOS release date is equivalent to the Win32_BIOS
    '    object property ReleaseDate
    '  - The second argument (arrBIOSInstances) is an array/collection of objects of class
    '    Win32_BIOS
    '
    ' The function returns a 0 if the systems management BIOS release date string was obtained
    ' successfully. It returns a negative integer if an error occurred retrieving it. Finally,
    ' it returns a positive integer if the systems management BIOS release date string was
    ' obtained, but multiple BIOS instances were present that contained data for the systems
    ' management BIOS release date string. When this happens, only the first Win32_BIOS
    ' instance containing data for the systems management BIOS release date string is used.
    '
    ' Example:
    '   intReturnCode = GetBIOSInstances(arrBIOSInstances)
    '   If intReturnCode >= 0 Then
    '       ' At least one Win32_BIOS instance was retrieved successfully
    '       intReturnCode = GetBIOSReleaseDateCIMDATETIMEStringUsingBIOSInstances(strBIOSReleaseDate, arrBIOSInstances)
    '       If intReturnCode >= 0 Then
    '           ' The systems management BIOS release date string was retrieved successfully
    '           ' and is stored in strBIOSReleaseDate
    '       End If
    '   End If
    '
    ' Version: 1.1.20230518.0
    'endregion FunctionMetadata ####################################################

    'region License ####################################################
    ' Copyright 2023 Frank Lesniak
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
    ' Microsoft, for publishing the document reference for Win32_BIOS:
    ' https://docs.microsoft.com/en-us/windows/win32/cimwin32prov/win32-bios
    '
    ' Jonathon Reinhart on StackExchange, who help confirm that SMBIOSBIOSVersion is an
    ' important BIOS attribute to retrieve. The examples are for Linux but helped the author to
    ' compare a cross-platform example.
    '
    ' The Ubuntu wiki also helped confirm that SMBIOSBIOSVersion is an important attribute,
    ' even though Ubuntu is a Linux platform - not Windows.
    'endregion Acknowledgements ####################################################

    'region DependsOn ####################################################
    ' TestObjectForData()
    ' TestObjectIsAnyTypeOfInteger()
    ' TestObjectIsStringContainingData()
    'endregion DependsOn ####################################################

    Dim intFunctionReturn
    Dim intReturnMultiplier

    Dim intTemp
    Dim strInterimResult
    Dim objBIOSInstance
    Dim strOldInterimResult
    Dim strResultToReturn
    Dim intCountOfBIOSes

    Err.Clear

    intFunctionReturn = 0
    intReturnMultiplier = 128
    strInterimResult = ""
    strResultToReturn = ""
    intCountOfBIOSes = 0

    If TestObjectForData(arrBIOSInstances) <> True Then
        intFunctionReturn = intFunctionReturn + (-1 * intReturnMultiplier)
    Else
        On Error Resume Next
        intTemp = arrBIOSInstances.Count
        If Err Then
            On Error Goto 0
            Err.Clear
            intFunctionReturn = intFunctionReturn + (-2 * intReturnMultiplier)
        Else
            On Error Goto 0
            If TestObjectIsAnyTypeOfInteger(intTemp) = False Then
                intFunctionReturn = intFunctionReturn + (-3 * intReturnMultiplier)
            Else
                If intTemp < 0 Then
                    intFunctionReturn = intFunctionReturn + (-4 * intReturnMultiplier)
                ElseIf intTemp = 0 Then
                    intFunctionReturn = intFunctionReturn + (-5 * intReturnMultiplier)
                Else
                    On Error Resume Next
                    For Each objBIOSInstance in arrBIOSInstances
                        If Err Then
                            Err.Clear
                        Else
                            strOldInterimResult = strInterimResult
                            strInterimResult = objBIOSInstance.ReleaseDate
                            If Err Then
                                Err.Clear
                                strInterimResult = strOldInterimResult
                            Else
                                If TestObjectForData(strInterimResult) <> True Then
                                    strInterimResult = strOldInterimResult
                                Else
                                    ' Found a result with real release date data
                                    If TestObjectIsStringContainingData(strResultToReturn) = False Then
                                        strResultToReturn = strInterimResult
                                    End If
                                    intCountOfBIOSes = intCountOfBIOSes + 1
                                End If
                            End If
                        End If
                    Next
                    On Error Goto 0
                    If Err Then
                        Err.Clear
                    End If
                End If
            End If
        End If
    End If

    If intFunctionReturn >= 0 Then
        ' No error has occurred yet
        If intCountOfBIOSes = 0 Then
            ' No result found
            intFunctionReturn = intFunctionReturn + (-5 * intReturnMultiplier)
        Else
            intFunctionReturn = intCountOfBIOSes - 1
        End If
    End If

    If intFunctionReturn >= 0 Then
        strBIOSReleaseDate = strResultToReturn
    End If
    
    GetBIOSReleaseDateCIMDATETIMEStringUsingBIOSInstances = intFunctionReturn
End Function
