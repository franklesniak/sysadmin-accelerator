Function GetBIOSReleaseDateCIMDATETIMEString(ByRef strBIOSReleaseDate)
    'region FunctionMetadata ####################################################
    ' This function obtains the computer's systems management BIOS release date in DMTF
    ' CIM_DATETIME string format, if available and configured by the computer's manufacturer.
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
    ' The function takes one positional argument (strBIOSReleaseDate), which is populated upon
    ' success with a string in CIM_DATETIME format (see above) containing the computer's
    ' systems management BIOS release date. The systems management BIOS release date is
    ' equivalent to the Win32_BIOS object property ReleaseDate
    '
    ' The function returns a 0 if the systems management BIOS release date string was obtained
    ' successfully. It returns a negative integer if an error occurred retrieving it. Finally,
    ' it returns a positive integer if the systems management BIOS release date string was
    ' obtained, but multiple BIOS instances were present that contained data for the systems
    ' management BIOS release date string. When this happens, only the first Win32_BIOS
    ' instance containing data for the systems management BIOS release date string is used.
    '
    ' Example:
    '   intReturnCode = GetBIOSReleaseDateCIMDATETIMEString(strBIOSReleaseDate)
    '   If intReturnCode >= 0 Then
    '       ' The systems management BIOS release date string was retrieved successfully and is
    '       ' stored in strBIOSReleaseDate
    '   End If
    '
    ' Version: 1.0.20210711.0
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
    ' GetBIOSInstances()
    ' GetBIOSReleaseDateCIMDATETIMEStringUsingBIOSInstances()
    'endregion DependsOn ####################################################

    Dim intFunctionReturn
    Dim arrBIOSInstances
    Dim strResult

    intFunctionReturn = 0

    intFunctionReturn = GetBIOSInstances(arrBIOSInstances)
    If intFunctionReturn >= 0 Then
        ' At least one Win32_BIOS instance was retrieved successfully
        intFunctionReturn = GetBIOSReleaseDateCIMDATETIMEStringUsingBIOSInstances(strResult, arrBIOSInstances)
        If intFunctionReturn >= 0 Then
            ' The computer's BIOS release date was retrieved successfully and is stored in
            ' strResult
            strBIOSReleaseDate = strResult
        End If
    End If
    
    GetBIOSReleaseDateCIMDATETIMEString = intFunctionReturn
End Function
