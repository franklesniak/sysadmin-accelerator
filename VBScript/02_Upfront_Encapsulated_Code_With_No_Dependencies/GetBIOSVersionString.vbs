Function GetBIOSVersionString(ByRef strBIOSVersion)
    'region FunctionMetadata ####################################################
    ' This function obtains the computer's systems management BIOS version number in string
    ' format, if available and configured by the computer's manufacturer.
    '
    ' The function takes one positional argument (strBIOSVersion), which is populated upon
    ' success with a string containing the computer's systems management BIOS version number in
    ' string format. The systems management BIOS version number is equivalent to the Win32_BIOS
    ' system property SMBIOSBIOSVersion
    '
    ' The function returns a 0 if the systems management BIOS version string was obtained
    ' successfully. It returns a negative integer if an error occurred retrieving it. Finally,
    ' it returns a positive integer if the systems management BIOS version string was obtained,
    ' but multiple BIOS instances were present that contained data for the systems management
    ' BIOS version string. When this happens, only the first Win32_BIOS instance containing
    ' data for the systems management BIOS version string is used.
    '
    ' Example:
    '   intReturnCode = GetBIOSVersionString(strBIOSVersion)
    '   If intReturnCode >= 0 Then
    '       ' The systems management BIOS version string was retrieved successfully and is
    '       ' stored in strBIOSVersion
    '   End If
    '
    ' Version: 1.0.20210707.0
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
    ' GetBIOSVersionStringUsingBIOSInstances()
    'endregion DependsOn ####################################################

    Dim intFunctionReturn
    Dim arrBIOSInstances
    Dim strResult

    intFunctionReturn = 0

    intFunctionReturn = GetBIOSInstances(arrBIOSInstances)
    If intFunctionReturn >= 0 Then
        ' At least one Win32_BIOS instance was retrieved successfully
        intFunctionReturn = GetBIOSVersionStringUsingBIOSInstances(strResult, arrBIOSInstances)
        If intFunctionReturn >= 0 Then
            ' The computer serial number was retrieved successfully and is stored in strResult
            strBIOSVersion = strResult
        End If
    End If
    
    GetBIOSVersionString = intFunctionReturn
End Function
