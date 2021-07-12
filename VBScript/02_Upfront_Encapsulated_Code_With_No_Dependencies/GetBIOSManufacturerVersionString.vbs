Function GetBIOSManufacturerVersionString(ByRef strBIOSManufacturerVersion)
    'region FunctionMetadata ####################################################
    ' This function obtains the computer's BIOS version number in string format as reported by
    ' the BIOS manufacturer, if available and configured by the computer's manufacturer.
    '
    ' NOTE: The BIOS manufacturer's BIOS version is not usually in .NET version format. For
    ' example, a Lenovo systems returns a version like "LENOVO - 1510"
    '
    ' NOTE: It is generally preferable to use the Systems Management BIOS version number
    ' instead of the BIOS manufacturer version number.
    '
    ' The function takes one positional argument (strBIOSManufacturerVersion), which is
    ' populated upon success with a string containing the computer's systems BIOS version
    ' number in string format as reported by the BIOS manufacturer. The BIOS manufacturer's
    ' BIOS version number is equivalent to the Win32_BIOS object property Version
    '
    ' The function returns a 0 if the BIOS version string (as reported by the BIOS
    ' manufacturer) was obtained successfully. It returns a negative integer if an error
    ' occurred retrieving it. Finally, it returns a positive integer if the BIOS manufacturer
    ' BIOS version string was obtained, but multiple BIOS instances were present that contained
    ' data for the BIOS manufacturer version string. When this happens, only the first
    ' Win32_BIOS instance containing data for the BIOS version string (as reported by the BIOS
    ' manufacturer) is used.
    '
    ' Example:
    '   intReturnCode = GetBIOSManufacturerVersionString(strBIOSManufacturerVersion)
    '   If intReturnCode >= 0 Then
    '       ' The BIOS version string as reported by the BIOS manufacturer was retrieved
    '       ' successfully and is stored in strBIOSManufacturerVersion
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
    ' GetBIOSManufacturerVersionStringUsingBIOSInstances()
    'endregion DependsOn ####################################################

    Dim intFunctionReturn
    Dim arrBIOSInstances
    Dim strResult

    intFunctionReturn = 0

    intFunctionReturn = GetBIOSInstances(arrBIOSInstances)
    If intFunctionReturn >= 0 Then
        ' At least one Win32_BIOS instance was retrieved successfully
        intFunctionReturn = GetBIOSManufacturerVersionStringUsingBIOSInstances(strResult, arrBIOSInstances)
        If intFunctionReturn >= 0 Then
            ' The computer's BIOS version string was retrieved successfully and is stored in
            ' strResult
            strBIOSManufacturerVersion = strResult
        End If
    End If
    
    GetBIOSManufacturerVersionString = intFunctionReturn
End Function
