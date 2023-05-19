Function GetWindowsOperatingSystemVersionNumberAsStringUsingWMI(ByRef strOperatingSystemVersion, ByVal boolExcludeRevisionNumber)
    'region FunctionMetadata ####################################################
    ' Safely obtains the operating system version number from Win32_OperatingSystem using WMI
    '
    ' Function takes two positional arguments:
    '   The first argument (strOperatingSystemVersion) will be populated with the operating
    '       system version in string format upon success
    '   The second argument (boolExcludeRevisionNumber) indicates whether the function should
    '       remove the "revision" portion of the operating system version (i.e.,
    '       major.minor.build.revision) from the operating system version string that the
    '       function returns via the first argument. Doing so can be useful because, at the
    '       time of writing, WMI does not accurately retrieve the revision number, which can be
    '       misleading or cause confusion. The valid values for boolExcludeRevisionNumber are:
    '           True = return just the major.minor.build portions of the operating system
    '               version; if WMI provides a revision number, it is removed.
    '           False = return the version number exactly as it was provided by WMI
    '           Null = same as False
    '
    ' The function returns 0 or a positive number if the operating system version number was
    ' retrieved successfully. A negative number is returned if the operating system version
    ' number was not retrieved successfully.
    '
    ' Example:
    '   intReturnCode = GetWindowsOperatingSystemVersionNumberAsStringUsingWMI(strOperatingSystemVersion, True)
    '   If intReturnCode = 0 Then
    '       ' strOperatingSystemVersion is populated with the operating system version number
    '       ' in string format. If applicable, the revision portion of the version number was
    '       ' trimmed.
    '   Else
    '       ' The operating system version number could not be retrieved via WMI. Usually this
    '       ' occurs because something in the operating system blocked the creation of the WMI
    '       ' object, something is wrong with the WMI database, or in the case of Windows 95,
    '       ' Windows 98, or Windows NT 4.0, it is likely that WMI is not installed.
    '   End If
    '
    ' Version: 1.2.20230518.0
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
    ' User "Shem Sargent" on Super User, who provided sample code for augmenting WMI-based
    ' version numbers with their revision number:
    ' https://superuser.com/a/1160428/334370
    'endregion Acknowledgements ####################################################

    'region DependsOn ####################################################
    ' TestObjectForData()
    ' GetOperatingSystemInstances()
    ' TestObjectIsStringContainingData()
    'endregion DependsOn ####################################################

    Dim intFunctionReturn
    Dim intReturnMultiplier
    Dim objWMI
    Dim intReturnCode
    Dim colOperatingSystems
    Dim objOperatingSystem
    Dim strWorkingOSVersion
    Dim strOldWorkingOSVersion
    Dim strOSVersionToReturn
    Dim arrVersionNumber
    Dim intCountOfOperatingSystems

    Err.Clear

    intFunctionReturn = 0
    intReturnMultiplier = 1
    strWorkingOSVersion = ""
    strOSVersionToReturn = ""
    intCountOfOperatingSystems = 0

    intReturnCode = GetOperatingSystemInstances(colOperatingSystems)
    If intReturnCode < 0 Then
        intFunctionReturn = intFunctionReturn + (intReturnCode * intReturnMultiplier)
    Else
        On Error Resume Next
        For Each objOperatingSystem in colOperatingSystems
            If Err Then
                Err.Clear
            Else
                strOldWorkingOSVersion = strWorkingOSVersion
                strWorkingOSVersion = objOperatingSystem.Version
                If Err Then
                    Err.Clear
                    strWorkingOSVersion = strOldWorkingOSVersion
                Else
                    If TestObjectIsStringContainingData(strWorkingOSVersion) <> True Then
                        strWorkingOSVersion = strOldWorkingOSVersion
                    Else
                        ' Found a result with real OS version data
                        If TestObjectForData(strOSVersionToReturn) = False Then
                            strOSVersionToReturn = strWorkingOSVersion
                        End If
                        intCountOfOperatingSystems = intCountOfOperatingSystems + 1
                    End If
                End If
            End If
        Next
        On Error Goto 0
        If Err Then
            Err.Clear
        End If
    End If

    If intFunctionReturn = 0 And TestObjectIsStringContainingData(strOSVersionToReturn) Then
        If boolExcludeRevisionNumber = True Then
            arrVersionNumber = Split(strOSVersionToReturn, ".")
            If UBound(arrVersionNumber) >= 3 Then
                ' Revision portion of version number is present
                strOSVersionToReturn = arrVersionNumber(0) & "." & arrVersionNumber(1) & "." & arrVersionNumber(2)
            End If
        End If
        strOperatingSystemVersion = strOSVersionToReturn
    End If
    
    GetWindowsOperatingSystemVersionNumberAsStringUsingWMI = intFunctionReturn
End Function
