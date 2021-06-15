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
    ' Version: 1.1.20210614.0
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
    ' User "Shem Sargent" on Super User, who provided sample code for augmenting WMI-based
    ' version numbers with their revision number:
    ' https://superuser.com/a/1160428/334370
    'endregion Acknowledgements ####################################################

    'region DependsOn ####################################################
    ' TestObjectForData()
    ' ConnectLocalWMINamespace()
    'endregion DependsOn ####################################################

    Dim intFunctionReturn
    Dim intReturnMultiplier
    Dim objWMI
    Dim intReturnCode
    Dim colOperatingSystems
    Dim objOperatingSystem
    Dim strWorkingOSVersion
    Dim arrVersionNumber

    Err.Clear

    intFunctionReturn = 0
    intReturnMultiplier = 1

    intReturnCode = ConnectLocalWMINamespace(objWMI, Null, Null)
    If intReturnCode <> 0 Then
        intFunctionReturn = intFunctionReturn + (intReturnCode * intReturnMultiplier)
    Else
        intReturnMultiplier = intReturnMultiplier * 16
        On Error Resume Next
        Set colOperatingSystems = objWMI.ExecQuery("Select Version From Win32_OperatingSystem")
        If Err Then
            On Error Goto 0
            Err.Clear
            intFunctionReturn = intFunctionReturn + (-1 * intReturnMultiplier)
        Else
            For Each objOperatingSystem in colOperatingSystems
                strWorkingOSVersion = objOperatingSystem.Version
            Next
            If Err Then
                On Error Goto 0
                Err.Clear
                intFunctionReturn = intFunctionReturn + (-2 * intReturnMultiplier)
            Else
                On Error Goto 0
                If TestObjectForData(strWorkingOSVersion) = False Then
                    intFunctionReturn = intFunctionReturn + (-3 * intReturnMultiplier)
                End If
            End If
        End If
    End If

    If intFunctionReturn = 0 Then
        If boolExcludeRevisionNumber = True Then
            arrVersionNumber = Split(strWorkingOSVersion, ".")
            If UBound(arrVersionNumber) >= 3 Then
                ' Revision portion of version number is present
                strWorkingOSVersion = arrVersionNumber(0) & "." & arrVersionNumber(1) & "." & arrVersionNumber(2)
            End If
        End If
        strOperatingSystemVersion = strWorkingOSVersion
    End If
    
    GetWindowsOperatingSystemVersionNumberAsStringUsingWMI = intFunctionReturn
End Function
