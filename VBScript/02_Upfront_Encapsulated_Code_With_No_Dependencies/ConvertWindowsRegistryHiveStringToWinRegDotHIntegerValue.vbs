Function ConvertWindowsRegistryHiveStringToWinRegDotHIntegerValue(ByRef lngRegistryHive, ByVal strRegistryHiveName)
    'region FunctionMetadata ####################################################
    ' Safely takes a string that contains a registry hive and converts it to a numerical value
    ' compatible with Windows subsystems that require the integer values specified in WinReg.h
    '
    ' Function takes two positional arguments:
    '   The first argument (lngRegistryHive) is set to a 32-bit integer upon success:
    '           "HKCU" Or "HKEY_CURRENT_USER" specified in the second argument
    '               (strRegistryHiveName): this argument (lngRegistryHive) will be set to
    '               &H80000001 (hex) = 2147483649
    '           "HKLM" Or "HKEY_LOCAL_MACHINE" specified in the second argument
    '               (strRegistryHiveName): this argument (lngRegistryHive) will be set to
    '               &H80000002 (hex) = 2147483650
    '           "HKDU" Or "HKEY_DEFAULT_USER" specified in the second argument
    '               (strRegistryHiveName): this argument (lngRegistryHive) will be set to &H4D2
    '               (hex) = 1234
    '               NOTE: This is a fake registry hive designation created by the function
    '               author to handle automatic mounting and unmounting of the default user
    '               profile's HKCU registry hive. This value should not be passed to Windows
    '               system calls that use WinReg.h values as it will result in an error.
    '               NOTE 2: If "HKDU" Or "HKEY_DEFAULT_USER" was specified in the second
    '               argument (strRegistryHiveName), the function will return 1
    '           "HKCR" Or "HKEY_CLASSES_ROOT" specified in the second argument
    '               (strRegistryHiveName): this argument (lngRegistryHive) will be set to
    '               &H80000000 (hex) = 2147483648
    '           "HKU" Or "HKEY_USERS" specified in the second argument (strRegistryHiveName):
    '               this argument (lngRegistryHive) will be set to &H80000003 (hex) =
    '               2147483651
    '           "HKCC" Or "HKEY_CURRENT_CONFIG" specified in the second argument
    '               (strRegistryHiveName): this argument (lngRegistryHive) will be set to
    '               &H80000005 (hex) = 2147483653
    '           "HKDD" Or "HKEY_DYN_DATA" specified in the second argument
    '               (strRegistryHiveName): this argument (lngRegistryHive) will be set to
    '               &H80000006 (hex) = 2147483654
    '           "HKPD" Or "HKEY_PERFORMANCE_DATA" specified in the second argument
    '               (strRegistryHiveName): this argument (lngRegistryHive) will be set to
    '               &H80000004 (hex) = 2147483652
    '   The second argument (strRegistryHiveName) is a string containing one of the following
    '       values:
    '           "HKCU" or "HKEY_CURRENT_USER"
    '           "HKLM" or "HKEY_LOCAL_MACHINE"
    '           "HKDU" or "HKEY_DEFAULT_USER" - a "fake" registry hive that references the
    '               default user profile's HKCU registry hive. This designation was created by
    '               the function author to facilitate downstream processing, i.e., automatic
    '               mounting and dismounting of the default user profile's HKCU registry hive.
    '           "HKCR" or "HKEY_CLASSES_ROOT" - a "fake" registry hive that represents a
    '               joining of HKCU\Software\Classes and HKLM\Software\Classes. Per Wikipedia,
    '               if a given value exists in both HKCU\Software\Classes and
    '               HKLM\Software\Classes, the one in HKCU\Software\Classes takes precedence.
    '           "HKU" or "HKEY_USERS"
    '           "HKCC" or "HKEY_CURRENT_CONFIG" - a "fake" registry hive that serves as an
    '               alias for "HKLM\SYSTEM\CurrentControlSet\Hardware Profiles\Current".
    '           "HKDD" or "HKEY_DYN_DATA" - only present in Windows 95, 98, and ME.
    '           "HKPD" or "HKEY_PERFORMANCE_DATA" - a "fake" registry hive that exposes
    '               performance information; not persistent/not stored on disk.
    '
    ' The function returns 0 if the registry hive was successfully converted to the equivalent
    ' integer value specified in WinReg.h. The function returns 1 if the registry hive
    ' specified was the fake "HKDU" / "HKEY_DEFAULT_USER" hive created by the function author
    ' to facilitate downstream processing and automatic mounting/unmounting of the default user
    ' profile's HKCU registry hive. A negative number is returned if the registry hive name was
    ' invalid and could not be converted.
    '
    ' Example:
    '   intReturnCode = ConvertWindowsRegistryHiveStringToWinRegDotHIntegerValue(lngRegistryHive, "HKEY_LOCAL_MACHINE")
    '   If intReturnCode >= 0 Then
    '       ' Conversion completed successfully
    '       ' lngRegistryHive equals 2147483650
    '   End If
    '
    ' Version: 1.0.20210614.0
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
    ' Microsoft, who published the list of Windows Registry hives present in WinReg.h on the
    ' following page:
    ' https://docs.microsoft.com/en-us/previous-versions/windows/desktop/regprov/enumkey-method-in-class-stdregprov
    '
    ' Stack Overflow user "TheMadTechnician", who listed additional values from WinReg.h:
    ' https://stackoverflow.com/a/24892338/2134110
    'endregion Acknowledgements ####################################################

    'region DependsOn ####################################################
    ' TestObjectIsStringContainingData()
    'endregion DependsOn ####################################################

    Dim intFunctionReturn
    Dim lngRegistryHiveStaging

    Err.Clear

    intFunctionReturn = 0

    If TestObjectIsStringContainingData() = False Then
        intFunctionReturn = -1
    Else
        ' strRegistryHiveName is a string and contains data
        Select Case UCase(strRegistryHiveName)
            Case "HKCU"
                lngRegistryHiveStaging = &H80000001
            Case "HKEY_CURRENT_USER"
                lngRegistryHiveStaging = &H80000001
            Case "HKLM"
                lngRegistryHiveStaging = &H80000002
            Case "HKEY_LOCAL_MACHINE"
                lngRegistryHiveStaging = &H80000002
            Case "HKDU"
                lngRegistryHiveStaging = 1234
                intFunctionReturn = 1
            Case "HKEY_DEFAULT_USER"
                lngRegistryHiveStaging = 1234
                intFunctionReturn = 1
            Case "HKCR"
                lngRegistryHiveStaging = &H80000000
            Case "HKEY_CLASSES_ROOT"
                lngRegistryHiveStaging = &H80000000
            Case "HKU"
                lngRegistryHiveStaging = &H80000003
            Case "HKEY_USERS"
                lngRegistryHiveStaging = &H80000003
            Case "HKCC"
                lngRegistryHiveStaging = &H80000005
            Case "HKEY_CURRENT_CONFIG"
                lngRegistryHiveStaging = &H80000005
            Case "HKDD"
                lngRegistryHiveStaging = &H80000006
            Case "HKEY_DYN_DATA"
                lngRegistryHiveStaging = &H80000006
            Case "HKPD"
                lngRegistryHiveStaging = &H80000004
            Case "HKEY_PERFORMANCE_DATA"
                lngRegistryHiveStaging = &H80000004
            Case Else
                intFunctionReturn = -2
        End Select
    End If

    If intFunctionReturn >= 0 Then
        lngRegistryHive = lngRegistryHiveStaging
    End If

    ConvertWindowsRegistryHiveStringToWinRegDotHIntegerValue = intFunctionReturn
End Function
