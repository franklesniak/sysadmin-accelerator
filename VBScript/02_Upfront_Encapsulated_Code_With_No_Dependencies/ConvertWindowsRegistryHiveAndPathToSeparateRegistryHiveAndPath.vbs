Function ConvertWindowsRegistryHiveAndPathToSeparateRegistryHiveAndPath(ByRef strRegistryHiveName, ByRef strRegistryPathWithoutHive, ByVal strPathToRegKey)
    'region FunctionMetadata ####################################################
    ' Safely takes a string that contains a registry hive and path and converts it to a
    ' separated registry hive and registry path.
    '
    ' Function takes three positional arguments:
    '   The first argument (strRegistryHiveName) will be populated upon success with a
    '       standardized registry hive name from the following list:
    '           "HKEY_CURRENT_USER"
    '           "HKEY_LOCAL_MACHINE"
    '           "HKEY_DEFAULT_USER" - a "fake" registry hive that references the default user
    '               profile's HKCU registry hive. This designation was created by the function
    '               author to facilitate downstream processing, i.e., automatic mounting and
    '               dismounting of the default user profile's HKCU registry hive.
    '           "HKEY_CLASSES_ROOT" - a "fake" registry hive that represents a joining of
    '               HKCU\Software\Classes and HKLM\Software\Classes. Per Wikipedia, if a given
    '               value exists in both HKCU\Software\Classes and HKLM\Software\Classes, the
    '               one in HKCU\Software\Classes takes precedence.
    '           "HKEY_USERS"
    '           "HKEY_CURRENT_CONFIG" - a "fake" registry hive that serves as an alias for
    '               "HKLM\SYSTEM\CurrentControlSet\Hardware Profiles\Current".
    '           "HKEY_DYN_DATA" - only present in Windows 95, 98, and ME.
    '           "HKEY_PERFORMANCE_DATA" - a "fake" registry hive that exposes performance
    '               information; not persistent/not stored on disk.
    '   The second argument (strRegistryPathWithoutHive) will be populated upon success with
    '       just the path portion of the registry key (i.e., the full path specified in the
    '       third argument minus the registry hive prefix)
    '   The third argument (strPathToRegKey) provides the path to the registry key that is to
    '       be converted. It must be specified with one of the following prefixes, which
    '       specifies the key's registry hive:
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
    '       For example, a valid specifiication for this third argument would be:
    '       "HKLM\Software\Microsoft\Windows"
    '
    ' The function returns 0 if the full registry path was successfully converted into its
    ' respective hive/path parts. A negative number is returned if the registry path was
    ' invalid and could not be converted.
    '
    ' Example:
    '   intReturnCode = ConvertWindowsRegistryHiveAndPathToSeparateRegistryHiveAndPath(strHive, strPathOnly, "HKLM\Software\Microsoft\Windows")
    '   If intReturnCode = 0 Then
    '       ' Conversion completed successfully
    '       ' strHive contains "HKEY_LOCAL_MACHINE"
    '       ' strPathOnly contains "Software\Microsoft\Windows"
    '   End If
    '
    ' Version: 2.0.20210614.0
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
    'endregion Acknowledgements ####################################################

    'region DependsOn ####################################################
    ' TestObjectIsStringContainingData()
    ' TestObjectForData()
    'endregion DependsOn ####################################################

    Dim intFunctionReturn
    Dim arrRegistryPath
    Dim intVariableType
    Dim intUpperBound
    Dim strRegistryHiveStaging
    Dim strRegistryPathStaging
    Dim intCounter
    Dim intCounterB

    Const REG_PATH_SEPARATOR = "\"

    Err.Clear

    intFunctionReturn = 0

    If TestObjectIsStringContainingData(strPathToRegKey) = False Then
        intFunctionReturn = -1
    Else
        On Error Resume Next
        arrRegistryPath = Split(strPathToRegKey, REG_PATH_SEPARATOR)
        If Err Then
            On Error Goto 0
            Err.Clear
            intFunctionReturn = -2
        Else
            intUpperBound = UBound(arrRegistryPath)
            If Err Then
                On Error Goto 0
                Err.Clear
                intFunctionReturn = -3
            Else
                On Error Goto 0
            End If
        End If
    End If

    If intFunctionReturn = 0 Then
        ' No error occurred
        ' strPathToRegKey is a string and contains data
        ' arrRegistryPath is the split of strPathToRegKey on "\"
        ' intUpperBound is the upper index of arrRegistryPath
        If intUpperBound < 1 Then
            intFunctionReturn = -4
        Else
            Select Case UCase(arrRegistryPath(0))
                Case "HKCU"
                    strRegistryHiveStaging = "HKEY_CURRENT_USER"
                Case "HKEY_CURRENT_USER"
                    strRegistryHiveStaging = "HKEY_CURRENT_USER"
                Case "HKLM"
                    strRegistryHiveStaging = "HKEY_LOCAL_MACHINE"
                Case "HKEY_LOCAL_MACHINE"
                    strRegistryHiveStaging = "HKEY_LOCAL_MACHINE"
                Case "HKDU"
                    strRegistryHiveStaging = "HKEY_DEFAULT_USER"
                Case "HKEY_DEFAULT_USER"
                    strRegistryHiveStaging = "HKEY_DEFAULT_USER"
                Case "HKCR"
                    strRegistryHiveStaging = "HKEY_CLASSES_ROOT"
                Case "HKEY_CLASSES_ROOT"
                    strRegistryHiveStaging = "HKEY_CLASSES_ROOT"
                Case "HKU"
                    strRegistryHiveStaging = "HKEY_USERS"
                Case "HKEY_USERS"
                    strRegistryHiveStaging = "HKEY_USERS"
                Case "HKCC"
                    strRegistryHiveStaging = "HKEY_CURRENT_CONFIG"
                Case "HKEY_CURRENT_CONFIG"
                    strRegistryHiveStaging = "HKEY_CURRENT_CONFIG"
                Case "HKDD"
                    strRegistryHiveStaging = "HKEY_DYN_DATA"
                Case "HKEY_DYN_DATA"
                    strRegistryHiveStaging = "HKEY_DYN_DATA"
                Case "HKPD"
                    strRegistryHiveStaging = "HKEY_PERFORMANCE_DATA"
                Case "HKEY_PERFORMANCE_DATA"
                    strRegistryHiveStaging = "HKEY_PERFORMANCE_DATA"
                Case Else
                    intFunctionReturn = -5
            End Select
        End If
    End If

    If intFunctionReturn = 0 Then
        ' No error occurred
        ' strPathToRegKey is a string and contains data
        ' arrRegistryPath is the split of strPathToRegKey on "\"
        ' intUpperBound is the upper index of arrRegistryPath, and it is at least 1
        ' strRegistryHiveStaging contains a normalized registry hive name
        intCounter = 1
        While TestObjectForData(arrRegistryPath(intCounter)) = False And intCounter <= intUpperBound
            intCounter = intCounter + 1
        Wend
        If intCounter > intUpperBound Then
            intFunctionReturn = -6
        Else
            strRegistryPathStaging = arrRegistryPath(intCounter)
            For intCounterB = intCounter + 1 To intUpperBound
                If TestObjectForData(arrRegistryPath(intCounterB)) = True Then
                    strRegistryPathStaging = strRegistryPathStaging & REG_PATH_SEPARATOR & arrRegistryPath(intCounterB)
                End If
            Next
        End If
    End If

    If intFunctionReturn = 0 Then
        strRegistryHiveName = strRegistryHiveStaging
        strRegistryPathWithoutHive = strRegistryPathStaging
    End If

    ConvertWindowsRegistryHiveAndPathToSeparateRegistryHiveAndPath = intFunctionReturn
End Function
