Function GetWindowsProfileVersionNumberStringFromOperatingSystemVersionNumber(ByRef strProfileVersionNumber, ByVal strOperatingSystemVersionNumber)
    'region FunctionMetadata ####################################################
    ' For a given operating system version number, gets the string version number of the
    ' Windows Profile.
    '
    ' Function takes two positional arguments:
    '   The first argument (strProfileVersionNumber) will be populated upon success with the
    '       version number of the Windows Profile for the operating system version specified by
    '       the second argument (strOperatingSystemVersionNumber).
    '   The second argument (strOperatingSystemVersionNumber) a string operating system version
    '       number in the format major.minor.build.revision. If the operating system version
    '       number is Windows 10, Windows Server 2016, or newer, than major.minor.build are
    '       minimally required. If the operating system version number is older than Windows 10
    '       and Windows Server 2016, then major.minor are minimally required.
    '
    ' The function returns 0 if the Windows profile version number was determined successfully.
    ' A negative number is returned if the Windows profile version number could not be
    ' determined.
    '
    ' Example:
    '   intReturnCode = GetWindowsProfileVersionNumberStringFromOperatingSystemVersionNumber(strProfileVersionNumber, "6.1.7601")
    '   If intReturnCode = 0 Then
    '       ' The profile version number was determined successfully
    '       ' strProfileVersionNumber contains "2.0"
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
    ' at https://github.com/franklesniak/VBScript_Resources
    'endregion DownloadLocationNotice ####################################################

    'region Acknowledgements ####################################################
    ' Citrix, for providing documentation on the version numbers of Windows profiles given the
    ' operating system version:
    ' https://docs.citrix.com/en-us/profile-management/current-release/how-it-works/about-profiles.html
    'endregion Acknowledgements ####################################################

    'region DependsOn ####################################################
    ' TestObjectIsStringContainingData()
    ' ConvertStringVersionNumberToMajorMinorBuildRevisionIntegers()
    'endregion DependsOn ####################################################

    Dim intFunctionReturn
    Dim intReturnOffset
    Dim boolTest
    Dim strWorkingProfileVersionNumber
    Dim intReturnCode
    Dim lngMajor
    Dim lngMinor
    Dim lngBuild
    Dim lngRevision

    Err.Clear

    intFunctionReturn = 0
    intReturnOffset = 0

    If TestObjectIsStringContainingData(strOperatingSystemVersionNumber) = False Then
        intFunctionReturn = -1
    Else
        intReturnOffset = intReturnOffset - 1
        intReturnCode = ConvertStringVersionNumberToMajorMinorBuildRevisionIntegers(lngMajor, lngMinor, lngBuild, lngRevision, strOperatingSystemVersionNumber)
        If intReturnCode < 0 Then
            intFunctionReturn = intReturnCode + intReturnOffset
        Else
            intReturnOffset = intReturnOffset - 17
        End If
    End If

    If intFunctionReturn = 0 Then
        ' No error occurred
        ' strOperatingSystemVersionNumber converted to lngMajor, lngMinor, lngBuild, lngRevision
        If lngMajor >= 10 Then
            If ((lngMajor > 10) Or (lngMajor = 10 And lngMinor > 0) Then
                ' Some release newer than Windows 10 or Windows Server v10
                ' TODO: Update when OS is released newer than Windows 10
                strWorkingProfileVersionNumber = "6.0"
                intReturnOffset = intReturnOffset - 3
            Else
                ' OS is 10.0.x
                If lngBuild < 0 Then
                    ' Build was not supplied. Error!
                    intFunctionReturn = -1 + intReturnOffset
                    intReturnOffset = intReturnOffset - 2
                Else
                    intReturnOffset = intReturnOffset - 3
                    If lngBuild >= 14393 Then
                        ' Windows 10 1607 or Windows Server 2016 or newer
                        strWorkingProfileVersionNumber = "6.0"
                    Else
                        ' Windows 10 1507 or Windows 10 1511
                        strWorkingProfileVersionNumber = "5.0"
                    End If
                End If
            End If
        Else
            ' OS version older than Windows 10
            intReturnOffset = intReturnOffset - 1
            If lngMajor = 6 Then
                'OS verson 6.x
                If lngMinor = 0 Or lngMinor = 1 Then
                    ' Windows Vista, Windows Server 2008, Windows 7, or Windows Server 2008 R2
                    strWorkingProfileVersionNumber = "2.0"
                    intReturnOffset = intReturnOffset - 2
                ElseIf lngMinor = 2 Then
                    'Windows 8 or Windows Server 2012
                    strWorkingProfileVersionNumber = "3.0"
                    intReturnOffset = intReturnOffset - 2
                ElseIf lngMinor = 3 Then
                    strWorkingProfileVersionNumber = "4.0"
                    intReturnOffset = intReturnOffset - 2
                ElseIf lngMinor = 4 Then
                    ' Early beta of Windows 10
                    strWorkingProfileVersionNumber = "5.0"
                    intReturnOffset = intReturnOffset - 2
                Else
                    intFunctionReturn = -1 + intReturnOffset
                    intReturnOffset = intReturnOffset - 1
                End If
            ElseIf lngMajor = 5 Then
                ' OS is 5.x
                intReturnOffset = intReturnOffset - 3
                strWorkingProfileVersionNumber = "1.0"
            Else
                ' OS is 4.x
                intReturnOffset = intReturnOffset - 2
                If lngBuild < 0 Then
                    ' Build was not supplied. Error!
                    intFunctionReturn = -1 + intReturnOffset
                Else
                    intReturnOffset = intReturnOffset - 1
                    If lngBuild = 1381 Then
                        ' Windows NT 4.0
                        strWorkingProfileVersionNumber = "1.0"
                    Else
                        ' Windows 95, Windows 98, or Windows ME
                        ' Technically, these operating systems supported roaming profiles, so
                        ' let's assign a version number:
                        strWorkingProfileVersionNumber = "0.1"
                    End If
                End If
            End If
        End If
    End If

    'Most-negative intFunctionReturn is -24

    If intFunctionReturn = 0 Then
        strProfileVersionNumber = strWorkingProfileVersionNumber
    End If
    GetWindowsProfileVersionNumberStringFromOperatingSystemVersionNumber = intFunctionReturn
End Function
