Function GetWindowsProfileMajorVersionNumberFromOperatingSystemVersionNumber(ByRef lngProfileMajorVersionNumber, ByVal strOperatingSystemVersionNumber)
    'region FunctionMetadata ####################################################
    ' For a given operating system version number, gets the major version number of the
    ' Windows Profile.
    '
    ' Function takes two positional arguments:
    '   The first argument (lngProfileMajorVersionNumber) will be populated upon success with
    '       the integer major version number of the Windows Profile for the operating system
    '       version specified by the second argument (strOperatingSystemVersionNumber).
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
    '   intReturnCode = GetWindowsProfileMajorVersionNumberFromOperatingSystemVersionNumber(lngProfileMajorVersionNumber, "6.1.7601")
    '   If intReturnCode = 0 Then
    '       ' The profile version number was determined successfully
    '       ' lngProfileMajorVersionNumber equals 2
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
    ' Citrix, for providing documentation on the version numbers of Windows profiles given the
    ' operating system version:
    ' https://docs.citrix.com/en-us/profile-management/current-release/how-it-works/about-profiles.html
    'endregion Acknowledgements ####################################################

    'region DependsOn ####################################################
    ' GetWindowsProfileVersionNumberStringFromOperatingSystemVersionNumber()
    ' ConvertStringVersionNumberToMajorMinorBuildRevisionIntegers()
    'endregion DependsOn ####################################################

    Dim intFunctionReturn
    Dim intReturnOffset
    Dim strProfileVersionNumber
    Dim intReturnCode
    Dim lngMajor
    Dim lngMinor
    Dim lngBuild
    Dim lngRevision

    intFunctionReturn = 0
    intReturnOffset = 0

    intReturnCode = GetWindowsProfileVersionNumberStringFromOperatingSystemVersionNumber(strProfileVersionNumber, strOperatingSystemVersionNumber)
    If intReturnCode < 0 Then
        intFunctionReturn = intReturnCode + intReturnOffset
    Else
        intReturnOffset = intReturnOffset - 24
        intReturnCode = ConvertStringVersionNumberToMajorMinorBuildRevisionIntegers(lngMajor, lngMinor, lngBuild, lngRevision, strProfileVersionNumber)
        If intReturnCode < 0 Then
            intFunctionReturn = intReturnCode + intReturnOffset
        Else
            intReturnOffset = intReturnOffset - 17
            If lngMajor < 1 Then
                intFunctionReturn = -1 + intReturnOffset
            Else
                intReturnOffset = intReturnOffset - 1
            End If
        End If
    End If

    ' Most negative possible intFunctionReturn is -42

    If intFunctionReturn = 0 Then
        lngProfileMajorVersionNumber = lngMajor
    End If
    GetWindowsProfileMajorVersionNumberFromOperatingSystemVersionNumber = intFunctionReturn
End Function
