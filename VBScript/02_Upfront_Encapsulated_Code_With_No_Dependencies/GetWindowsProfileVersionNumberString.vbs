Function GetWindowsProfileVersionNumberString(ByRef strProfileVersionNumber)
    'region FunctionMetadata ####################################################
    ' Gets the string version number of the Windows Profile for the current operating system
    '
    ' Function takes one positional argument (strProfileVersionNumber), which will be populated
    ' upon success with the version number of the Windows Profile for the current operating
    ' system.
    '
    ' The function returns 0 if the Windows profile version number was determined successfully.
    ' A negative number is returned if the Windows profile version number could not be
    ' determined.
    '
    ' Example:
    '   intReturnCode = GetWindowsProfileVersionNumberString(strProfileVersionNumber)
    '   If intReturnCode = 0 Then
    '       ' The profile version number was determined successfully
    '       ' strProfileVersionNumber contains a version number like "2.0"
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

    'region DependsOn ####################################################
    ' GetWindowsOperatingSystemVersionNumberAsString()
    ' GetWindowsProfileVersionNumberStringFromOperatingSystemVersionNumber()
    'endregion DependsOn ####################################################

    Dim intFunctionReturn
    Dim intReturnOffset
    Dim intReturnCode
    Dim strOperatingSystemVersion
    Dim strWorkingProfileVersionNumber

    intFunctionReturn = 0
    intReturnOffset = 0

    ' Get OS version number without using registry:
    intReturnCode = GetWindowsOperatingSystemVersionNumberAsString(strOperatingSystemVersion, 1, 1, Null, -1, Array(1, 2, 3, 4, 5, 6, 7))
    If intReturnCode < 0 Then
        intFunctionReturn = intReturnOffset + intReturnCode
    Else
        intReturnOffset = intReturnOffset - 33825
        intReturnCode = GetWindowsProfileVersionNumberStringFromOperatingSystemVersionNumber(strWorkingProfileVersionNumber, strOperatingSystemVersion)
        If intReturnCode < 0 Then
            intFunctionReturn = intReturnOffset + intReturnCode
        Else
            intReturnOffset = intReturnOffset - 24
        End If
    End If

    'Most-negative intFunctionReturn is -33849

    If intFunctionReturn = 0 Then
        strProfileVersionNumber = strWorkingProfileVersionNumber
    End If
    GetWindowsProfileVersionNumberString = intFunctionReturn
End Function
