Function GetWindowsProfileMajorVersionNumber(ByRef lngProfileMajorVersionNumber)
    'region FunctionMetadata ####################################################
    ' Gets the major version number of the Windows Profile for the current operating system.
    '
    ' Function takes one positional argument (lngProfileMajorVersionNumber), which will be
    ' populated upon success with the integer major version number of the Windows Profile for
    ' the current operating system.
    '
    ' The function returns 0 if the Windows profile version number was determined successfully.
    ' A negative number is returned if the Windows profile version number could not be
    ' determined.
    '
    ' Example:
    '   intReturnCode = GetWindowsProfileMajorVersionNumber(lngProfileMajorVersionNumber)
    '   If intReturnCode = 0 Then
    '       ' The profile version number was determined successfully
    '       ' lngProfileMajorVersionNumber is set to an integer, e.g., 2
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
    ' GetWindowsProfileVersionNumberString()
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

    intReturnCode = GetWindowsProfileVersionNumberString(strProfileVersionNumber)
    If intReturnCode < 0 Then
        intFunctionReturn = intReturnCode + intReturnOffset
    Else
        intReturnOffset = intReturnOffset - 33849
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

    ' Most negative possible intFunctionReturn is -33867

    If intFunctionReturn = 0 Then
        lngProfileMajorVersionNumber = lngMajor
    End If
    GetWindowsProfileMajorVersionNumber = intFunctionReturn
End Function
