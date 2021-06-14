Function TestOperatingSystemVersionRestrictsScriptProcessToX86Architecture(ByRef boolOnlyX86Architecture, ByVal lngMajor, ByVal lngMinor, ByVal lngBuild)
    'region FunctionMetadata ####################################################
    ' Returns true if the supplied operating system version was only written for IA32 (Intel
    ' 32-bit x86). For example, Windows 95 was only written for IA32; if lngMajor was 4,
    ' lngMinor was 0, and lngBuild was 950, this function would set boolOnlyX86Architecture to
    ' True
    '
    ' This function takes four arguments:
    '   - The first argument (boolOnlyX86Architecture) is populated upon success to either True
    '       or False, to reflect whether the supplied operating system version was only
    '       available in Intel IA32 (32-bit x86) processor architecture (True), or if it was
    '       available for other architectures (False).
    '   - The second argument (lngMajor) is set to an integer, reflecting the major portion of
    '       the version number to be evaluated.
    '   - The third argument (lngMinor) is set to an integer, reflecting the minor portion of
    '       the version number to be evaluated.
    '   - The fourth argument (lngBuild) is set to an integer, reflecting the build portion of
    '       the version number to be evaluated.
    '
    ' The function returns 0 if the script was able to evaluate whether or not the supplied
    ' operating system version was only available for IA32 (Intel 32-bit x86) processor
    ' architectures.
    '
    ' Example:
    '   intReturnCode = TestOperatingSystemVersionRestrictsScriptProcessToX86Architecture(boolOnlyX86Architecture, 4, 0, 0)
    '   If intReturnCode = 0 Then
    '       ' Evaluation was successful
    '       ' boolOnlyX86Architecture is True (Windows 95)
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
    ' BetaWiki, for documenting all the different versions of Windows NT 4.0, which is the only
    ' Microsoft operating system in the 4.x version range that supported non Intel IA-32
    ' (32-bit x86) operating systems.
    ' https://betawiki.net/wiki/Windows_NT_4.0
    'endregion Acknowledgements ####################################################

    'region DependsOn ####################################################
    ' None!
    'endregion DependsOn ####################################################

    Dim intFunctionReturn
    Dim boolTest
    Dim boolTempEvaluation

    Err.Clear

    intFunctionReturn = 0
    
    On Error Resume Next
    boolTest = (lngMajor = 4)
    If Err Then
        On Error Goto 0
        Err.Clear
        intFunctionReturn = -1
    Else
        On Error Goto 0
        If boolTest = False Then
            boolTempEvaluation = False
        Else
            ' 4.x
            boolTempEvaluation = True
            
            ' Check to see if it's NT4, which would change to False
            On Error Resume Next
            boolTest = (lngMinor = 0)
            If Err Then
                On Error Goto 0
                Err.Clear
                intFunctionReturn = -2
            Else
                On Error Goto 0
                If boolTest = True Then
                    On Error Resume Next
                    Select Case lngBuild
                        Case 1381
                            ' NT 4.0 RTM - SP6a + Security Rollup
                            boolTempEvaluation = False
                        Case 1096
                            ' NT 4.0 Beta 1
                            boolTempEvaluation = False
                        Case 1116
                            ' NT 4.0 Beta 1
                            boolTempEvaluation = False
                        Case 1124
                            ' NT 4.0 Beta 1
                            boolTempEvaluation = False
                        Case 1130
                            ' NT 4.0 Beta 1
                            boolTempEvaluation = False
                        Case 1141
                            ' NT 4.0 Beta 1
                            boolTempEvaluation = False
                        Case 1150
                            ' NT 4.0 Beta 1
                            boolTempEvaluation = False
                        Case 1166
                            ' NT 4.0 Beta 1
                            boolTempEvaluation = False
                        Case 1175
                            ' NT 4.0 Beta 1
                            boolTempEvaluation = False
                        Case 1227
                            ' NT 4.0 Beta 1
                            boolTempEvaluation = False
                        Case 1234
                            ' NT 4.0 Beta 1
                            boolTempEvaluation = False
                        Case 1249
                            ' NT 4.0 Beta 2
                            boolTempEvaluation = False
                        Case 1261
                            ' NT 4.0 Beta 2
                            boolTempEvaluation = False
                        Case 1264
                            ' NT 4.0 Beta 2
                            boolTempEvaluation = False
                        Case 1273
                            ' NT 4.0 Beta 2
                            boolTempEvaluation = False
                        Case 1287
                            ' NT 4.0 Beta 2
                            boolTempEvaluation = False
                        Case 1293
                            ' NT 4.0 Beta 2
                            boolTempEvaluation = False
                        Case 1297
                            ' NT 4.0 Beta 2
                            boolTempEvaluation = False
                        Case 1314
                            ' NT 4.0 Beta 2
                            boolTempEvaluation = False
                        Case 1326
                            ' NT 4.0 Release Candidate 1
                            boolTempEvaluation = False
                        Case 1327
                            ' NT 4.0 Release Candidate 1
                            boolTempEvaluation = False
                        Case 1332
                            ' NT 4.0 Release Candidate 1
                            boolTempEvaluation = False
                        Case 1345
                            ' NT 4.0 Release Candidate 1
                            boolTempEvaluation = False
                        Case 1353
                            ' NT 4.0 Release Candidate 2
                            boolTempEvaluation = False
                        Case 1369
                            ' NT 4.0 Release Candidate 2
                            boolTempEvaluation = False
                    End Select
                    If Err Then
                        On Error Goto 0
                        Err.Clear
                        intFunctionReturn = -3
                    Else
                        On Error Goto 0
                    End If
                End If
            End If
        End If
    End If

    If intFunctionReturn = 0 Then
        boolOnlyX86Architecture = boolTempEvaluation
    End If

    TestOperatingSystemVersionRestrictsScriptProcessToX86Architecture = intFunctionReturn
End Function
