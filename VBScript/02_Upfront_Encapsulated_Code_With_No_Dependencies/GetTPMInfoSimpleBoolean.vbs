Function GetTPMInfoSimpleBoolean(ByRef boolTPMPresent, ByRef boolEnabledTPMPresent, ByRef boolActivatedTPMPresent, ByRef boolReadyTPMPresent)
    'region FunctionMetadata #######################################################
    ' This function retrieves the most commonly needed TPM information, namely:
    '  - Whether a TPM was present
    '  - If a TPM was present, and if this script was run with elevated privileges:
    '     - Whether at least one TPM was enabled
    '     - Whether at least one TPM was activated
    '     - Whether at least one TPM was ready
    '
    ' **NOTE**: If this function is called in a script run without elevated privileges,
    '           it will return potentially misleading results (see definitions for
    '           "False" results, below). To cover the non-admin rights scenario better,
    '           use GetTPMInfo() instead
    '
    ' The function takes four positional arguments:
    '  - The first argument (boolTPMPresent) is boolean value that will be populated
    '    with either:
    '         True = a TPM was present
    '         False = either a TPM was not present, or it could not be determined
    '             whether a TPM was present. The latter is usually caused by
    '             restrictions that prevent access to WMI
    '  - The second argument (boolEnabledTPMPresent) is a boolean value that will be
    '    populated with either:
    '         True = at least one enabled TPM was present
    '         False = either no enabled TPM was present, or it could not be determined
    '             whether an enabled TPM was present. The latter is usually caused by
    '             the script being run without elevated permissions
    '  - The third argument (boolActivatedTPMPresent) is a boolean value that will be
    '    populated with either:
    '         True = at least one activated TPM was present
    '         False = either no activated TPM was present, or it could not be
    '             determined whether an activated TPM was present. The latter is
    '             usually caused by the script being run without elevated permissions
    '  - The fourth argument (boolReadyTPMPresent) is a boolean value that will be
    '    populated with either:
    '         True = at least one ready TPM was present
    '         False = either no ready TPM was present, or it could not be determined
    '             whether a ready TPM was present. The latter is usually caused by the
    '             script being run without elevated permissions
    '
    ' The function returns a 0 if all TPM information was retrieved successfully. It
    ' returns a positive integer if it was determined whether a TPM was present, but
    ' the TPM(s)' enabled status, activation status, or readiness could not be
    ' determined. In other words, if the function returned 0, then a False result for
    ' the second, third, or fourth argument should be regarded as no TPM present.
    ' However, if the function returned a positive integer, then a False result in the
    ' second, third, or fourth argument should be regarded as "the result could not be
    ' determined". The function returns a negative integer if an error occurred and no
    ' information could be determined.
    '
    ' Example:
    '   intReturnCode = GetTPMInfoSimpleBoolean(boolTPMPresent, boolEnabledTPMPresent, boolActivatedTPMPresent, boolReadyTPMPresent)
    '   If intReturnCode >= 0 Then
    '       ' boolTPMPresent is populated as expected
    '   End If
    '   If intReturnCode = 0 Then
    '       ' boolEnabledTPMPresent, boolActivatedTPMPresent, and
    '       ' boolReadyTPMPresent are also populated as expected
    '   End If
    '
    ' Version: 1.0.20210617.1
    'endregion FunctionMetadata #######################################################

    'region License ################################################################
    ' Copyright 2021 Frank Lesniak
    '
    ' Permission is hereby granted, free of charge, to any person obtaining a copy of
    ' this software and associated documentation files (the "Software"), to deal in the
    ' Software without restriction, including without limitation the rights to use,
    ' copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the
    ' Software, and to permit persons to whom the Software is furnished to do so,
    ' subject to the following conditions:
    '
    ' The above copyright notice and this permission notice shall be included in all
    ' copies or substantial portions of the Software.
    '
    ' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    ' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
    ' FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR
    ' COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN
    ' AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
    ' WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
    'endregion License ################################################################

    'region DownloadLocationNotice #################################################
    ' The most up-to-date version of this script can be found on the author's GitHub
    ' repository at https://github.com/franklesniak/sysadmin-accelerator
    'endregion DownloadLocationNotice #################################################

    'region Acknowledgements #######################################################
    ' Microsoft and Sven Aelterman, for providing ZTICheckForTPM/ZTICheckForTPM_v2,
    ' which inspired this approach.
    'endregion Acknowledgements #######################################################

    'region DependsOn ##############################################################
    ' GetTPMInfo()
    'endregion DependsOn ##############################################################

    Dim intFunctionReturn
    Dim intPreliminaryTriStateTPMPresent
    Dim intPreliminaryTriStateEnabledTPMPresent
    Dim intPreliminaryTriStateActivatedTPMPresent
    Dim intPreliminaryTriStateReadyTPMPresent

    intFunctionReturn = GetTPMInfo(intPreliminaryTriStateTPMPresent, intPreliminaryTriStateEnabledTPMPresent, intPreliminaryTriStateActivatedTPMPresent, intPreliminaryTriStateReadyTPMPresent)

    If intPreliminaryTriStateTPMPresent = -1 Then
        boolTPMPresent = True
    ElseIf intPreliminaryTriStateTPMPresent = 0 Then
        boolTPMPresent = False
    ElseIf intPreliminaryTriStateTPMPresent = 1 Then
        ' Make sure this gets flagged as an error state if not already:
        If intFunctionReturn >= 0 Then intFunctionReturn = -1
    End If

    If intFunctionReturn >= 0 Then
        If intPreliminaryTriStateEnabledTPMPresent = -1 Then
            boolEnabledTPMPresent = True
        ElseIf intPreliminaryTriStateEnabledTPMPresent = 0 Then
            boolEnabledTPMPresent = False
        ElseIf intPreliminaryTriStateEnabledTPMPresent = 1 Then
            If intFunctionReturn <= 0 Then intFunctionReturn = 1
            boolEnabledTPMPresent = False
        End If

        If intPreliminaryTriStateActivatedTPMPresent = -1 Then
            boolActivatedTPMPresent = True
        ElseIf intPreliminaryTriStateActivatedTPMPresent = 0 Then
            boolActivatedTPMPresent = False
        ElseIf intPreliminaryTriStateActivatedTPMPresent = 1 Then
            If intFunctionReturn <= 0 Then intFunctionReturn = 1
            boolActivatedTPMPresent = False
        End If

        If intPreliminaryTriStateReadyTPMPresent = -1 Then
            boolReadyTPMPresent = True
        ElseIf intPreliminaryTriStateReadyTPMPresent = 0 Then
            boolReadyTPMPresent = False
        ElseIf intPreliminaryTriStateReadyTPMPresent = 1 Then
            If intFunctionReturn <= 0 Then intFunctionReturn = 1
            boolReadyTPMPresent = False
        End If
    End If

    GetTPMInfoSimpleBoolean = intFunctionReturn
End Function
