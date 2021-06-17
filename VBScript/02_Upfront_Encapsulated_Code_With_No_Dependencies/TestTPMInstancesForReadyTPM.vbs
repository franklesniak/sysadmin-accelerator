Function TestTPMInstancesForReadyTPM(ByRef boolReadyTPMPresent)
    'region FunctionMetadata ####################################################
    ' This function tests each TPM on the system to see if there is at least one ready TPM
    ' present.
    '
    ' The function takes one positional argument (boolReadyTPMPresent), which is populated
    ' upon success with a boolean True or False. True means that at least one ready TPM was
    ' present, while False means that no ready TPM was present.
    '
    ' The function returns a 0 if the test was performed successfully. It returns a negative
    ' integer if an error occurred performing the test.
    '
    ' Example:
    '   intReturnCode = TestTPMInstancesForReadyTPM(boolReadyTPMPresent)
    '   If intReturnCode = 0 Then
    '       ' The test was performed successfully
    '       ' boolReadyTPMPresent is True if at least one TPM was present that is ready
    '       ' boolReadyTPMPresent is False if no TPMs were found ready
    '   End If
    '
    ' Version: 1.0.20210615.0
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
    ' Microsoft and Sven Aelterman, for providing ZTICheckForTPM/ZTICheckForTPM_v2, which
    ' inspired this approach.
    'endregion Acknowledgements ####################################################

    'region DependsOn ####################################################
    ' GetTPMInstances()
    ' TestTPMInstancesForReadyTPMUsingTPMInstances()
    'endregion DependsOn ####################################################

    Dim intFunctionReturn
    Dim intReturnMultiplier
    Dim intReturnCode
    Dim arrTPMInstances
    Dim boolResult

    Err.Clear

    intFunctionReturn = 0
    intReturnMultiplier = 1

    intReturnCode = GetTPMInstances(arrTPMInstances)
    If intReturnCode < 0 Then
        intFunctionReturn = intFunctionReturn + (intReturnCode * intReturnMultiplier)
    Else
        ' No error occurred retrieving TPMs
        intReturnCode = TestTPMInstancesForReadyTPMUsingTPMInstances(boolResult, arrTPMInstances)
        If intReturnCode < 0 Then
            intFunctionReturn = intFunctionReturn + (intReturnCode * intReturnMultiplier)
        End if
    End If

    If intFunctionReturn = 0 Then
        boolReadyTPMPresent = boolResult
    End If
    
    TestTPMInstancesForReadyTPM = intFunctionReturn
End Function
