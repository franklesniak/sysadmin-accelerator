Function GetSerialNumberUsingBIOSInstances(ByRef strSerialNumber, ByVal arrBIOSInstances)
    'region FunctionMetadata ####################################################
    ' Assuming that arrBIOSInstances represents an array / collection of the available BIOS
    ' instances (of type Win32_BIOS), this function obtains the computer's serial number, if
    ' available and configured by the computer's manufacturer
    '
    ' The function takes two positional arguments:
    '  - The first argument (strSerialNumber) is populated upon success with a string
    '    containing the computer's serial number as reported by WMI.
    '  - The second argument (arrBIOSInstances) is an array/collection of objects of class
    '    Win32_BIOS
    '
    ' The function returns a 0 if the serial number was obtained successfully. It returns a
    ' negative integer if an error occurred retrieving the serial number. Finally, it returns
    ' a positive integer if the serial number was obtained, but multiple BIOS instances were
    ' present that contained data for the serial number. When this happens, only the first
    ' Win32_BIOS instance containing data for the serial number is used.
    '
    ' Example:
    '   intReturnCode = GetBIOSInstances(arrBIOSInstances)
    '   If intReturnCode >= 0 Then
    '       ' At least one Win32_BIOS instance was retrieved successfully
    '       intReturnCode = GetSerialNumberUsingBIOSInstances(strSerialNumber, arrBIOSInstances)
    '       If intReturnCode >= 0 Then
    '           ' The computer serial number was retrieved successfully and is stored in
    '           ' strSerialNumber
    '       End If
    '   End If
    '
    ' Version: 1.0.20210624.0
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
    ' None!
    'endregion Acknowledgements ####################################################

    'region DependsOn ####################################################
    ' TestObjectForData()
    ' TestObjectIsAnyTypeOfInteger()
    'endregion DependsOn ####################################################

    Dim intFunctionReturn
    Dim intReturnMultiplier
    Dim intTemp
    Dim intCounterA
    Dim strInterimResult
    Dim strOldInterimResult
    Dim strResultToReturn
    Dim intCountOfBIOSes

    Err.Clear

    intFunctionReturn = 0
    intReturnMultiplier = 128
    strInterimResult = ""
    strResultToReturn = ""
    intCountOfBIOSes = 0

    If TestObjectForData(arrBIOSInstances) <> True Then
        intFunctionReturn = intFunctionReturn + (-1 * intReturnMultiplier)
    Else
        On Error Resume Next
        intTemp = arrBIOSInstances.Count
        If Err Then
            On Error Goto 0
            Err.Clear
            intFunctionReturn = intFunctionReturn + (-2 * intReturnMultiplier)
        Else
            On Error Goto 0
            If TestObjectIsAnyTypeOfInteger(intTemp) = False Then
                intFunctionReturn = intFunctionReturn + (-3 * intReturnMultiplier)
            Else
                If intTemp < 0 Then
                    intFunctionReturn = intFunctionReturn + (-4 * intReturnMultiplier)
                ElseIf intTemp = 0 Then
                    intFunctionReturn = intFunctionReturn + (-5 * intReturnMultiplier)
                Else
                    For intCounterA = 0 To (intTemp - 1)
                        strOldInterimResult = strInterimResult
                        On Error Resume Next
                        strInterimResult = arrBIOSInstances.ItemIndex(intCounterA).SerialNumber
                        If Err Then
                            On Error Goto 0
                            Err.Clear
                            strInterimResult = strOldInterimResult
                        Else
                            On Error Goto 0
                            If TestObjectForData(strInterimResult) <> True Then
                                strInterimResult = strOldInterimResult
                            Else
                                ' Found a result with real model data
                                If TestObjectForData(strResultToReturn) = False Then
                                    strResultToReturn = strInterimResult
                                End If
                                intCountOfBIOSes = intCountOfBIOSes + 1
                            End If
                        End If
                    Next
                End If
            End If
        End If
    End If

    If intFunctionReturn >= 0 Then
        ' No error has occurred yet
        If intCountOfBIOSes = 0 Then
            ' No result found
            intFunctionReturn = intFunctionReturn + (-5 * intReturnMultiplier)
        Else
            intFunctionReturn = intCountOfBIOSes - 1
        End If
    End If

    If intFunctionReturn >= 0 Then
        strSerialNumber = strResultToReturn
    End If
    
    GetSerialNumberUsingBIOSInstances = intFunctionReturn
End Function
