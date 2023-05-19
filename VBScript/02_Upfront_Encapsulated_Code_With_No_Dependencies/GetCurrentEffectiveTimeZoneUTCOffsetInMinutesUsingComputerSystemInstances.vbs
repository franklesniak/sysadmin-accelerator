Function GetCurrentEffectiveTimeZoneUTCOffsetInMinutesUsingComputerSystemInstances(ByRef intCurrentEffectiveTimeZoneUTCOffsetInMinutes, ByVal arrComputerSystemInstances)
    'region FunctionMetadata ####################################################
    ' Assuming that arrComputerSystemInstances represents an array / collection of the
    ' available computer system instances (of type Win32_ComputerSystem), this function obtains
    ' the computer's current effective time zone UTC offset (in minutes), if available. For
    ' example, for a computer in Central (US) Standard Time (CST), the time zone UTC offset
    ' would be -360 because CST is GMT-6.
    '
    ' The function takes two positional arguments:
    '  - The first argument (intCurrentEffectiveTimeZoneUTCOffsetInMinutes) is populated upon
    '    success with a string containing the computer's current effective time zone UTC offset
    '    (in minutes) as reported by WMI.
    '  - The second argument (arrComputerSystemInstances) is an array/collection of objects of
    '    class Win32_ComputerSystem
    '
    ' The function returns a 0 if the current effective time zone UTC offset (in minutes) was
    ' obtained successfully. It returns a negative integer if an error occurred retrieving the
    ' time zone offset. Finally, it returns a positive integer if the time zone offset was
    ' obtained, but multiple computer system instances were present that contained data for the
    ' time zone offset. When this happens, only the first Win32_ComputerSystem instance
    ' containing data for the time zone offset is used.
    '
    ' Example:
    '   intReturnCode = GetComputerSystemInstances(arrComputerSystemInstances)
    '   If intReturnCode >= 0 Then
    '       ' At least one Win32_ComputerSystem instance was retrieved successfully
    '       intReturnCode = GetCurrentEffectiveTimeZoneUTCOffsetInMinutesUsingComputerSystemInstances(intCurrentEffectiveTimeZoneUTCOffsetInMinutes, arrComputerSystemInstances)
    '       If intReturnCode >= 0 Then
    '           ' The computer's current effective time zone UTC offset (in minutes) was
    '           ' retrieved successfully and is stored in
    '           ' intCurrentEffectiveTimeZoneUTCOffsetInMinutes
    '       End If
    '   End If
    '
    ' Version: 1.1.20230518.0
    'endregion FunctionMetadata ####################################################

    'region License ####################################################
    ' Copyright 2023 Frank Lesniak
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
    Dim intInterimResult
    Dim objComputerSystemInstance
    Dim intOldInterimResult
    Dim intResultToReturn
    Dim intCountOfComputerSystems

    Err.Clear

    intFunctionReturn = 0
    intReturnMultiplier = 128
    intInterimResult = Null
    intResultToReturn = Null
    intCountOfComputerSystems = 0

    If TestObjectForData(arrComputerSystemInstances) <> True Then
        intFunctionReturn = intFunctionReturn + (-1 * intReturnMultiplier)
    Else
        On Error Resume Next
        intTemp = arrComputerSystemInstances.Count
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
                    On Error Resume Next
                    For Each objComputerSystemInstance in arrComputerSystemInstances
                        If Err Then
                            Err.Clear
                        Else
                            intOldInterimResult = intInterimResult
                            intInterimResult = objComputerSystemInstance.CurrentTimeZone
                            If Err Then
                                Err.Clear
                                intInterimResult = intOldInterimResult
                            Else
                                If TestObjectForData(intInterimResult) <> True Then
                                    intInterimResult = intOldInterimResult
                                Else
                                    ' Found a result with real time zone data
                                    If TestObjectForData(intResultToReturn) = False Then
                                        intResultToReturn = intInterimResult
                                    End If
                                    intCountOfComputerSystems = intCountOfComputerSystems + 1
                                End If
                            End If
                        End If
                    Next
                    On Error Goto 0
                    If Err Then
                        Err.Clear
                    End If
                End If
            End If
        End If
    End If

    If intFunctionReturn >= 0 Then
        ' No error has occurred yet
        If intCountOfComputerSystems = 0 Then
            ' No result found
            intFunctionReturn = intFunctionReturn + (-5 * intReturnMultiplier)
        Else
            intFunctionReturn = intCountOfComputerSystems - 1
        End If
    End If

    If intFunctionReturn >= 0 Then
        intCurrentEffectiveTimeZoneUTCOffsetInMinutes = intResultToReturn
    End If
    
    GetCurrentEffectiveTimeZoneUTCOffsetInMinutesUsingComputerSystemInstances = intFunctionReturn
End Function
