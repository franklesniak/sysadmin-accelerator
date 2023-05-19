Function GetWindowsAutopilotHardwareHashUsingMDMDevDetailExt01Instances(ByRef strWindowsAutopilotHardwareHash, ByVal arrMDMDevDetailExt01Instances)
    'region FunctionMetadata #######################################################
    ' Assuming that arrMDMDevDetailExt01Instances represents an array / collection of
    ' the instances of the class MDM_DevDetail_Ext01., this function obtains the
    ' Windows Autopilot hardware hash (i.e., a raw blob used to identify a device in
    ' the cloud)
    '
    ' The function takes two positional arguments:
    '  - The first argument (strWindowsAutopilotHardwareHash) is populated upon success
    '    with a string containing the Windows Autopilot hardware hash.
    '  - The second argument (arrMDMDevDetailExt01Instances) is an array/collection of
    '    objects of class MDM_DevDetail_Ext01
    '
    ' The function returns a 0 if the Windows Autopilot hardware hash was obtained
    ' successfully. It returns a negative integer if an error occurred retrieving the
    ' Windows Autopilot hardware hash. Finally, it returns a positive integer if the
    ' Windows Autopilot hardware hash was obtained, but multiple MDM_DevDetail_Ext01
    ' instances were present that contained data for the Windows Autopilot hardware
    ' hash. When this happens, only the first MDM_DevDetail_Ext01 instance containing
    ' data for the Windows Autopilot hardware hash is used.
    '
    ' Example:
    '   intReturnCode = GetMDMDevDetailExt01Instances(arrMDMDevDetailExt01Instances)
    '   If intReturnCode >= 0 Then
    '       ' At least one MDM_DevDetail_Ext01 instance was retrieved successfully
    '       intReturnCode = GetWindowsAutopilotHardwareHashUsingMDMDevDetailExt01Instances(strWindowsAutopilotHardwareHash, arrMDMDevDetailExt01Instances)
    '       If intReturnCode >= 0 Then
    '           ' The Windows Autopilot hardware hash was retrieved successfully and is
    '           ' stored in strWindowsAutopilotHardwareHash
    '       End If
    '   End If
    '
    ' Version: 1.1.20230518.0
    'endregion FunctionMetadata #######################################################

    'region License ################################################################
    ' Copyright 2023 Frank Lesniak
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
    ' Michael Niehaus, who wrote the script Get-WindowsAutoPilotInfo, which is where I
    ' learned about this WMI namespace:
    ' https://www.powershellgallery.com/packages/Get-WindowsAutoPilotInfo/
    'endregion Acknowledgements #######################################################

    'region DependsOn ##############################################################
    ' TestObjectForData()
    ' TestObjectIsAnyTypeOfInteger()
    'endregion DependsOn ##############################################################

    Dim intFunctionReturn
    Dim intReturnMultiplier
    Dim intTemp
    Dim strInterimResult
    Dim objMDMDevDetailExt01Instance
    Dim strOldInterimResult
    Dim strResultToReturn
    Dim intCountOfHardwareHashes

    Err.Clear

    intFunctionReturn = 0
    intReturnMultiplier = 128
    strInterimResult = ""
    strResultToReturn = ""
    intCountOfHardwareHashes = 0

    If TestObjectForData(arrMDMDevDetailExt01Instances) <> True Then
        intFunctionReturn = intFunctionReturn + (-1 * intReturnMultiplier)
    Else
        On Error Resume Next
        intTemp = arrMDMDevDetailExt01Instances.Count
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
                    For Each objMDMDevDetailExt01Instance in arrMDMDevDetailExt01Instances
                        If Err Then
                            Err.Clear
                        Else
                            strOldInterimResult = strInterimResult
                            strInterimResult = objMDMDevDetailExt01Instance.DeviceHardwareData
                            If Err Then
                                Err.Clear
                                strInterimResult = strOldInterimResult
                            Else
                                If TestObjectForData(strInterimResult) <> True Then
                                    strInterimResult = strOldInterimResult
                                Else
                                    ' Found a result with a real hardware hash
                                    If TestObjectForData(strResultToReturn) = False Then
                                        strResultToReturn = strInterimResult
                                    End If
                                    intCountOfHardwareHashes = intCountOfHardwareHashes + 1
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
        If intCountOfHardwareHashes = 0 Then
            ' No result found
            intFunctionReturn = intFunctionReturn + (-5 * intReturnMultiplier)
        Else
            intFunctionReturn = intCountOfHardwareHashes - 1
        End If
    End If

    If intFunctionReturn >= 0 Then
        strWindowsAutopilotHardwareHash = strResultToReturn
    End If
    
    GetWindowsAutopilotHardwareHashUsingMDMDevDetailExt01Instances = intFunctionReturn
End Function
