Function GetWindowsAutopilotHardwareHash(ByRef strWindowsAutopilotHardwareHash)
    'region FunctionMetadata #######################################################
    ' This function obtains the Windows Autopilot hardware hash (i.e., a raw blob used
    ' to identify a device in the cloud)
    '
    ' The function takes one positional argument (strWindowsAutopilotHardwareHash),
    ' which is populated upon success with a string containing the Windows Autopilot
    ' hardware hash.
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
    '   intReturnCode = GetWindowsAutopilotHardwareHash(strWindowsAutopilotHardwareHash)
    '   If intReturnCode >= 0 Then
    '       ' The Windows Autopilot hardware hash was retrieved successfully and is
    '       ' stored in strWindowsAutopilotHardwareHash
    '   End If
    '
    ' Version: 1.0.20230424.0
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
    ' The most up-to-date version of this script can be found on the author's GitHub repository
    ' at https://github.com/franklesniak/sysadmin-accelerator
    'endregion DownloadLocationNotice #################################################

    'region Acknowledgements #######################################################
    ' Michael Niehaus, who wrote the script Get-WindowsAutoPilotInfo, which is where I
    ' learned about this WMI namespace:
    ' https://www.powershellgallery.com/packages/Get-WindowsAutoPilotInfo/
    'endregion Acknowledgements #######################################################

    'region DependsOn ##############################################################
    ' GetMDMDevDetailExt01Instances()
    ' GetWindowsAutopilotHardwareHashUsingMDMDevDetailExt01Instances()
    'endregion DependsOn ##############################################################

    Dim intFunctionReturn
    Dim arrMDMDevDetailExt01Instances
    Dim strResult

    intFunctionReturn = 0

    intFunctionReturn = GetMDMDevDetailExt01Instances(arrMDMDevDetailExt01Instances)
    If intFunctionReturn >= 0 Then
        ' At least one MDM_DevDetail_Ext01 instance was retrieved successfully
        intFunctionReturn = GetWindowsAutopilotHardwareHashUsingMDMDevDetailExt01Instances(strResult, arrMDMDevDetailExt01Instances)
        If intFunctionReturn >= 0 Then
            ' The computer manufacturer was retrieved successfully and is stored in strResult
            strWindowsAutopilotHardwareHash = strResult
        End If
    End If
    
    GetWindowsAutopilotHardwareHash = intFunctionReturn
End Function
