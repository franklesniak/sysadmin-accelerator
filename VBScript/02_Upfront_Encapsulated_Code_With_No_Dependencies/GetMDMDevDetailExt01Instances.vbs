Function GetMDMDevDetailExt01Instances(ByRef arrMDMDevDetailExt01Instances)
    'region FunctionMetadata #######################################################
    ' This function retrieves the available instances as objects of the class
    ' MDM_DevDetail_Ext01. This class is one of several that provides device-specific
    ' parameters relevant to the Open Mobile Alliance (OMA) device management (DM;
    ' together: OMA-DM) server.
    '
    ' The MDM_DevDetail_Ext01 class contains two useful properties:
    '  - DeviceHardwareData: A string that contains the device's unique identifier
    '    (hardware hash). This property was added in Windows 10 version 1703. It
    '    returns a base64 encoded string of the hardware parameters of a device.
    '  - WLANMACAddress: A string that contains the MAC address of the device's active
    '    wireless network adapter. This property was added in Windows 10 version 1511.
    '
    ' The function takes one positional argument (arrMDMDevDetailExt01Instances), which
    ' is populated upon success with a collection of instances of the type
    ' MDM_DevDetail_Ext01
    '
    ' The function returns 0 if MDM_DevDetail_Ext01 instances were retrieved
    ' successfully, and there was one MDM_DevDetail_Ext01 instance (as expected). If no
    ' MDM_DevDetail_Ext01 objects could be retrieved, then the function returns a
    ' negative number. If there are unexpectedly multiple instances of 
    ' MDM_DevDetail_Ext01, then the function returns a positive number equal to the
    ' number of WMI instances retrieved minus one.
    '
    ' Example:
    '   intReturnCode = GetMDMDevDetailExt01Instances(arrMDMDevDetailExt01Instances)
    '   If intReturnCode = 0 Then
    '       ' The MDM_DevDetail_Ext01 instance was retrieved successfully and is
    '       ' available at arrMDMDevDetailExt01Instances.ItemIndex(0)
    '   ElseIf intReturnCode > 0 Then
    '       ' More than one MDM_DevDetail_Ext01 instance was retrieved, which is
    '       ' unexpected.
    '   Else
    '       ' An error occurred and no MDM_DevDetail_Ext01 instances were retrieved
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
    ' The most up-to-date version of this script can be found on the author's GitHub
    ' repository at https://github.com/franklesniak/sysadmin-accelerator
    'endregion DownloadLocationNotice #################################################

    'region Acknowledgements #######################################################
    ' Michael Niehaus, who wrote the script Get-WindowsAutoPilotInfo, which is where I
    ' learned about this WMI namespace:
    ' https://www.powershellgallery.com/packages/Get-WindowsAutoPilotInfo/
    '
    ' Microsoft, for publishing some details on the MDM_DevDetail_Ext01 class:
    ' https://learn.microsoft.com/en-us/windows/win32/dmwmibridgeprov/mdm-devdetail-ext01
    '
    ' Microsoft, for publishing details on the DevDetail CSP:
    ' https://learn.microsoft.com/en-us/windows/client-management/mdm/devdetail-csp
    'endregion Acknowledgements #######################################################

    'region DependsOn ##############################################################
    ' ConnectLocalMDMWMIBridgeNamespace()
    ' GetMDMDevDetailExt01InstancesUsingWMINamespace()
    'endregion DependsOn ##############################################################

    Dim intFunctionReturn
    Dim intReturnMultiplier
    Dim intReturnCode
    Dim objSWbemServicesMDMWMIBridgeNamespace

    intFunctionReturn = 0
    intReturnMultiplier = 1

    intReturnCode = ConnectLocalMDMWMIBridgeNamespace(objSWbemServicesMDMWMIBridgeNamespace)
    If intReturnCode < 0 Then
        intFunctionReturn = intFunctionReturn + (intReturnCode * intReturnMultiplier)
    Else
        intReturnCode = GetMDMDevDetailExt01InstancesUsingWMINamespace(arrMDMDevDetailExt01Instances, objSWbemServicesMDMWMIBridgeNamespace)
        intFunctionReturn = intFunctionReturn + (intReturnCode * intReturnMultiplier)
    End If
    
    GetMDMDevDetailExt01Instances = intFunctionReturn
End Function
