Function ConnectLocalMDMWMIBridgeNamespace(ByRef objSWbemServicesMDMWMIBridgeNamespace)
    'region FunctionMetadata #######################################################
    ' Connects to the local computer's mobile device management (MDM) WMI Bridge
    ' Namespace ("root\cimv2\mdm\dmmap"). Requires administrative rights for a
    ' successful connection
    '
    ' The function returns 0 upon success. When the function fails, it returns a negative
    ' integer between -1 and -10 (inclusive), depending on the specific failure that occurs.
    '
    ' Example:
    '   intReturnCode = ConnectLocalMDMWMIBridgeNamespace(objSWbemServicesMDMWMIBridgeNamespace)
    '   If intReturnCode = 0 Then
    '       ' Successfully connected to the local computer's MDM WMI Bridge Namespace
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
    ' Peter van der Woude, who told me that the WMI namespace root\cimv2\mdm\dmmap is
    ' the MDM WMI Bridge provider:
    ' https://www.petervanderwoude.nl/post/windows-10-mdm-powershell-scripting/
    'endregion Acknowledgements #######################################################

    'region DependsOn ##############################################################
    ' ConnectLocalWMINamespace()
    'endregion DependsOn ##############################################################

    Dim intFunctionReturnCode
    intFunctionReturnCode = 0

    intFunctionReturnCode = ConnectLocalWMINamespace(objSWbemServicesMDMWMIBridgeNamespace, "root\cimv2\mdm\dmmap", Null)

    ConnectLocalMDMWMIBridgeNamespace = intFunctionReturnCode
End Function
