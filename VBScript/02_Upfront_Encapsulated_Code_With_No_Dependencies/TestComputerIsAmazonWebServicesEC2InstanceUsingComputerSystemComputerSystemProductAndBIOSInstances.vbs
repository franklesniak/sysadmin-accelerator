Function TestComputerIsAmazonWebServicesEC2InstanceUsingComputerSystemComputerSystemProductAndBIOSInstances(ByRef boolIsAmazonWebServicesEC2Instance, ByVal arrComputerSystemInstances, ByVal arrComputerSystemProductInstances, ByVal arrBIOSInstances)
    'region FunctionMetadata #######################################################
    ' Assuming that arrComputerSystemInstances represents an array / collection of the
    ' available computer system instances (of type Win32_ComputerSystem),
    ' arrComputerSystemProductInstances represents an array / collection of the
    ' available computer system instances (of type Win32_ComputerSystemProduct), and
    ' arrBIOSInstances represents an array / collection of the available BIOS instances
    ' (of type Win32_BIOS), this function determines if the computer is an Amazon Web
    ' Services (AWS) EC2 instance
    '
    ' The function takes four positional arguments:
    '  - The first argument (boolIsAmazonWebServicesEC2Instance) is populated upon
    '    success with a boolean value: True when the computer was determined to be an
    '    AWS EC2 instance, False when the computer was determined to not be an AWS EC2
    '    instance
    '  - The second argument (arrComputerSystemInstances) is a WMI collection/array
    '    that must be pre-populated with a collection of Win32_ComputerSystem objects
    '  - The third argument (arrComputerSystemProductInstances) is a WMI collection/
    '    array that must be pre-populated with a collection of
    '    Win32_ComputerSystemProduct objects
    '  - The fourth argument (arrBIOSInstances) is a WMI collection/array that must be
    '    pre-populated with a collection of Win32_BIOS objects
    '
    ' The function returns a 0 when the function successfully evaluated whether the
    ' computer is an AWS EC2 instance. The function returns a negative integer if an
    ' error occurred.
    '
    ' Note: for Citrix Xen virtual machines, this function will return a false positive
    ' on one out of every 4096 virtual machines. This is because AWS EC2 instances are
    ' detectable when the first three digits of their UUID are EC2, but Citrix Xen
    ' virtual machines will have their UUID start with EC2 one out of every 4096 times.
    '
    ' Example:
    '   intReturnCode = ConnectLocalWMINamespace(objSWbemServicesWMINamespace, Null, Null)
    '   If intReturnCode = 0 Then
    '       ' Successfully connected to the local computer's root\CIMv2 WMI Namespace
    '       intReturnCode = GetComputerSystemInstancesUsingWMINamespace(arrComputerSystemInstances, objSWbemServicesWMINamespace)
    '       If intReturnCode >= 0 Then
    '           ' At least one Win32_ComputerSystem instance was retrieved successfully
    '           intReturnCode = GetComputerSystemProductInstancesUsingWMINamespace(arrComputerSystemProductInstances, objSWbemServicesWMINamespace)
    '           If intReturnCode >= 0 Then
    '               ' At least one Win32_ComputerSystemProduct instance was retrieved
    '               ' successfully
    '               intReturnCode = GetBIOSInstancesUsingWMINamespace(arrBIOSInstances, objSWbemServicesWMINamespace)
    '               If intReturnCode >= 0 Then
    '                   ' At least one Win32_BIOS instance was retrieved successfully
    '                   intReturnCode = TestComputerIsAmazonWebServicesEC2InstanceUsingComputerSystemComputerSystemProductAndBIOSInstances(boolIsAmazonWebServicesEC2Instance, arrComputerSystemInstances, arrComputerSystemProductInstances, arrBIOSInstances)
    '                   If intReturnCode = 0 Then
    '                       ' Successfully tested whether this system is an AWS EC2
    '                       ' instance
    '                       If boolIsAmazonWebServicesEC2Instance = True Then
    '                           ' Computer is an AWS EC2 instance
    '                       Else
    '                           ' Computer is not an AWS EC2 instance
    '                       End If
    '                   End If
    '               End If
    '           End If
    '       End If
    '   End If
    '
    ' Version: 1.0.20230423.0
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

    'region DependsOn ##############################################################
    ' TestObjectForData()
    ' GetComputerManufacturerUsingComputerSystemInstances()
    ' GetComputerUUIDUsingComputerSystemProductInstances()
    ' GetSMBIOSVersionStringUsingBIOSInstances()
    ' TestComputerIsAmazonWebServicesEC2InstanceUsingManufacturerSMBIOSVersionAndUUID()
    'endregion DependsOn ##############################################################

    Dim intFunctionReturn
    Dim intReturnMultiplier
    Dim strComputerManufacturer
    Dim strComputerUUID
    Dim strSMBIOSVersion
    Dim boolInterimResult

    intFunctionReturn = 0
    intReturnMultiplier = 1

    If TestObjectForData(arrComputerSystemInstances) <> True Then
        intFunctionReturn = intFunctionReturn + (-1 * intReturnMultiplier)
    Else
        If TestObjectForData(arrComputerSystemProductInstances) <> True Then
            intFunctionReturn = intFunctionReturn + (-2 * intReturnMultiplier)
        Else
            If TestObjectForData(arrBIOSInstances) <> True Then
                intFunctionReturn = intFunctionReturn + (-3 * intReturnMultiplier)
            End If
        End If
    End If

    If intFunctionReturn = 0 Then
        ' No error occurred
        intReturnMultiplier = intReturnMultiplier * 4
        intReturnCode = GetComputerManufacturerUsingComputerSystemInstances(strComputerManufacturer, arrComputerSystemInstances)
        If intReturnCode >= 0 Then
            ' The computer manufacturer was retrieved successfully and is stored in
            ' strComputerManufacturer
        Else
            ' An error occurred
            intFunctionReturn = intFunctionReturn + (intReturnCode * intReturnMultiplier)
        End If
    End If

    If intFunctionReturn = 0 Then
        ' No error occurred
        intReturnMultiplier = intReturnMultiplier * 1024
        intReturnCode = GetComputerUUIDUsingComputerSystemProductInstances(strComputerUUID, arrComputerSystemProductInstances)
        If intReturnCode >= 0 Then
            ' The computer's UUID was retrieved successfully and is stored in
            ' strComputerUUID
        Else
            ' An error occurred
            intFunctionReturn = intFunctionReturn + (intReturnCode * intReturnMultiplier)
        End If
    End If

    If intFunctionReturn = 0 Then
        ' No error occurred
        intReturnMultiplier = intReturnMultiplier * 1024
        intReturnCode = GetSMBIOSVersionStringUsingBIOSInstances(strSMBIOSVersion, arrBIOSInstances)
        If intReturnCode >= 0 Then
            ' The systems management BIOS version string was retrieved
            ' successfully and is stored in strSMBIOSVersion
        Else
            ' An error occurred
            intFunctionReturn = intFunctionReturn + (intReturnCode * intReturnMultiplier)
        End If
    End If

    If intFunctionReturn = 0 Then
        ' No error occurred
        intReturnMultiplier = intReturnMultiplier * 1024
        intReturnCode = TestComputerIsAmazonWebServicesEC2InstanceUsingManufacturerSMBIOSVersionAndUUID(boolInterimResult, strComputerManufacturer, strSMBIOSVersion, strComputerUUID)
        If intReturnCode = 0 Then
            ' Successfully tested whether this system is an AWS EC2 instance
        Else
            ' An error occurred
            intFunctionReturn = intFunctionReturn + (intReturnCode * intReturnMultiplier)
        End If
    End If

    If intFunctionReturn = 0 Then
        ' No error occurred
        boolIsAmazonWebServicesEC2Instance = boolInterimResult
    End If
    
    TestComputerIsAmazonWebServicesEC2InstanceUsingComputerSystemComputerSystemProductAndBIOSInstances = intFunctionReturn
End Function
