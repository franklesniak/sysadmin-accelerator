Function TestComputerIsAmazonWebServicesEC2Instance(ByRef boolIsAmazonWebServicesEC2Instance)
    'region FunctionMetadata #######################################################
    ' This function determines if the computer is an Amazon Web Services (AWS) EC2
    ' instance
    '
    ' The function takes one positional argument (boolIsAmazonWebServicesEC2Instance),
    ' which is populated upon success with a boolean value: True when the computer was
    ' determined to be an AWS EC2 instance, False when the computer was determined to
    ' not be an AWS EC2 instance
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
    '   intReturnCode = TestComputerIsAmazonWebServicesEC2Instance(boolIsAmazonWebServicesEC2Instance)
    '   If intReturnCode = 0 Then
    '       ' Successfully tested whether this system is an AWS EC2 instance
    '       If boolIsAmazonWebServicesEC2Instance = True Then
    '           ' Computer is an AWS EC2 instance
    '       Else
    '           ' Computer is not an AWS EC2 instance
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
    ' ConnectLocalWMINamespace()
    ' GetComputerSystemInstancesUsingWMINamespace()
    ' GetComputerSystemProductInstancesUsingWMINamespace()
    ' GetBIOSInstancesUsingWMINamespace()
    'endregion DependsOn ##############################################################

    Dim intFunctionReturn
    Dim intReturnMultiplier
    Dim intReturnCode
    Dim objSWbemServicesWMINamespace
    Dim arrComputerSystemInstances
    Dim arrComputerSystemProductInstances
    Dim arrBIOSInstances
    Dim boolInterimResult

    intFunctionReturn = 0
    intReturnMultiplier = 1

    intReturnCode = ConnectLocalWMINamespace(objSWbemServicesWMINamespace, Null, Null)
    If intReturnCode = 0 Then
        ' Successfully connected to the local computer's root\CIMv2 WMI Namespace
    Else
        ' An error occurred
        intFunctionReturn = intFunctionReturn + (intReturnCode * intReturnMultiplier)
    End If

    If intFunctionReturn = 0 Then
        ' No error occurred
        intReturnMultiplier = intReturnMultiplier * 16
        intReturnCode = GetComputerSystemInstancesUsingWMINamespace(arrComputerSystemInstances, objSWbemServicesWMINamespace)
        If intReturnCode >= 0 Then
            ' At least one Win32_ComputerSystem instance was retrieved successfully
        Else
            ' An error occurred
            intFunctionReturn = intFunctionReturn + (intReturnCode * intReturnMultiplier)
        End If
    End If

    If intFunctionReturn = 0 Then
        ' No error occurred
        intReturnMultiplier = intReturnMultiplier * 8
        intReturnCode = GetComputerSystemProductInstancesUsingWMINamespace(arrComputerSystemProductInstances, objSWbemServicesWMINamespace)
        If intReturnCode >= 0 Then
            ' At least one Win32_ComputerSystemProduct instance was retrieved
            ' successfully
        Else
            ' An error occurred
            intFunctionReturn = intFunctionReturn + (intReturnCode * intReturnMultiplier)
        End If
    End If

    If intFunctionReturn = 0 Then
        ' No error occurred
        intReturnMultiplier = intReturnMultiplier * 8
        intReturnCode = GetBIOSInstancesUsingWMINamespace(arrBIOSInstances, objSWbemServicesWMINamespace)
        If intReturnCode >= 0 Then
            ' At least one Win32_BIOS instance was retrieved successfully
        Else
            ' An error occurred
            intFunctionReturn = intFunctionReturn + (intReturnCode * intReturnMultiplier)
        End If
    End If

    If intFunctionReturn = 0 Then
        ' No error occurred
        intReturnMultiplier = intReturnMultiplier * 8
        intReturnCode = TestComputerIsAmazonWebServicesEC2InstanceUsingComputerSystemComputerSystemProductAndBIOSInstances(boolInterimResult, arrComputerSystemInstances, arrComputerSystemProductInstances, arrBIOSInstances)
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
    
    TestComputerIsAmazonWebServicesEC2Instance = intFunctionReturn
End Function
