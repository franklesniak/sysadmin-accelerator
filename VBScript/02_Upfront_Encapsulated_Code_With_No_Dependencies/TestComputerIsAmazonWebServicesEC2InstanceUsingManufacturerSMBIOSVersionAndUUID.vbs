Function TestComputerIsAmazonWebServicesEC2InstanceUsingManufacturerSMBIOSVersionAndUUID(ByRef boolIsAmazonWebServicesEC2Instance, ByVal strComputerManufacturer, ByVal strSMBIOSVersion, ByVal strUUID)
    'region FunctionMetadata #######################################################
    ' Assuming that strComputerManufacturer is a string that is populated with the
    ' computer's manufacturer, strSMBIOSVersion is a string that is populated with the
    ' computer's SMBIOS version and strUUID is a string populated with the computer's
    ' universal unique ID (derived from Win32_ComputerSystemProduct -> UUID), this
    ' function determines if the computer is an Amazon Web Services (AWS) EC2 instance
    '
    ' The function takes four positional arguments:
    '  - The first argument (boolIsAmazonWebServicesEC2Instance) is populated upon
    '    success with a boolean value: True when the computer was determined to be an
    '    AWS EC2 instance, False when the computer was determined to not be an AWS EC2
    '    instance
    '  - The second argument (strComputerManufacturer) is a string that must be pre-
    '    populated with the computer manufacturer
    '  - The third argument (strSMBIOSVersion) is a string that must be pre-populated
    '    with the SMBIOS version of the computer
    '  - The fourth argument (strUUID) is a string that must be pre-populated with the
    '    computer's universal unique ID (derived from Win32_ComputerSystemProduct ->
    '    UUID)
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
    '           intReturnCode = GetComputerManufacturerUsingComputerSystemInstances(strComputerManufacturer, arrComputerSystemInstances)
    '           If intReturnCode >= 0 Then
    '               ' The computer manufacturer was retrieved successfully and is
    '               ' stored in strComputerManufacturer
    '               intReturnCode = GetComputerSystemProductInstancesUsingWMINamespace(arrComputerSystemProductInstances, objSWbemServicesWMINamespace)
    '               If intReturnCode >= 0 Then
    '                   ' At least one Win32_ComputerSystemProduct instance was
    '                   ' retrieved successfully
    '                   intReturnCode = GetComputerUUIDUsingComputerSystemProductInstances(strComputerUUID, arrComputerSystemProductInstances)
    '                   If intReturnCode >= 0 Then
    '                       ' The computer's UUID was retrieved successfully and is
    '                       ' stored in strComputerUUID
    '                       intReturnCode = GetBIOSInstancesUsingWMINamespace(arrBIOSInstances, objSWbemServicesWMINamespace)
    '                       If intReturnCode >= 0 Then
    '                           ' At least one Win32_BIOS instance was retrieved
    '                           ' successfully
    '                           intReturnCode = GetSMBIOSVersionStringUsingBIOSInstances(strSMBIOSVersion, arrBIOSInstances)
    '                           If intReturnCode >= 0 Then
    '                               ' The systems management BIOS version string was
    '                               ' retrieved successfully and is stored in
    '                               ' strSMBIOSVersion
    '                               intReturnCode = TestComputerIsAmazonWebServicesEC2InstanceUsingManufacturerSMBIOSVersionAndUUID(boolIsAmazonWebServicesEC2Instance, strComputerManufacturer, strSMBIOSVersion, strComputerUUID)
    '                               If intReturnCode = 0 Then
    '                                   ' Successfully tested whether this system is an
    '                                   ' AWS EC2 instance
    '                                   If boolIsAmazonWebServicesEC2Instance = True Then
    '                                       ' Computer is an AWS EC2 instance
    '                                   Else
    '                                       ' Computer is not an AWS EC2 instance
    '                                   End If
    '                               End If
    '                           End If
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

    'region Acknowledgements #######################################################
    ' Microsoft, for affirming my assumptions that SMBIOS versions older than 2.6 were
    ' not required to follow little-endian byte order:
    ' https://learn.microsoft.com/en-us/azure/virtual-machines/instance-metadata-service?tabs=windows
    '
    ' Amazon Web Services for providing good documentation on how to determine whether
    ' a given system is an Amazon Web Services EC2 instance:
    ' https://docs.aws.amazon.com/AWSEC2/latest/WindowsGuide/identify_ec2_instances.html
    'endregion Acknowledgements #######################################################

    'region DependsOn ##############################################################
    ' TestObjectIsStringGUID()
    ' TestObjectIsStringContainingData()
    ' ConvertStringVersionNumberToMajorMinorBuildRevisionIntegers()
    'endregion DependsOn ##############################################################

    Dim intFunctionReturn
    Dim intReturnMultiplier
    Dim intReturnCode
    Dim boolResult
    Dim arrSMBIOSVersion
    Dim strModifiedSMBIOSVersion
    Dim lngMajor
    Dim lngMinor
    Dim lngBuild
    Dim lngRevision
    Dim boolSMBIOSVersionIsOlderThan2Point6
    Dim boolInterimResult

    intFunctionReturn = 0
    intReturnMultiplier = 1024

    If TestObjectIsStringContainingData(strComputerManufacturer) <> True Then
        intFunctionReturn = intFunctionReturn + (-1 * intReturnMultiplier)
    Else
        ' strComputerManufacturer is a string and it is not empty
        If strComputerManufacturer <> "Xen" Then
            ' Not a Citrix Xen virtual machine
            ' Not a candidate for running on AWS EC2
            boolInterimResult = False
            intFunctionReturn = 1
        End If
    End If

    If intFunctionReturn = 0 Then
        ' No error occurred
        ' Computer is a Citrix Xen virtual machine
        If TestObjectIsStringContainingData(strSMBIOSVersion) <> True Then
            intFunctionReturn = intFunctionReturn + (-2 * intReturnMultiplier)
        Else
            ' strSMBIOSVersion is a string and it is not empty
            arrSMBIOSVersion = Split(strSMBIOSVersion, ".")
            If UBound(arrSMBIOSVersion) < 1 Then
                boolInterimResult = False
                intFunctionReturn = 1
            Else
                ' strSMBIOSVersion contains at least two version number components
                strModifiedSMBIOSVersion = arrSMBIOSVersion(0) & "." & arrSMBIOSVersion(1)
                intReturnCode = ConvertStringVersionNumberToMajorMinorBuildRevisionIntegers(lngMajor, lngMinor, lngBuild, lngRevision, strModifiedSMBIOSVersion)
                If intReturnCode <> 0 Then
                    boolInterimResult = False
                    intFunctionReturn = 1
                Else
                    ' The first two SMBIOS version number components were successfully
                    ' converted to integers
                    If lngMajor < 2 Then
                        ' SMBIOS version is less than 2.0
                        boolSMBIOSVersionIsOlderThan2Point6 = True
                    ElseIf lngMajor = 2 Then
                        ' SMBIOS version is 2.0 or greater
                        If lngMinor < 6 Then
                            ' SMBIOS version is less than 2.6
                            boolSMBIOSVersionIsOlderThan2Point6 = True
                        Else
                            ' SMBIOS version is 2.6 or greater
                            boolSMBIOSVersionIsOlderThan2Point6 = False
                        End If
                    Else
                        ' SMBIOS version is 2.6 or greater
                        boolSMBIOSVersionIsOlderThan2Point6 = False
                    End If
                End If
            End If
        End If
    End If

    If intFunctionReturn = 0 Then
        ' No error occurred
        ' Computer is a Citrix Xen virtual machine
        ' SMBIOS version is in a format that indicates whether it is 2.6 or greater
        intReturnCode = TestObjectIsStringGUID(boolResult, strUUID)
        If intReturnCode <> 0 Then
            intFunctionReturn = intFunctionReturn + (-3 * intReturnMultiplier)
        Else
            If boolResult <> True Then
                intFunctionReturn = intFunctionReturn + (-4 * intReturnMultiplier)
            Else
                ' The UUID is a valid GUID
            End If
        End If
    End If

    If intFunctionReturn = 0 Then
        ' No error occurred
        ' Computer is a Citrix Xen virtual machine
        ' SMBIOS version is in a format that indicates whether it is 2.6 or greater
        ' The UUID is a valid GUID
        If Left(strUUID, 1) = "{" Or Left(strUUID, 1) = "(" Then
            'UUID is wrapped by a curly brace or parenthesis
            If UCase(Mid(strUUID, 2, 3)) = "EC2" Then
                ' Amazon Web Services EC2 instance
                boolInterimResult = True
            Else
                ' Not an Amazon Web Services EC2 instance
                boolInterimResult = False
            End If
        Else
            'UUID is not wrapped by a curly brace or parenthesis
            If UCase(Left(strUUID, 3)) = "EC2" Then
                ' Amazon Web Services EC2 instance
                boolInterimResult = True
            Else
                ' Not an Amazon Web Services EC2 instance
                boolInterimResult = False
            End If
        End If
        
        If boolInterimResult = False Then
            If boolSMBIOSVersionIsOlderThan2Point6 = True Then
                If Left(strUUID, 1) = "{" Or Left(strUUID, 1) = "(" Then
                    'UUID is wrapped by a curly brace or parenthesis
                    If (UCase(Mid(strUUID, 8, 2)) + UCase(Mid(strUUID, 6, 1))) = "EC2" Then
                        ' Amazon Web Services EC2 instance
                        boolInterimResult = True
                    End If
                Else
                    'UUID is not wrapped by a curly brace or parenthesis
                    If (UCase(Mid(strUUID, 7, 2)) + UCase(Mid(strUUID, 5, 1))) = "EC2" Then
                        ' Amazon Web Services EC2 instance
                        boolInterimResult = True
                    End If
                End If
            End If
        End If
    End If


    If intFunctionReturn = 1 Then
        intFunctionReturn = 0
    End If

    If intFunctionReturn = 0 Then
        boolIsAmazonWebServicesEC2Instance = boolInterimResult
    End If
    
    TestComputerIsAmazonWebServicesEC2InstanceUsingManufacturerSMBIOSVersionAndUUID = intFunctionReturn
End Function
