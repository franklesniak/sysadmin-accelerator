Function TestComputerIsVirtualMachineUsingManufacturerAndModel(ByRef boolIsVM, ByVal strComputerManufacturer, ByVal strComputerModel)
    'region FunctionMetadata #######################################################
    ' Assuming that strComputerManufacturer is a string that is populated with the
    ' computer manufacturer and strComputerModel is a string that is populated with the
    ' computer model name or model number, this function determines if the computer is
    ' a virtual machine
    '
    ' The function takes three positional arguments:
    '  - The first argument (boolIsVM) is populated upon success with a boolean value:
    '    True when the computer was determined to be a virtual machine, False when the
    '    computer was determined to not be a virtual machine
    '  - The second argument (strComputerManufacturer) is a string that must be pre-
    '    populated with the computer manufacturer
    '  - The third argument (strComputerModel) is a string that must be pre-populated
    '    with the computer's model name or model number
    '
    ' The function returns a 0 when the function successfully evaluated whether the
    ' computer is a virtual machine. The function returns a negative integer if an
    ' error occurred.
    '
    ' Example:
    '   intReturnCode = GetComputerSystemInstances(arrComputerSystemInstances)
    '   If intReturnCode >= 0 Then
    '       ' At least one Win32_ComputerSystem instance was retrieved successfully
    '       intReturnCode = GetComputerManufacturerUsingComputerSystemInstances(strComputerManufacturer, arrComputerSystemInstances)
    '       If intReturnCode >= 0 Then
    '           ' The computer manufacturer was retrieved successfully and is stored in
    '           ' strComputerManufacturer
    '           intReturnCode = GetComputerModelUsingComputerSystemInstances(strComputerModel, arrComputerSystemInstances)
    '           If intReturnCode >= 0 Then
    '               ' The computer model name/number was retrieved successfully and is
    '               ' stored in strComputerModel
    '               intReturnCode = TestComputerIsVirtualMachineUsingManufacturerAndModel(boolIsVM, strComputerManufacturer, strComputerModel)
    '               If intReturnCode = 0 Then
    '                   ' Successfully tested whether this system is a VM
    '                   If boolIsVM = True Then
    '                       ' Computer is a VM
    '                   Else
    '                       ' Computer is not a VM
    '                   End If
    '               End If
    '           End If
    '       End If
    '  End If
    '
    ' Version: 1.1.20230423.0
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
    ' Microsoft, for (intentionally or not) making the Microsoft Deployment Toolkit
    ' (MDT) with its source code viewable based on it being written in VBS/WSH. MDT has
    ' a function that determines whether a given system is a desktop/laptop/server/VM,
    ' which was useful in determining how to approach this function
    '
    ' Google, for documenting the procedure for detecting whether a virtual machine is
    ' running on Google Compute Engine:
    ' https://cloud.google.com/compute/docs/instances/detect-compute-engine#windows-vm_1
    '
    ' Michael Niehaus for comfirming that the model number of a VMware virtual machine
    ' may show up as "VMware20,1":
    ' https://oofhours.com/2022/11/25/now-released-vmware-fusion-for-running-windows-on-arm-on-m1-m2-macs/
    'endregion Acknowledgements #######################################################

    'region DependsOn ##############################################################
    ' TestObjectIsStringContainingData()
    'endregion DependsOn ##############################################################

    Dim intFunctionReturn
    Dim intReturnMultiplier
    Dim boolInterimResult

    intFunctionReturn = 0
    intReturnMultiplier = 2048

    If TestObjectIsStringContainingData(strComputerModel) <> True Then
        intFunctionReturn = intFunctionReturn + (-1 * intReturnMultiplier)
    Else
        boolInterimResult = False
        If strComputerModel = "Virtual Machine" Then
            ' Microsoft virtual machine
            boolInterimResult = True
        ElseIf strComputerModel = "VMware Virtual Platform" Or strComputerModel = "VMware7,1" Or strComputerModel = "VMware20,1" Then
            ' VMware virtual machine
            boolInterimResult = True
        ElseIf strComputerModel = "VirtualBox" Then
            ' VirtualBox virtual machine
            boolInterimResult = True
        ElseIf strComputerModel = "Parallels Virtual Platform" Then
            ' Parallels virtual machine
            boolInterimResult = True
        ElseIf strComputerModel = "Google Compute Engine" Then
            ' Google Compute Engine virtual machine
            boolInterimResult = True
        Else
            If TestObjectIsStringContainingData(strComputerManufacturer) <> True Then
                intFunctionReturn = intFunctionReturn + (-2 * intReturnMultiplier)
            Else
                If strComputerManufacturer = "Xen" Then
                    ' Citrix Xen virtual machine
                    ' Note: could also be running on AWS
                    boolInterimResult = True
                ElseIf strComputerManufacturer = "QEMU" Then
                    ' QEMU / KVM virtual machine
                    boolInterimResult = True
                End If
            End If
        End If
    End If

    If intFunctionReturn = 0 Then
        boolIsVM = boolInterimResult
    End If
    
    TestComputerIsVirtualMachineUsingManufacturerAndModel = intFunctionReturn
End Function
