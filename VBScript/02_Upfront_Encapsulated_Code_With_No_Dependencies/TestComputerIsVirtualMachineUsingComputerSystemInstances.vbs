Function TestComputerIsVirtualMachineUsingComputerSystemInstances(ByRef boolIsVM, ByVal arrComputerSystemInstances)
    'region FunctionMetadata ####################################################
    ' Assuming that arrComputerSystemInstances represents an array / collection of the
    ' available computer system instances (of type Win32_ComputerSystem), this function
    ' determines if the computer is a virtual machine
    '
    ' The function takes two positional arguments:
    '  - The first argument (boolIsVM) is populated upon success with a boolean value: True
    '    when the computer was determined to be a virtual machine, False when the computer was
    '    determined to not be a virtual machine
    '  - The second argument (arrComputerSystemInstances) is an array/collection of objects of class
    '    Win32_ComputerSystem
    '
    ' The function returns a 0 when the function successfully evaluated whether the computer is
    ' a virtual machine. The function returns a negative integer if an error occurred.
    '
    ' Example:
    '   intReturnCode = GetComputerSystemInstances(arrComputerSystemInstances)
    '   If intReturnCode >= 0 Then
    '       ' At least one Win32_ComputerSystem instance was retrieved successfully
    '       intReturnCode = TestComputerIsVirtualMachineUsingComputerSystemInstances(boolIsVM, arrComputerSystemInstances)
    '       If intReturnCode = 0 Then
    '           ' Successfully tested whether this system is a VM
    '           If boolIsVM = True Then
    '               ' Computer is a VM
    '           Else
    '               ' Computer is not a VM
    '           End If
    '       End If
    '  End If
    '
    ' Version: 1.0.20210625.0
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
    ' GetComputerManufacturerUsingComputerSystemInstances()
    ' GetComputerModelUsingComputerSystemInstances()
    ' TestComputerIsVirtualMachineUsingManufacturerAndModel()
    'endregion DependsOn ####################################################

    Dim intFunctionReturn
    Dim intReturnMultiplier
    Dim intReturnCode
    Dim boolInterimResult
    Dim strComputerManufacturer
    Dim strComputerModel

    intFunctionReturn = 0
    intReturnMultiplier = 1

    If TestObjectForData(arrComputerSystemInstances) <> True Then
        intFunctionReturn = intFunctionReturn + (-1 * intReturnMultiplier)
    Else
        intReturnCode = GetComputerManufacturerUsingComputerSystemInstances(strComputerManufacturer, arrComputerSystemInstances)
        If intReturnCode < 0 Then
            intFunctionReturn = intFunctionReturn + (intReturnCode * intReturnMultiplier)
        Else
            ' intReturnCode >= 0
            ' The computer manufacturer was retrieved successfully and is stored in
            ' strComputerManufacturer
            intReturnMultiplier = intReturnMultiplier * 2
            intReturnCode = GetComputerModelUsingComputerSystemInstances(strComputerModel, arrComputerSystemInstances)
            If intReturnCode < 0 Then
                intFunctionReturn = intFunctionReturn + (intReturnCode * intReturnMultiplier)
            Else
                ' intReturnCode >= 0
                ' The computer model name/number was retrieved successfully and is stored in
                ' strComputerModel
                intReturnMultiplier = 1
                intReturnCode = TestComputerIsVirtualMachineUsingManufacturerAndModel(boolInterimResult, strComputerManufacturer, strComputerModel)
                If intReturnCode <> 0 Then
                    intFunctionReturn = intFunctionReturn + (intReturnCode * intReturnMultiplier)
                End If
            End If
        End If
    End If

    If intFunctionReturn = 0 Then
        boolIsVM = boolInterimResult
    End If
    
    TestComputerIsVirtualMachineUsingComputerSystemInstances = intFunctionReturn
End Function
