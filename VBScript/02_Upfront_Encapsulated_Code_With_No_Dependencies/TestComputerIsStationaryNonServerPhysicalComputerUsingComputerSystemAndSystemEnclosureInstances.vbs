Function TestComputerIsStationaryNonServerPhysicalComputerUsingComputerSystemAndSystemEnclosureInstances(ByRef boolIsStationaryNonServerPhysicalComputer, ByVal arrComputerSystemInstances, ByVal arrSystemEnclosureInstances)
    'region FunctionMetadata ####################################################
    ' Assuming that arrComputerSystemInstances represents an array / collection of the
    ' available computer system instances (of type Win32_ComputerSystem) and
    ' arrSystemEnclosureInstances is an array/collection of Win32_SystemEnclosure objects, this
    ' function determines if the computer is a stationary, non-server, physical computer (e.g.,
    ' a physical desktop)
    '
    ' The function takes three positional arguments:
    '  - The first argument (boolIsStationaryNonServerPhysicalComputer) is populated upon
    '    success with a boolean value: True when the computer was determined to be a
    '    stationary, non-server, physical computer (e.g., a physical desktop), False otherwise
    '  - The second argument (arrComputerSystemInstances) is a WMI collection/array that must
    '    be pre-populated with a collection of Win32_ComputerSystem objects
    '  - The third argument (arrSystemEnclosureInstances) is a WMI collection/array that must
    '    be pre-populated with a collection of Win32_SystemEnclosure objects
    '
    ' The function returns a 0 when the function successfully evaluated whether the computer is
    ' a stationary, non-server, physical computer (i.e., a physical desktop). The function
    ' returns a negative integer if an error occurred.
    '
    ' Example:
    '   intReturnCode = ConnectLocalWMINamespace(objSWbemServicesWMINamespace, Null, Null)
    '   If intReturnCode = 0 Then
    '       ' Successfully connected to the local computer's root\CIMv2 WMI Namespace
    '       intReturnCode = GetComputerSystemInstancesUsingWMINamespace(arrComputerSystemInstances, objSWbemServicesWMINamespace)
    '       If intReturnCode >= 0 Then
    '           ' At least one Win32_ComputerSystem instance was retrieved successfully
    '           intReturnCode = GetSystemEnclosureInstancesUsingWMINamespace(arrSystemEnclosureInstances, objSWbemServicesWMINamespace)
    '           If intReturnCode > 0 Then
    '               ' One or more Win32_SystemEnclosure instances were retrieved. The first
    '               ' instance is available at arrSystemEnclosureInstances.ItemIndex(0) and the
    '               ' number of instances is available at arrSystemEnclosureInstances.Count. In
    '               ' other words, the upper array boundary/index is
    '               ' (arrSystemEnclosureInstances.Count - 1).
    '               intReturnCode = TestComputerIsStationaryNonServerPhysicalComputerUsingComputerSystemAndSystemEnclosureInstances(boolIsStationaryNonServerPhysicalComputer, arrComputerSystemInstances, arrSystemEnclosureInstances)
    '               If intReturnCode = 0 Then
    '                   ' Successfully tested whether this system is a stationary, non-server
    '                   ' physical computer
    '                   If boolIsStationaryNonServerPhysicalComputer = True Then
    '                       ' Computer is a stationary, non-server physical computer (e.g.,
    '                       ' desktop system)
    '                   Else
    '                       ' Computer is not a stationary, non-server physical computer, i.e.,
    '                       ' it is a portable computer (laptop, tablet, etc.), a physical
    '                       ' server chassis, or it is a virtual machine
    '                   End If
    '               End If
    '           End If
    '       End If
    '  End If
    '
    ' Version: 1.0.20210627.0
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
    ' TestComputerIsStationaryNonServerPhysicalComputerUsingManufacturerModelAndSystemEnclosureInstances()
    'endregion DependsOn ####################################################

    Dim intFunctionReturn
    Dim intReturnMultiplier
    Dim intReturnCode
    Dim intOffset
    Dim boolInterimResult
    Dim strComputerManufacturer
    Dim strComputerModel

    intFunctionReturn = 0
    intReturnMultiplier = 131072
    intOffset = 0

    boolInterimResult = False

    If TestObjectForData(arrComputerSystemInstances) <> True Then
        intFunctionReturn = intFunctionReturn + (-1 * intReturnMultiplier) + intOffset
    Else
        intFunctionReturn = intFunctionReturn * 2
        intReturnCode = GetComputerManufacturerUsingComputerSystemInstances(strComputerManufacturer, arrComputerSystemInstances)
        If intReturnCode < 0 Then
            ' Error occurred
            intFunctionReturn = intFunctionReturn + (intReturnCode * intReturnMultiplier) + intOffset
        Else
            ' intReturnCode >= 0
            ' The computer manufacturer was retrieved successfully and is stored in
            ' strComputerManufacturer
            intOffset = 536870912
            intReturnCode = GetComputerModelUsingComputerSystemInstances(strComputerModel, arrComputerSystemInstances)
            If intReturnCode < 0 Then
                ' Error occurred
                intFunctionReturn = intFunctionReturn + (intReturnCode * intReturnMultiplier) + intOffset
            Else
                intReturnMultiplier = 1
                intOffset = 0
                intReturnCode = TestComputerIsStationaryNonServerPhysicalComputerUsingManufacturerModelAndSystemEnclosureInstances(boolInterimResult, strComputerManufacturer, strComputerModel, arrSystemEnclosureInstances)
                If intReturnCode <> 0 Then
                    ' Error occurred
                    intFunctionReturn = intFunctionReturn + (intReturnCode * intReturnMultiplier)
                Else
                    'Success
                End If
            End If
        End If
    End If

    If intFunctionReturn = 0 Then
        boolIsStationaryNonServerPhysicalComputer = boolInterimResult
    End If
    
    TestComputerIsStationaryNonServerPhysicalComputerUsingComputerSystemAndSystemEnclosureInstances = intFunctionReturn
End Function
