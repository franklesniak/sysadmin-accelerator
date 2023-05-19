Function TestComputerIsStationaryNonServerPhysicalComputerUsingManufacturerModelAndSystemEnclosureInstances(ByRef boolIsStationaryNonServerPhysicalComputer, ByVal strComputerManufacturer, ByVal strComputerModel, ByVal arrSystemEnclosureInstances)
    'region FunctionMetadata ####################################################
    ' Assuming that strComputerManufacturer is a string that is populated with the computer
    ' manufacturer, strComputerModel is a string that is populated with the computer model
    ' name or model number, and arrSystemEnclosureInstances is an array/collection of
    ' Win32_SystemEnclosure objects, this function determines if the computer is a stationary,
    ' non-server, physical computer (i.e., a physical desktop)
    '
    ' The function takes four positional arguments:
    '  - The first argument (boolIsStationaryNonServerPhysicalComputer) is populated upon
    '    success with a boolean value: True when the computer was determined to be a
    '    stationary, non-server, physical computer (i.e., a physical desktop), False otherwise
    '  - The second argument (strComputerManufacturer) is a string that must be pre-populated
    '    with the computer manufacturer
    '  - The third argument (strComputerModel) is a string that must be pre-populated with the
    '    computer's model name or model number
    '  - The fourth argument (arrSystemEnclosureInstances) is a WMI collection/array that must
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
    '           intReturnCode = GetComputerManufacturerUsingComputerSystemInstances(strComputerManufacturer, arrComputerSystemInstances)
    '           If intReturnCode >= 0 Then
    '               ' The computer manufacturer was retrieved successfully and is stored in
    '               ' strComputerManufacturer
    '               intReturnCode = GetComputerModelUsingComputerSystemInstances(strComputerModel, arrComputerSystemInstances)
    '               If intReturnCode >= 0 Then
    '                   ' The computer model name/number was retrieved successfully and is
    '                   ' stored in strComputerModel
    '                   intReturnCode = GetSystemEnclosureInstancesUsingWMINamespace(arrSystemEnclosureInstances, objSWbemServicesWMINamespace)
    '                   If intReturnCode > 0 Then
    '                       ' One or more Win32_SystemEnclosure instances were retrieved. The
    '                       ' first instance is available at
    '                       ' arrSystemEnclosureInstances.ItemIndex(0) and the number of
    '                       ' instances is available at arrSystemEnclosureInstances.Count. In
    '                       ' other words, the upper array boundary/index is
    '                       ' (arrSystemEnclosureInstances.Count - 1).
    '                       intReturnCode = TestComputerIsStationaryNonServerPhysicalComputerUsingManufacturerModelAndSystemEnclosureInstances(boolIsStationaryNonServerPhysicalComputer, strComputerManufacturer, strComputerModel, arrSystemEnclosureInstances)
    '                       If intReturnCode = 0 Then
    '                           ' Successfully tested whether this system is a stationary,
    '                           ' non-server physical computer
    '                           If boolIsStationaryNonServerPhysicalComputer = True Then
    '                               ' Computer is a stationary, non-server physical computer
    '                               ' (e.g., desktop system)
    '                           Else
    '                               ' Computer is not a stationary, non-server physical
    '                               ' computer, i.e., it is a portable computer (laptop,
    '                               ' tablet, etc.), a physical server chassis, or it is a
    '                               ' virtual machine
    '                           End If
    '                       End If
    '                   End If
    '               End If
    '           End If
    '       End If
    '  End If
    '
    ' Version: 1.1.20230518.0
    'endregion FunctionMetadata ####################################################

    'region License ####################################################
    ' Copyright 2023 Frank Lesniak
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
    ' Microsoft, for (intentionally or not) making the Microsoft Deployment Toolkit (MDT) with
    ' its source code viewable based on it being written in VBS/WSH. MDT has a function that
    ' determines whether a given system is a desktop/laptop/server/VM, which was useful in
    ' determining how to approach this function
    'endregion Acknowledgements ####################################################

    'region DependsOn ####################################################
    ' TestComputerIsVirtualMachineUsingManufacturerAndModel()
    ' TestObjectForData()
    ' TestObjectIsAnyTypeOfInteger()
    ' TestSystemEnclosureInstanceIsDockingStation()
    ' TestObjectIsStringContainingData()
    ' TestWin32SystemEnclosureChassisTypeIsStationaryNonServerComputer()
    'endregion DependsOn ####################################################

    Dim intFunctionReturn
    Dim intReturnMultiplier
    Dim boolInterimResult
    Dim objSystemEnclosureInstance
    Dim intReturnCode
    Dim boolIsVirtualMachine
    Dim intTemp
    Dim boolInstanceIsDockingStation
    Dim arrChassisTypes
    Dim intChassisType

    Const VARTYPE_ARRAY = 8204

    intFunctionReturn = 0
    intReturnMultiplier = 1

    boolInterimResult = False

    intReturnCode = TestComputerIsVirtualMachineUsingManufacturerAndModel(boolIsVirtualMachine, strComputerManufacturer, strComputerModel)
    If intReturnCode <> 0 Then
        intFunctionReturn = intFunctionReturn + (intReturnCode * intReturnMultiplier)
    Else
        intReturnMultiplier = intReturnMultiplier * 4
        ' boolIsVirtualMachine is populated with True or False
        If boolIsVirtualMachine = False Then
            If TestObjectForData(arrSystemEnclosureInstances) <> True Then
                intFunctionReturn = intFunctionReturn + (-1 * intReturnMultiplier)
            Else
                On Error Resume Next
                intTemp = arrSystemEnclosureInstances.Count
                If Err Then
                    On Error Goto 0
                    Err.Clear
                    intFunctionReturn = intFunctionReturn + (-2 * intReturnMultiplier)
                Else
                    On Error Goto 0
                    If TestObjectIsAnyTypeOfInteger(intTemp) <> True Then
                        intFunctionReturn = intFunctionReturn + (-3 * intReturnMultiplier)
                    Else
                        ' intTemp is an integer
                        intReturnMultiplier = intReturnMultiplier * 4
                        On Error Resume Next
                        For Each objSystemEnclosureInstance in arrSystemEnclosureInstances
                            If Err Then
                                Err.Clear
                            Else
                                If intFunctionReturn = 0 And boolInterimResult = False Then
                                    intReturnCode = TestSystemEnclosureInstanceIsDockingStation(boolInstanceIsDockingStation, objSystemEnclosureInstance)
                                    If intReturnCode <> 0 Then
                                        intFunctionReturn = intFunctionReturn + (intReturnCode * intReturnMultiplier)
                                    Else
                                        If boolInstanceIsDockingStation = False Then
                                            ' The instance specified by
                                            ' objSystemEnclosureInstance is not a docking
                                            ' station
                                            arrChassisTypes = objSystemEnclosureInstance.ChassisTypes
                                            If Err Then
                                                Err.Clear
                                                intFunctionReturn = intFunctionReturn + (-1 * intReturnMultiplier * 128 * 8)
                                            Else
                                                If TestObjectForData(arrChassisTypes) <> True Then
                                                    intFunctionReturn = intFunctionReturn + (-2 * intReturnMultiplier * 128 * 8)
                                                Else
                                                    If VarType(arrChassisTypes) = VARTYPE_ARRAY Then
                                                        ' arrChassisTypes is an array
                                                        For Each intChassisType in arrChassisTypes
                                                            If Err Then
                                                                Err.Clear
                                                            Else
                                                                If TestObjectIsAnyTypeOfInteger(intChassisType) <> True Then
                                                                    If TestObjectIsStringContainingData(intChassisType) <> True Then
                                                                        If intFunctionReturn >= 0 Then
                                                                            intFunctionReturn = intFunctionReturn + (-3 * intReturnMultiplier * 128 * 8)
                                                                        End If
                                                                    Else
                                                                        ' intChassisType was a string. Try to convert it to int
                                                                        intChassisType = CInt(intChassisType)
                                                                        If Err Then
                                                                            Err.Clear
                                                                            intFunctionReturn = intFunctionReturn + (-4 * intReturnMultiplier * 128 * 8)
                                                                        Else
                                                                            ' intChassisType is now an integer
                                                                            If TestWin32SystemEnclosureChassisTypeIsStationaryNonServerComputer(intChassisType) = True Then
                                                                                boolInterimResult = True
                                                                            End If
                                                                        End If
                                                                    End If
                                                                Else
                                                                    ' intChassisType is an integer
                                                                    If TestWin32SystemEnclosureChassisTypeIsStationaryNonServerComputer(intChassisType) = True Then
                                                                        boolInterimResult = True
                                                                    End If
                                                                End If
                                                            End If
                                                        Next
                                                        If Err Then
                                                            Err.Clear
                                                        End If
                                                    ElseIf TestObjectIsAnyTypeOfInteger(arrChassisTypes) = True Then
                                                        ' arrChassisTypes is a single integer
                                                        boolInterimResult = TestWin32SystemEnclosureChassisTypeIsStationaryNonServerComputer(intChassisType)
                                                    Else
                                                        ' arrChassisTypes was not an array nor an
                                                        ' integer
                                                        intFunctionReturn = intFunctionReturn + (-5 * intReturnMultiplier * 128 * 8)
                                                    End If
                                                End If
                                            End If
                                        End If
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
    End If

    If intFunctionReturn = 0 Then
        boolIsStationaryNonServerPhysicalComputer = boolInterimResult
    End If
    
    TestComputerIsStationaryNonServerPhysicalComputerUsingManufacturerModelAndSystemEnclosureInstances = intFunctionReturn
End Function
