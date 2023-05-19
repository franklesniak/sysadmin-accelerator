Function TestSystemEnclosureInstanceIsDockingStation(ByRef boolInstanceIsDockingStation, ByVal objSystemEnclosureInstance)
    'region FunctionMetadata ####################################################
    ' Assuming that objSystemEnclosureInstance represents an instance of a
    ' Win32_SystemEnclosure object, this function tests it to see if it is a docking station.
    '
    ' The function takes two positional arguments:
    '  - The first argument (boolInstanceIsDockingStation) is populated upon success with a boolean
    '    True or False. True means that the system enclosure represented by
    '    objSystemEnclosureInstance is a docking station, while False means that it is not a
    '    docking station.
    '  - The second argument (objSystemEnclosureInstance) is instance (object) of class
    '    Win32_SystemEnclosure
    '
    ' The function returns a 0 if the test was performed successfully. It returns a negative
    ' integer if an error occurred performing the test.
    '
    ' Example:
    '   intReturnCode = GetSystemEnclosureInstances(arrSystemEnclosureInstances)
    '   If intReturnCode > 0 Then
    '       ' At least one system enclosure instance was retrieved successfully
    '       For Each objSystemEnclosureInstance in arrSystemEnclosureInstances
    '           intReturnCode = TestSystemEnclosureInstanceIsDockingStation(boolInstanceIsDockingStation, objSystemEnclosureInstance)
    '           If intReturnCode = 0 Then
    '               If boolInstanceIsDockingStation = True Then
    '                   ' The test was successful and the system enclosure was a docking
    '                   ' station
    '               Else
    '                   ' The test was successful and the system enclosure was not a docking
    '                   ' station
    '               End If
    '           Else
    '               ' The test was not successful
    '           End If
    '       Next
    '   End If
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
    ' None!
    'endregion Acknowledgements ####################################################

    'region DependsOn ####################################################
    ' TestObjectForData()
    ' TestObjectIsAnyTypeOfInteger()
    ' TestObjectIsStringContainingData()
    ' TestWin32SystemEnclosureChassisTypeIsDockingStation()
    'endregion DependsOn ####################################################

    Dim intFunctionReturn
    Dim intReturnMultiplier
    Dim arrChassisTypes
    Dim intChassisType
    Dim boolInterimResult

    Const VARTYPE_ARRAY = 8204

    Err.Clear

    intFunctionReturn = 0
    intReturnMultiplier = 128

    If TestObjectForData(objSystemEnclosureInstance) <> True Then
        intFunctionReturn = intFunctionReturn + (-1 * intReturnMultiplier)
    Else
        On Error Resume Next
        arrChassisTypes = objSystemEnclosureInstance.ChassisTypes
        If Err Then
            On Error Goto 0
            Err.Clear
            intFunctionReturn = intFunctionReturn + (-2 * intReturnMultiplier)
        Else
            On Error Goto 0
            If TestObjectForData(arrChassisTypes) <> True Then
                intFunctionReturn = intFunctionReturn + (-3 * intReturnMultiplier)
            Else
                If VarType(arrChassisTypes) = VARTYPE_ARRAY Then
                    ' arrChassisTypes is an array
                    boolInterimResult = False
                    On Error Resume Next
                    For Each intChassisType in arrChassisTypes
                        If Err Then
                            Err.Clear
                        Else
                            If TestObjectIsAnyTypeOfInteger(intChassisType) <> True Then
                                If TestObjectIsStringContainingData(intChassisType) <> True Then
                                    If intFunctionReturn >= 0 Then
                                        intFunctionReturn = intFunctionReturn + (-5 * intReturnMultiplier)
                                    End If
                                Else
                                    ' intChassisType was a string. Try to convert it to int
                                    intChassisType = CInt(intChassisType)
                                    If Err Then
                                        Err.Clear
                                        intFunctionReturn = intFunctionReturn + (-6 * intReturnMultiplier)
                                    Else
                                        ' intChassisType is now an integer
                                        If TestWin32SystemEnclosureChassisTypeIsDockingStation(intChassisType) = True Then
                                            boolInterimResult = True
                                        End If
                                    End If
                                End If
                            Else
                                ' intChassisType is an integer
                                If TestWin32SystemEnclosureChassisTypeIsDockingStation(intChassisType) = True Then
                                    boolInterimResult = True
                                End If
                            End If
                        End If
                    Next
                    On Error Goto 0
                    If Err Then
                        Err.Clear
                    End If
                ElseIf TestObjectIsAnyTypeOfInteger(arrChassisTypes) = True Then
                    ' arrChassisTypes is a single integer
                    boolInterimResult = TestWin32SystemEnclosureChassisTypeIsDockingStation(arrChassisTypes)
                Else
                    intFunctionReturn = intFunctionReturn + (-4 * intReturnMultiplier)
                End If
            End If
        End If
    End If

    If intFunctionReturn = 0 Then
        boolInstanceIsDockingStation = boolInterimResult
    End If
    
    TestSystemEnclosureInstanceIsDockingStation = intFunctionReturn
End Function
