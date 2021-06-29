Function TestComputerIsPortablePhysicalComputer(ByRef boolIsPortablePhysicalComputer)
    'region FunctionMetadata ####################################################
    ' This function determines if the computer is a portable physical computer (e.g., a laptop
    ' or tablet)
    '
    ' The function takes one positional argument (boolIsPortablePhysicalComputer), which is
    ' populated upon success with a boolean value: True when the computer was determined to be
    ' a portable physical computer (i.e., a laptop or tablet), False otherwise
    '
    ' The function returns a 0 when the function successfully evaluated whether the computer is
    ' a portable physical computer (i.e., a laptop or tablet). The function returns a negative
    ' integer if an error occurred.
    '
    ' Example:
    '   intReturnCode = TestComputerIsPortablePhysicalComputer(boolIsPortablePhysicalComputer)
    '   If intReturnCode = 0 Then
    '       ' Successfully tested whether this system is a portable physical computer
    '       If boolIsPortablePhysicalComputer = True Then
    '           ' Computer is a portable physical computer (e.g., laptop or tablet system)
    '       Else
    '           ' Computer is not a portable physical computer, i.e., it is a stationary,
    '           ' non-server computer (i.e., a desktop), a physical server chassis, or it is a
    '           ' virtual machine
    '       End If
    '   End If
    '
    ' Version: 1.0.20210629.0
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
    ' ConnectLocalWMINamespace()
    ' GetComputerSystemInstancesUsingWMINamespace()
    ' GetSystemEnclosureInstancesUsingWMINamespace()
    ' TestComputerIsPortablePhysicalComputerUsingComputerSystemAndSystemEnclosureInstances()
    'endregion DependsOn ####################################################

    Dim intFunctionReturn
    Dim intReturnMultiplier
    Dim intOffset
    Dim boolInterimResult
    Dim intReturnCode
    Dim objSWbemServicesWMINamespace
    Dim arrComputerSystemInstances
    Dim arrSystemEnclosureInstances

    intFunctionReturn = 0
    intReturnMultiplier = 1
    intOffset = 1073741824

    boolInterimResult = False

    intReturnCode = ConnectLocalWMINamespace(objSWbemServicesWMINamespace, Null, Null)
    If intReturnCode <> 0 Then
        ' Error occurred
        intFunctionReturn = intFunctionReturn + (intReturnCode * intReturnMultiplier) + intOffset
    Else
        ' intReturnCode = 0
        intReturnMultiplier = intReturnMultiplier * 16
        ' Successfully connected to the local computer's root\CIMv2 WMI Namespace
        intReturnCode = GetComputerSystemInstancesUsingWMINamespace(arrComputerSystemInstances, objSWbemServicesWMINamespace)
        If intReturnCode < 0 Then
            ' Error occurred
            intFunctionReturn = intFunctionReturn + (intReturnCode * intReturnMultiplier) + intOffset
        Else
            ' intReturnCode >= 0
            intReturnMultiplier = intReturnMultiplier * 8
            ' At least one Win32_ComputerSystem instance was retrieved successfully
            intReturnCode = GetSystemEnclosureInstancesUsingWMINamespace(arrSystemEnclosureInstances, objSWbemServicesWMINamespace)
            If intReturnCode < 0 Then
                ' Error occurred
                intFunctionReturn = intFunctionReturn + (intReturnCode * intReturnMultiplier) + intOffset
            Else
                ' intReturnCode >= 0
                intReturnMultiplier = 1
                intOffset = 0
                ' One or more Win32_SystemEnclosure instances were retrieved. The first
                ' instance is available at arrSystemEnclosureInstances.ItemIndex(0) and the
                ' number of instances is available at arrSystemEnclosureInstances.Count. In
                ' other words, the upper array boundary/index is
                ' (arrSystemEnclosureInstances.Count - 1).
                intReturnCode = TestComputerIsPortablePhysicalComputerUsingComputerSystemAndSystemEnclosureInstances(boolInterimResult, arrComputerSystemInstances, arrSystemEnclosureInstances)
                If intReturnCode <> 0 Then
                    ' Error occurred
                    intFunctionReturn = intFunctionReturn + (intReturnCode * intReturnMultiplier) + intOffset
                Else
                    ' intReturnCode = 0
                    ' Successfully tested whether this system is a stationary, non-server
                    ' physical computer
                End If
            End If
        End If
    End If

    If intFunctionReturn = 0 Then
        boolIsPortablePhysicalComputer = boolInterimResult
    End If
    
    TestComputerIsPortablePhysicalComputer = intFunctionReturn
End Function
