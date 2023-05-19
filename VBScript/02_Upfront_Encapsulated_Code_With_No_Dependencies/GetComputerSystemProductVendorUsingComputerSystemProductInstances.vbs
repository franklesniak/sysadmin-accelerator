Function GetComputerSystemProductVendorUsingComputerSystemProductInstances(ByRef strComputerSystemProductVendor, ByVal arrComputerSystemProductInstances)
    'region FunctionMetadata #######################################################
    ' Assuming that arrComputerSystemProductInstances represents an array / collection
    ' of the available computer system product instances (of type
    ' Win32_ComputerSystemProduct), this function obtains the computer system product
    ' vendor.
    '
    ' The function takes two positional arguments:
    '  - The first argument (strComputerSystemProductVendor) is populated upon success
    '    with a string containing the computer system product vendor as reported by
    '    WMI.
    '  - The second argument (arrComputerSystemProductInstances) is an array/collection
    '    of objects of class Win32_ComputerSystemProduct
    '
    ' The function returns a 0 if the computer system product vendor was obtained
    ' successfully. It returns a negative integer if an error occurred retrieving the
    ' computer system product vendor. Finally, it returns a positive integer if the
    ' computer system product vendor was obtained, but multiple computer system product
    ' instances were present that contained data for the computer system product
    ' vendor. When this happens, only the first Win32_ComputerSystemProduct instance
    ' containing data for the computer system product vendor is used.
    '
    ' Example:
    '   intReturnCode = GetComputerSystemProductInstances(arrComputerSystemProductInstances)
    '   If intReturnCode >= 0 Then
    '       ' At least one Win32_ComputerSystemProduct instance was retrieved
    '       ' successfully
    '       intReturnCode = GetComputerSystemProductVendorUsingComputerSystemProductInstances(strComputerSystemProductVendor, arrComputerSystemProductInstances)
    '       If intReturnCode >= 0 Then
    '           ' The computer system product vendor was retrieved successfully and is
    '           ' stored in strComputerSystemProductVendor
    '       End If
    '   End If
    '
    ' Version: 1.1.20230518.0
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
    ' None!
    'endregion Acknowledgements #######################################################

    'region DependsOn ##############################################################
    ' TestObjectForData()
    ' TestObjectIsAnyTypeOfInteger()
    'endregion DependsOn ##############################################################

    Dim intFunctionReturn
    Dim intReturnMultiplier
    Dim intTemp
    Dim strInterimResult
    Dim objComputerSystemProductInstance
    Dim strOldInterimResult
    Dim strResultToReturn
    Dim intCountOfComputerSystemProductVendors

    Err.Clear

    intFunctionReturn = 0
    intReturnMultiplier = 128
    strInterimResult = ""
    strResultToReturn = ""
    intCountOfComputerSystemProductVendors = 0

    If TestObjectForData(arrComputerSystemProductInstances) <> True Then
        intFunctionReturn = intFunctionReturn + (-1 * intReturnMultiplier)
    Else
        On Error Resume Next
        intTemp = arrComputerSystemProductInstances.Count
        If Err Then
            On Error Goto 0
            Err.Clear
            intFunctionReturn = intFunctionReturn + (-2 * intReturnMultiplier)
        Else
            On Error Goto 0
            If TestObjectIsAnyTypeOfInteger(intTemp) = False Then
                intFunctionReturn = intFunctionReturn + (-3 * intReturnMultiplier)
            Else
                If intTemp < 0 Then
                    intFunctionReturn = intFunctionReturn + (-4 * intReturnMultiplier)
                ElseIf intTemp = 0 Then
                    intFunctionReturn = intFunctionReturn + (-5 * intReturnMultiplier)
                Else
                    On Error Resume Next
                    For Each objComputerSystemProductInstance in arrComputerSystemProductInstances
                        If Err Then
                            Err.Clear
                        Else
                            strOldInterimResult = strInterimResult
                            strInterimResult = objComputerSystemProductInstance.Vendor
                            If Err Then
                                Err.Clear
                                strInterimResult = strOldInterimResult
                            Else
                                If TestObjectForData(strInterimResult) <> True Then
                                    strInterimResult = strOldInterimResult
                                Else
                                    ' Found a result with real data
                                    If TestObjectForData(strResultToReturn) = False Then
                                        strResultToReturn = strInterimResult
                                    End If
                                    intCountOfComputerSystemProductVendors = intCountOfComputerSystemProductVendors + 1
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

    If intFunctionReturn >= 0 Then
        ' No error has occurred yet
        If intCountOfComputerSystemProductVendors = 0 Then
            ' No result found
            intFunctionReturn = intFunctionReturn + (-5 * intReturnMultiplier)
        Else
            intFunctionReturn = intCountOfComputerSystemProductVendors - 1
        End If
    End If

    If intFunctionReturn >= 0 Then
        strComputerSystemProductVendor = strResultToReturn
    End If
    
    GetComputerSystemProductVendorUsingComputerSystemProductInstances = intFunctionReturn
End Function
