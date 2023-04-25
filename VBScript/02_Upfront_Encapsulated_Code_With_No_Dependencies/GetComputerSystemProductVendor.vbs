Function GetComputerSystemProductVendor(ByRef strComputerSystemProductVendor)
    'region FunctionMetadata #######################################################
    ' This function obtains the computer system product vendor.
    '
    ' The function takes one positional argument (strComputerSystemProductVendor),
    ' which is populated upon success with a string containing the computer system's
    ' product vendor as reported by WMI.
    '
    ' The function returns a 0 if the computer system product vendor was obtained
    ' successfully. It returns a negative integer if an error occurred retrieving the
    ' computer system product vendor. Finally, it returns a positive integer if the
    ' computer system product vendor was obtained, but multiple computer system
    ' product instances were present that contained data for the computer system
    ' product vendor. When this happens, only the first
    ' Win32_ComputerSystemProductProduct instance containing data for the computer
    ' system product vendor is used.
    '
    ' Example:
    '   intReturnCode = GetComputerSystemProductVendor(strComputerSystemProductVendor)
    '   If intReturnCode >= 0 Then
    '       '  The computer system product vendor was retrieved successfully and is
    '       ' stored in strComputerSystemProductVendor
    '   End If
    '
    ' Version: 1.0.20230424.0
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

    'region DependsOn ####################################################
    ' GetComputerSystemProductInstances()
    ' GetComputerSystemProductVendorUsingComputerSystemProductInstances()
    'endregion DependsOn ####################################################

    Dim intFunctionReturn
    Dim arrComputerSystemProductInstances
    Dim strResult

    intFunctionReturn = 0

    intFunctionReturn = GetComputerSystemProductInstances(arrComputerSystemProductInstances)
    If intFunctionReturn >= 0 Then
        ' At least one Win32_ComputerSystemProduct instance was retrieved successfully
        intFunctionReturn = GetComputerSystemProductVendorUsingComputerSystemProductInstances(strResult, arrComputerSystemProductInstances)
        If intFunctionReturn >= 0 Then
            ' The computer system product vendor was retrieved successfully and is
            ' stored in strResult
            strComputerSystemProductVendor = strResult
        End If
    End If
    
    GetComputerSystemProductVendor = intFunctionReturn
End Function
