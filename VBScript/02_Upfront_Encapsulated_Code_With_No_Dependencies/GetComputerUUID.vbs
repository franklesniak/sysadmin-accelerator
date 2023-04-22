Function GetComputerUUID(ByRef strComputerUUID)
    'region FunctionMetadata #######################################################
    ' This function obtains the computer's universally unique ID (UUID), which is a
    ' 128-bit identifier represented as a string in GUID format
    ' (AAAAAAAA-BBBB-CCCC-DDDD-EEEEEEEEEEEE)
    '
    ' The function takes one positional argument (strComputerUUID), which is populated
    ' upon success with a string containing the computer's UUID as reported by WMI.
    '
    ' The function returns a 0 if the computer's UUID was obtained successfully. It
    ' returns a negative integer if an error occurred retrieving the UUID. Finally, it
    ' returns a positive integer if the computer's UUID was obtained, but multiple
    ' computer system product instances were present that contained data for the UUID.
    ' When this happens, only the first Win32_ComputerSystemProduct instance containing
    ' data for the UUID is used.
    '
    ' Example:
    '   intReturnCode = GetComputerUUID(strComputerUUID)
    '   If intReturnCode >= 0 Then
    '       ' The computer's UUID was retrieved successfully and is stored in
    '       ' strComputerUUID
    '   End If
    '
    ' Version: 1.0.20230422.0
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
    ' The most up-to-date version of this script can be found on the author's GitHub repository
    ' at https://github.com/franklesniak/sysadmin-accelerator
    'endregion DownloadLocationNotice #################################################

    'region Acknowledgements #######################################################
    ' None!
    'endregion Acknowledgements #######################################################

    'region DependsOn ##############################################################
    ' GetComputerSystemProductInstances()
    ' GetComputerUUIDUsingComputerSystemProductInstances()
    'endregion DependsOn ##############################################################

    Dim intFunctionReturn
    Dim arrComputerSystemProductInstances
    Dim strResult

    intFunctionReturn = 0

    intFunctionReturn = GetComputerSystemProductInstances(arrComputerSystemProductInstances)
    If intFunctionReturn >= 0 Then
        ' At least one Win32_ComputerSystemProduct instance was retrieved successfully
        intFunctionReturn = GetComputerUUIDUsingComputerSystemProductInstances(strResult, arrComputerSystemProductInstances)
        If intFunctionReturn >= 0 Then
            ' The computer UUID was retrieved successfully and is stored in strResult
            strComputerUUID = strResult
        End If
    End If
    
    GetComputerUUID = intFunctionReturn
End Function
