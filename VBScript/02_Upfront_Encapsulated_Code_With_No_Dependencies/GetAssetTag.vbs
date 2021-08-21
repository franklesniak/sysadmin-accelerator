Function GetAssetTag(ByRef strAssetTag)
    'region FunctionMetadata ####################################################
    ' This function obtains the computer's asset tag, if available and configured by the
    ' computer's manufacturer or administrator/owner
    '
    ' The function takes one positional argument (strAssetTag), which is populated upon success
    ' with a string containing the computer's asset tag as reported by the System Management
    ' BIOS (SMBIOS), via WMI.
    '  - The second argument (arrSystemEnclosureInstances) is an array/collection of objects of
    '    class Win32_SystemEnclosure
    '
    ' The function returns a 0 if the asset tag was obtained successfully. It returns a
    ' negative integer if an error occurred retrieving the asset tag. Finally, it returns a
    ' positive integer if the asset tag was obtained, but multiple SystemEnclosure instances
    ' were present that contained data for the asset tag. When this happens, only the first
    ' Win32_SystemEnclosure instance containing data for the asset tag is used.
    '
    ' Example:
    '   intReturnCode = GetAssetTag(strAssetTag)
    '   If intReturnCode >= 0 Then
    '      ' The computer's asset tag was retrieved successfully and is stored in strAssetTag
    '   End If
    '
    ' Version: 1.0.20210812.0
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
    ' GetSystemEnclosureInstances()
    ' GetAssetTagUsingSystemEnclosureInstances()
    'endregion DependsOn ####################################################

    Dim intFunctionReturn
    Dim arrSystemEnclosureInstances
    Dim strResult

    intFunctionReturn = 0

    intFunctionReturn = GetSystemEnclosureInstances(arrSystemEnclosureInstances)
    If intFunctionReturn >= 0 Then
        ' At least one Win32_SystemEnclosure instance was retrieved successfully
        intFunctionReturn = GetAssetTagUsingSystemEnclosureInstances(strResult, arrSystemEnclosureInstances)
        If intFunctionReturn >= 0 Then
            ' The computer's asset tag was retrieved successfully and is stored in strResult
            strAssetTag = strResult
        End If
    End If
    
    GetAssetTag = intFunctionReturn
End Function
