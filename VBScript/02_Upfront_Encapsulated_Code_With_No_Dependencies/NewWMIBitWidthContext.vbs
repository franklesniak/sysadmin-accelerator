Function NewWMIBitWidthContext(ByRef objSWbemNamedValueSetContext, ByVal intTargetWMIProviderArchitectureBitWidth)
    'region FunctionMetadata ####################################################
    ' Safely creates a SWbemNamedValueSet object for use with setting the bit-width "context"
    ' when connecting to or working with WMI.
    '
    ' Function takes three positional arguments:
    '   The first argument (objSWbemNamedValueSetContext) will be populated with the
    '       SWbemNamedValueSet (WMI context) object upon successful creation and configuration.
    '   The second argument (intTargetWMIProviderArchitectureBitWidth) specifies a target bit
    '       width "context" to use when opening WMI. For example, supplying 32 or 64 will force
    '       a respective 32- or 64-bit context when opening the WMI connection. This feature is
    '       commonly used when connecting to the "root\default" WMI namespace and then using
    '       the StdRegProv class to connect to the Windows registry.
    '
    ' The function returns 0 if the SWbemNamedValueSet (WMI context) object
    '       objSWbemNamedValueSetContext was created successfully; a negative number otherwise.
    '
    ' Example:
    '   intReturnCode = NewWMIBitWidthContext(objWMIContext, 32)
    '   If intReturnCode = 0 Then
    '       ' objWMIContext is initialized and configured to instruct WMI to use a 32-bit
    '       ' context.
    '   End If
    '
    ' Version: 1.1.20210613.0
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

    'region DependsOn ####################################################
    ' TestObjectIsAnyTypeOfInteger()
    'endregion DependsOn ####################################################

    Dim intReturnCode
    Dim objSWbemNamedValueSetTemp

    Err.Clear

    intReturnCode = 0

    If TestObjectIsAnyTypeOfInteger(intTargetWMIProviderArchitectureBitWidth) = False Then
        intReturnCode = -1
    Else
        On Error Resume Next
        Set objSWbemNamedValueSetTemp = CreateObject("WbemScripting.SWbemNamedValueSet")
        If Err Then
            On Error Goto 0
            Err.Clear
            intReturnCode = -2
        Else
            objSWbemNamedValueSetTemp.Add "__ProviderArchitecture", intTargetWMIProviderArchitectureBitWidth
            If Err Then
                On Error Goto 0
                Err.Clear
                intReturnCode = -3
            Else
                objSWbemNamedValueSetTemp.Add "__RequiredArchitecture", True
                If Err Then
                    On Error Goto 0
                    Err.Clear
                    intReturnCode = -4
                Else
                    On Error Goto 0
                End If
            End If
        End If
    End If

    If intReturnCode = 0 Then
        ' No error occurred
        ' At this point, we've only configured a temporary variable; we still need to configure
        ' objSWbemNamedValueSetContext:
        Set objSWbemNamedValueSetTemp = Nothing
        On Error Resume Next
        Set objSWbemNamedValueSetContext = CreateObject("WbemScripting.SWbemNamedValueSet")
        If Err Then
            On Error Goto 0
            Err.Clear
            intReturnCode = -5
        Else
            objSWbemNamedValueSetContext.Add "__ProviderArchitecture", intTargetWMIProviderArchitectureBitWidth
            If Err Then
                On Error Goto 0
                Err.Clear
                intReturnCode = -6
            Else
                objSWbemNamedValueSetContext.Add "__RequiredArchitecture", True
                If Err Then
                    On Error Goto 0
                    Err.Clear
                    intReturnCode = -7
                Else
                    On Error Goto 0
                End If
            End If
        End If
    End If

    NewWMIBitWidthContext = intReturnCode
End Function
