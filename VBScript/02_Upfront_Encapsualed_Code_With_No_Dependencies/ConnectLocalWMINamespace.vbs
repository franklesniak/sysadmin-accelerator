Function ConnectLocalWMINamespace(ByRef objSWbemServicesWMINamespace, ByVal strTargetWMINamespace, ByVal objSWbemNamedValueSetContextOrIntTargetWMIProviderArchitectureBitWidth)
    'region FunctionMetadata ####################################################
    ' Safely creates a SWbemServices object with a connection to the specified WMI namespace on
    ' the local computer.
    '
    ' Function takes three positional arguments:
    '   The first argument (objSWbemServicesWMINamespace) will be populated with the
    '       SWbemServices (WMI connection) object upon successful connection.
    '   The second argument (strTargetWMINamespace) specifies the namespace target to which
    '       this function will connect. If vbNullString ("") or Null is passed, the function
    '       defaults to "root\cimv2", which is the most commonly-used WMI namespace.
    '   The third argument
    '       (objSWbemNamedValueSetContextOrIntTargetWMIProviderArchitectureBitWidth) specifies
    '       either a SWbemNamedValueSet that sets the required bit-width to use when opening
    '       the WMI connection, **or** it specifies an integer target bit width "context" to
    '       use when opening WMI. For example, supplying 32 or 64 will force a respective 32-
    '       or 64-bit context when opening the WMI connection. Generally, when using this
    '       function, it is recommended to use SWbemNamedValueSet instead of an integer. This
    '       feature is commonly used when connecting to the "root\default" WMI namespace and
    '       then using the StdRegProv class to connect to the Windows registry. If Null is
    '       passed, the function defaults to the context supplied by the VBScript process that
    '       is running this script.
    '
    ' The function returns 0 if the SWbemServices (WMI connection) object
    '       objSWbemServicesWMINamespace was created successfully; a negative number otherwise.
    '
    ' Example 1:
    '   intReturnCode = ConnectLocalWMINamespace(objWMI, Null, Null)
    '   If intReturnCode = 0 Then
    '       ' objWMI is initialized and connected to the root\CIMv2 namespace
    '       Set colOS = objWMI.InstancesOf("Win32_OperatingSystem")
    '       For Each objOS in colOS
    '           WScript.Echo(objOS.Caption)
    '       Next
    '   End If
    '
    ' Example 2:
    '   Const HKEY_CLASSES_ROOT     = &H80000000
    '   Const HKEY_CURRENT_USER     = &H80000001
    '   Const HKEY_LOCAL_MACHINE    = &H80000002
    '   Const HKEY_USERS            = &H80000003
    '   intReturnCode = NewWMIBitWidthContext(objWMIContext, 32)
    '   If intReturnCode = 0 Then
    '       intReturnCode = ConnectLocalWMINamespace(objWMI, "root\default", objWMIContext)
    '       If intReturnCode = 0 Then
    '           ' objWMI is initialized and connected to the root\default namespace
    '           ' Create the StdRegProv:
    '           Set objStdRegProv = objWMI.Get("StdRegProv")
    '           ' Create a registry key in the 32-bit process context:
    '           Set objInParams = objStdRegProv.Methods_("CreateKey").Inparameters
    '           objInParams.hDefKey = HKEY_CURRENT_USER
    '           objInParams.sSubKeyName = "SOFTWARE\West Monroe Partners\Temp"
    '           Set objOutParams = objStdRegProv.ExecMethod_("CreateKey",objInParams,,objWMIContext)
    '           intReturnCode = objOutParams.ReturnValue
    '       End If
    '   End If
    '
    ' Example 3:
    '   intReturnCode = ConnectLocalWMINamespace(objWMI, Null, 64)
    '   If intReturnCode = 0 Then
    '       ' objWMI is initialized and connected to the root\cimv2 namespace
    '       Set colWinSATs = objWMI.ExecQuery("Select * From Win32_WinSAT")
    '       For Each objWinSAT in colWinSATs
    '           WScript.Echo(objWinSAT.WinSATAssessmentState)
    '       Next
    '   End If
    '
    ' Version: 2.2.20210613.0
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
    ' at https://github.com/franklesniak/VBScript_Resources
    'endregion DownloadLocationNotice ####################################################

    'region DependsOn ####################################################
    ' TestObjectForData()
    ' TestObjectIsAnyTypeOfInteger()
    ' NewWMIBitWidthContext() <- not a strict dependency, but connecting to an alternative
    '                            bit-width context requires this function
    'endregion DependsOn ####################################################

    Dim strEffectiveComputerName
    Dim intReturnCode
    Dim strEffectiveNamespace
    Dim objSWbemLocator
    Dim objSWbemNamedValueSetContext
    Dim objSWbemServicesTemp

    Const wbemImpersonationLevelImpersonate = 3
    strEffectiveComputerName = "."

    Err.Clear

    intReturnCode = 0
    
    If TestObjectForData(strTargetWMINamespace) = False Then
        strEffectiveNamespace = "root\cimv2"
    Else
        strEffectiveNamespace = strTargetWMINamespace
    End If

    On Error Resume Next
    Set objSWbemLocator = CreateObject("Wbemscripting.SWbemLocator")
    If Err Then
        On Error Goto 0
        Err.Clear
        intReturnCode = -1
    Else
        On Error Goto 0
    End If

    If intReturnCode = 0 Then
        ' No error occurred
        If TestObjectForData(objSWbemNamedValueSetContextOrIntTargetWMIProviderArchitectureBitWidth) = True Then
            ' objSWbemNamedValueSetContextOrIntTargetWMIProviderArchitectureBitWidth parameter
            ' was supplied
            If TestObjectIsAnyTypeOfInteger(objSWbemNamedValueSetContextOrIntTargetWMIProviderArchitectureBitWidth) = True Then
                ' objSWbemNamedValueSetContextOrIntTargetWMIProviderArchitectureBitWidth is an
                ' integer
                On Error Resume Next
                Set objSWbemNamedValueSetContext = CreateObject("WbemScripting.SWbemNamedValueSet")
                If Err Then
                    On Error Goto 0
                    Err.Clear
                    intReturnCode = -2
                Else
                    objSWbemNamedValueSetContext.Add "__ProviderArchitecture", objSWbemNamedValueSetContextOrIntTargetWMIProviderArchitectureBitWidth
                    If Err Then
                        On Error Goto 0
                        Err.Clear
                        intReturnCode = -3
                    Else
                        objSWbemNamedValueSetContext.Add "__RequiredArchitecture", True
                        If Err Then
                            On Error Goto 0
                            Err.Clear
                            intReturnCode = -4
                        Else
                            Set objSWbemServicesTemp = objSWbemLocator.ConnectServer(strEffectiveComputerName, strEffectiveNamespace,,,,,,objSWbemNamedValueSetContext)
                            If Err Then
                                On Error Goto 0
                                Err.Clear
                                intReturnCode = -5
                            Else
                                On Error Goto 0
                            End If
                        End If
                    End If
                End If
            Else
                ' objSWbemNamedValueSetContextOrIntTargetWMIProviderArchitectureBitWidth is not
                ' an integer; it is probably a SWbemNamedValueSet
                On Error Resume Next
                Set objSWbemServicesTemp = objSWbemLocator.ConnectServer(strEffectiveComputerName, strEffectiveNamespace,,,,,,objSWbemNamedValueSetContextOrIntTargetWMIProviderArchitectureBitWidth)
                If Err Then
                    On Error Goto 0
                    Err.Clear
                    intReturnCode = -6
                Else
                    On Error Goto 0
                End If
            End If
        Else
            ' objSWbemNamedValueSetContextOrIntTargetWMIProviderArchitectureBitWidth parameter
            ' was not supplied
            On Error Resume Next
            Set objSWbemServicesTemp = objSWbemLocator.ConnectServer(strEffectiveComputerName, strEffectiveNamespace)
            If Err Then
                On Error Goto 0
                Err.Clear
                intReturnCode = -7
            Else
                On Error Goto 0
            End If
        End If

        If intReturnCode = 0 Then
            ' No error occurred
            On Error Resume Next
            objSWbemServicesTemp.Security_.ImpersonationLevel = wbemImpersonationLevelImpersonate
            If Err Then
                On Error Goto 0
                Err.Clear
                intReturnCode = -8
            Else
                On Error Goto 0
            End If
        End If
    End If

    If intReturnCode = 0 Then
        ' No error occurred
        ' We fully connected to WMI, but did so with a "dummy" object...
        ' ... so, let's connect using the real object
        Set objSWbemServicesTemp = Nothing
        If TestObjectForData(objSWbemNamedValueSetContextOrIntTargetWMIProviderArchitectureBitWidth) = True Then
            ' objSWbemNamedValueSetContextOrIntTargetWMIProviderArchitectureBitWidth parameter
            ' was supplied
            If TestObjectIsAnyTypeOfInteger(objSWbemNamedValueSetContextOrIntTargetWMIProviderArchitectureBitWidth) = True Then
                ' objSWbemNamedValueSetContextOrIntTargetWMIProviderArchitectureBitWidth is an
                ' integer
                ' objSWbemNamedValueSetContext already constructed
                On Error Resume Next
                Set objSWbemServicesWMINamespace = objSWbemLocator.ConnectServer(strEffectiveComputerName, strEffectiveNamespace,,,,,,objSWbemNamedValueSetContext)
                If Err Then
                    On Error Goto 0
                    Err.Clear
                    intReturnCode = -9
                Else
                    On Error Goto 0
                End If
            Else
                ' objSWbemNamedValueSetContextOrIntTargetWMIProviderArchitectureBitWidth is not
                ' an integer; it is probably a SWbemNamedValueSet
                On Error Resume Next
                Set objSWbemServicesWMINamespace = objSWbemLocator.ConnectServer(strEffectiveComputerName, strEffectiveNamespace,,,,,,objSWbemNamedValueSetContextOrIntTargetWMIProviderArchitectureBitWidth)
                If Err Then
                    On Error Goto 0
                    Err.Clear
                    intReturnCode = -10
                Else
                    On Error Goto 0
                End If
            End If
        Else
            ' objSWbemNamedValueSetContextOrIntTargetWMIProviderArchitectureBitWidth parameter
            ' was not supplied
            On Error Resume Next
            Set objSWbemServicesWMINamespace = objSWbemLocator.ConnectServer(strEffectiveComputerName, strEffectiveNamespace)
            If Err Then
                On Error Goto 0
                Err.Clear
                intReturnCode = -11
            Else
                On Error Goto 0
            End If
        End If
        If intReturnCode = 0 Then
            ' No error occurred
            On Error Resume Next
            objSWbemServicesWMINamespace.Security_.ImpersonationLevel = wbemImpersonationLevelImpersonate
            If Err Then
                On Error Goto 0
                Err.Clear
                intReturnCode = -12
            Else
                On Error Goto 0
            End If
        End If
    End If

    ConnectLocalWMINamespace = intReturnCode
End Function
