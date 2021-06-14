Function GetExecutableProcessorArchitectureFromFile(ByRef strProcessorArchitecture, ByVal strPathToFile)
    'region FunctionMetadata ####################################################
    ' Executable files have a header that tells the operating system what "machine type" or
    ' processor architecture the executable targets. This function reads the executable's
    ' header and determines the target processor architecture (i.e., the value obtained from
    ' the environment variable "PROCESSOR_ARCHITECTURE", from the perspective of the executable
    ' if this executable were running).
    '
    ' This function takes two arguments:
    '   - The first argument (strProcessorArchitecture) is populated upon success with the
    '       processor architecture of the executable, i.e., the value that the executable would
    '       see in the environment variable "PROCESSOR_ARCHITECTURE" if it were running.
    '   - The second argument (strPathToFile) contains the full path of the executable to
    '       evaluate.
    '
    ' The function returns 0 if the executable's processor architecture (machine type) was
    ' determined succesfully. It returns a negative number if it were not determined
    ' successfully
    '
    ' Example:
    '   intReturnCode = GetExecutableProcessorArchitectureFromFile(strProcessorArchitecture, "C:\Windows\explorer.exe")
    '   If intReturnCode = 0 Then
    '       ' The header of explorer.exe was read and its processor architecture was
    '       ' successfully determined. strProcessorArchitecture is populated, e.g., as "AMD64".
    '   End If
    '
    ' The known processor architectures are as follows:
    '
    ' x86 = Intel IA32 (32-bit x86 or compatible), 32-bit
    ' AMD64 = AMD64, Intel x86-64 (Intel x64), or compatible, 64-bit
    ' IA64 = Intel Itanium, 64-bit (Windows XP, Windows Server 2003, and Windows Server 2008
    '        only)
    ' ARM = ARM (Native ARM operating systems include Windows 8 RT, Windows 8.1 RT, and Windows
    '       10 Mobile/IoT Core only*; however, newer ARM64 releases of Windows 10 can run ARM
    '       processes), 32-bit
    ' ARM64 = ARM64, (Windows 10 and newer only**), 64-bit
    ' ALPHA = Alpha/DEC (Windows NT4 family, only), 32-bit
    ' ALPHA64 = Alpha/DEC (Windows 2000 pre-release versions, only), 64-bit
    ' MIPS = MIPS (Windows NT 3.51 / 4.0 families, only), 32-bit
    ' PPC = PowerPC (Windows NT4 family, only), 32-bit
    '
    ' *  = Windows CE / Windows Mobile also had support for ARM. However, those operating
    '      systems did not include support for VBScript to the knowledge of the author
    ' ** = At the time of writing, Microsoft is rumored to be working on an ARM version of
    '      Windows Server
    '
    ' Version: 1.0.20210614.0
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

    'region Acknowledgements ####################################################
    ' Microsoft, for providing a current reference on the SYSTEM_INFO struct, used by the
    ' GetSystemInfo Win32 function. This reference does not show the exact text of the
    ' PROCESSOR_ARCHITECTURE environment variable, but shows the universe of what's possible on
    ' a core system API:
    ' https://docs.microsoft.com/en-us/windows/win32/api/sysinfoapi/ns-sysinfoapi-system_info#members
    '
    ' Microsoft, for including in the MSDN Library Jan 2003 information on this same SYSTEM_INFO
    ' struct that pre-dates Windows 2000 and enumerates additional processor architectures
    ' (MIPS, ALPHA, PowerPC, IA32_ON_WIN64). The MSDN Library Jan 2003 also lists SHX and ARM,
    ' explains nuiances in accessing environment variables on pre-Windows 2000 operating
    ' systems (namely that VBScript in Windows 9x can only access per-process environment
    ' variables), and that the PROCESSOR_ARCHITECTURE system environment variable is not
    ' available on Windows 98/ME.
    ' (link unavailable, check Internet Archive for source)
    '
    ' Adam Haile, for confirming that there is no VBScript support for Windows CE/Mobile:
    ' https://stackoverflow.com/a/28838/2134110
    '
    ' Wikipedia for listing the operating systems that included Windows Scripting Host support:
    ' https://en.wikipedia.org/wiki/Windows_Script_Host#Version_history
    '
    ' "guga" for the first post in this thread that tipped me off to the SYSTEM_INFO struct and
    ' additional architectures:
    ' http://masm32.com/board/index.php?topic=3401.0
    '
    ' Ron Loveless and Andrew C. Wilson for authoring this guide that confirmed the values for
    ' PROCESSOR_ARCHITECTURE for MIPS, Alpha, and PowerPC architecture:
    ' http://www-personal.umich.edu/~acwilson/unattend-nt/tech-doc-draft-acwilson.html
    '
    ' IBM, for publishing "Windows NT Systems Management", which provided a second confirmation
    ' of the MIPS, Alpha, and PowerPC architecture values for PROCESS_ARCHITECTURE:
    ' https://www.infania.net/misc/basil.holloway/ALL%20PDF/sg242107.pdf
    '
    ' Microsoft, for publishing the reference to the PE header machine types that translates
    ' their integer values to their respective architecture.
    ' https://docs.microsoft.com/en-us/windows/win32/debug/pe-format?redirectedfrom=MSDN#machine-types
    '
    ' StackExchange user "jumpjack", for giving me a pointer on how to read the PE header in
    ' VBScript:
    ' https://superuser.com/a/1027161/334370
    'endregion Acknowledgements ####################################################

    'region DependsOn ####################################################
    ' ReadCharactersFromFile()
    ' TestObjectForData()
    ' ConvertPortableExecutableMachineTypeToProcessorArchitecture()
    'endregion DependsOn ####################################################

    Dim intFunctionReturn
    Dim intReturnMultiplier
    Dim intReturnCode
    Dim strData
    Dim intLength
    Dim strCharacter
    Dim intLowerByte
    Dim intUpperByte
    Dim intPEHeaderOffset
    Dim intPECharacter
    Dim intMachineType
    Dim strWorkingProcessorArchitecture

    Const PE_HEADER_FILE_OFFSET_POSITION = &H3C

    Err.Clear

    intFunctionReturn = 0
    intReturnMultiplier = 1

    intReturnCode = ReadCharactersFromFile(strData, strPathToFile, PE_HEADER_FILE_OFFSET_POSITION + 2, Null)
    If intReturnCode < 0 Then
        intFunctionReturn = intFunctionReturn + (intReturnCode * intReturnMultiplier)
    Else
        intReturnMultiplier = intReturnMultiplier * 16
        If TestObjectForData(strData) = False Then
            intFunctionReturn = intFunctionReturn + (-1 * intReturnMultiplier)
        Else
            On Error Resume Next
            intLength = Len(strData)
            If Err Then
                On Error Goto 0
                Err.Clear
                intFunctionReturn = intFunctionReturn + (-2 * intReturnMultiplier)
            Else
                On Error Goto 0
                If intLength < (PE_HEADER_FILE_OFFSET_POSITION + 2) Then
                    intFunctionReturn = intFunctionReturn + (-3 * intReturnMultiplier)
                Else
                    ' The file offset for the PE header is specified at index &H3C (60), with zero-indexing
                    ' Since Mid() is 1-indexed, we start with 61:
                    On Error Resume Next
                    strCharacter = Mid(strData, PE_HEADER_FILE_OFFSET_POSITION + 1, 1)
                    If Err Then
                        On Error Goto 0
                        Err.Clear
                        intFunctionReturn = intFunctionReturn + (-4 * intReturnMultiplier)
                    Else
                        intLowerByte = Asc(strCharacter)
                        If Err Then
                            On Error Goto 0
                            Err.Clear
                            intFunctionReturn = intFunctionReturn + (-5 * intReturnMultiplier)
                        Else
                            strCharacter = Mid(strData, PE_HEADER_FILE_OFFSET_POSITION + 1 + 1, 1)
                            If Err Then
                                On Error Goto 0
                                Err.Clear
                                intFunctionReturn = intFunctionReturn + (-6 * intReturnMultiplier)
                            Else
                                intUpperByte = Asc(strCharacter)
                                If Err Then
                                    On Error Goto 0
                                    Err.Clear
                                    intFunctionReturn = intFunctionReturn + (-7 * intReturnMultiplier)
                                Else
                                    On Error Goto 0
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If

    If intFunctionReturn = 0 Then
        ' intLowerByte has the lower byte of the PE offset address
        ' intUpperByte has the upper byte of the PE offset address
        intPEHeaderOffset = (256 * intUpperByte) + intLowerByte
        intReturnMultiplier = intReturnMultiplier * 8
        intReturnCode = ReadCharactersFromFile(strData, strPathToFile, intPEHeaderOffset + 6, Null)
        If intReturnCode < 0 Then
            intFunctionReturn = intFunctionReturn + (intReturnCode * intReturnMultiplier)
        Else
            intReturnMultiplier = intReturnMultiplier * 16
            If TestObjectForData(strData) = False Then
                intFunctionReturn = intFunctionReturn + (-1 * intReturnMultiplier)
            Else
                On Error Resume Next
                intLength = Len(strData)
                If Err Then
                    On Error Goto 0
                    Err.Clear
                    intFunctionReturn = intFunctionReturn + (-2 * intReturnMultiplier)
                Else
                    On Error Goto 0
                    If intLength < (intPEHeaderOffset + 6) Then
                        intFunctionReturn = intFunctionReturn + (-3 * intReturnMultiplier)
                    Else
                        On Error Resume Next
                        strCharacter = Mid(strData, intPEHeaderOffset + 1, 1)
                        If Err Then
                            On Error Goto 0
                            Err.Clear
                            intFunctionReturn = intFunctionReturn + (-4 * intReturnMultiplier)
                        Else
                            intPECharacter = Asc(strCharacter)
                            If Err Then
                                On Error Goto 0
                                Err.Clear
                                intFunctionReturn = intFunctionReturn + (-5 * intReturnMultiplier)
                            Else
                                On Error Goto 0
                                If intPECharacter <> Asc("P") Then
                                    intFunctionReturn = intFunctionReturn + (-6 * intReturnMultiplier)
                                Else
                                    On Error Resume Next
                                    strCharacter = Mid(strData, intPEHeaderOffset + 1 + 1, 1)
                                    If Err Then
                                        On Error Goto 0
                                        Err.Clear
                                        intFunctionReturn = intFunctionReturn + (-7 * intReturnMultiplier)
                                    Else
                                        intPECharacter = Asc(strCharacter)
                                        If Err Then
                                            On Error Goto 0
                                            Err.Clear
                                            intFunctionReturn = intFunctionReturn + (-8 * intReturnMultiplier)
                                        Else
                                            On Error Goto 0
                                            If intPECharacter <> Asc("E") Then
                                                intFunctionReturn = intFunctionReturn + (-9 * intReturnMultiplier)
                                            Else
                                                On Error Resume Next
                                                strCharacter = Mid(strData, intPEHeaderOffset + 1 + 2, 1)
                                                If Err Then
                                                    On Error Goto 0
                                                    Err.Clear
                                                    intFunctionReturn = intFunctionReturn + (-10 * intReturnMultiplier)
                                                Else
                                                    intPECharacter = Asc(strCharacter)
                                                    If Err Then
                                                        On Error Goto 0
                                                        Err.Clear
                                                        intFunctionReturn = intFunctionReturn + (-11 * intReturnMultiplier)
                                                    Else
                                                        On Error Goto 0
                                                        If intPECharacter <> 0 Then
                                                            intFunctionReturn = intFunctionReturn + (-12 * intReturnMultiplier)
                                                        Else
                                                            On Error Resume Next
                                                            strCharacter = Mid(strData, intPEHeaderOffset + 1 + 3, 1)
                                                            If Err Then
                                                                On Error Goto 0
                                                                Err.Clear
                                                                intFunctionReturn = intFunctionReturn + (-13 * intReturnMultiplier)
                                                            Else
                                                                intPECharacter = Asc(strCharacter)
                                                                If Err Then
                                                                    On Error Goto 0
                                                                    Err.Clear
                                                                    intFunctionReturn = intFunctionReturn + (-14 * intReturnMultiplier)
                                                                Else
                                                                    On Error Goto 0
                                                                    If intPECharacter <> 0 Then
                                                                        intFunctionReturn = intFunctionReturn + (-15 * intReturnMultiplier)
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If

    If intFunctionReturn = 0 Then
        ' intPEHeaderOffset confirmed to be PE header
        On Error Resume Next
        strCharacter = Mid(strData, intPEHeaderOffset + 1 + 4, 1)
        If Err Then
            On Error Goto 0
            Err.Clear
            intFunctionReturn = intFunctionReturn + (-16 * intReturnMultiplier)
        Else
            intLowerByte = Asc(strCharacter)
            If Err Then
                On Error Goto 0
                Err.Clear
                intFunctionReturn = intFunctionReturn + (-17 * intReturnMultiplier)
            Else
                strCharacter = Mid(strData, intPEHeaderOffset + 1 + 5, 1)
                If Err Then
                    On Error Goto 0
                    Err.Clear
                    intFunctionReturn = intFunctionReturn + (-18 * intReturnMultiplier)
                Else
                    intUpperByte = Asc(strCharacter)
                    If Err Then
                        On Error Goto 0
                        Err.Clear
                        intFunctionReturn = intFunctionReturn + (-19 * intReturnMultiplier)
                    Else
                        On Error Goto 0
                    End If
                End If
            End If
        End If
    End If

    If intFunctionReturn = 0 Then
        ' intLowerByte contains lower byte of PE header machine type
        ' intUpperByte contains upper byte of PE header machine type
        intMachineType = (256 * intUpperByte) + intLowerByte
        intReturnMultiplier = intReturnMultiplier * 32
        intReturnCode = ConvertPortableExecutableMachineTypeToProcessorArchitecture(strWorkingProcessorArchitecture, intMachineType)
        If intReturnCode < 0 Then
            intFunctionReturn = intFunctionReturn + (intReturnCode * intReturnMultiplier)
        Else
            If TestObjectForData(strWorkingProcessorArchitecture) = False Then
                intReturnMultiplier = intReturnMultiplier * 16
                intFunctionReturn = intFunctionReturn + (-1 * intReturnMultiplier)
            End If
        End If
    End If

    If intFunctionReturn = 0 Then
        strProcessorArchitecture = strWorkingProcessorArchitecture
    End If

    GetExecutableProcessorArchitectureFromFile = intFunctionReturn
End Function
