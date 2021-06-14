Function TestProcessorArchitectureIs64Bit(ByRef boolProcessorArchitectureIs64Bit, ByVal strProcessorArchitecture)
    'region FunctionMetadata ####################################################
    ' Safely determines whether or not the supplied processor architecture is 64-bit.
    '
    ' This function takes two arguments:
    '   The first argument (boolProcessorArchitectureIs64Bit) is populated upon success with
    '       True if the processor architecture supplied in the second argument
    '       (strProcessorArchitecture) is 64-bit. It returns False if the processor
    '       architecture supplied in the second argument (strProcessorArchitecture) not 64-bit.
    '   The second argument (strProcessorArchitecture) is a string containing the processor
    '       architecture to evaluate.
    '
    ' The function returns 0 if the processor architecture was evaluated successfully; it
    ' returns a negative number if the processor architecture was not evaluated successfully.
    '
    ' Example:
    '   intReturnCode = TestProcessorArchitectureIs64Bit(boolProcessorArchitectureIs64Bit, "AMD64")
    '   If intReturnCode >= 0 Then
    '       ' Processor architecture ("AMD64") evaluated successfully
    '       ' boolProcessorArchitectureIs64Bit is: True
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
    'endregion Acknowledgements ####################################################

    'region DependsOn ####################################################
    ' TestObjectIsStringContainingData()
    'endregion DependsOn ####################################################

    Dim intFunctionReturn
    Dim boolWorkingProcessorArchitectureIs64Bit

    Err.Clear

    intFunctionReturn = 0

    If TestObjectIsStringContainingData(strProcessorArchitecture) = False Then
        intFunctionReturn = -1
    Else
        Select Case UCase(strProcessorArchitecture)
            Case "X86"
                boolWorkingProcessorArchitectureIs64Bit = False
            Case "AMD64"
                boolWorkingProcessorArchitectureIs64Bit = True
            Case "IA64"
                boolWorkingProcessorArchitectureIs64Bit = True
            Case "ARM"
                boolWorkingProcessorArchitectureIs64Bit = False
            Case "ARM64"
                boolWorkingProcessorArchitectureIs64Bit = True
            Case "ALPHA"
                boolWorkingProcessorArchitectureIs64Bit = False
            Case "ALPHA64"
                boolWorkingProcessorArchitectureIs64Bit = True
            Case "MIPS"
                boolWorkingProcessorArchitectureIs64Bit = False
            Case "PPC"
                boolWorkingProcessorArchitectureIs64Bit = False
            Case Else
                intFunctionReturn = -2
        End Select
    End If

    If intFunctionReturn = 0 Then
        boolProcessorArchitectureIs64Bit = boolWorkingProcessorArchitectureIs64Bit
    End If

    TestProcessorArchitectureIs64Bit = intFunctionReturn
End Function
