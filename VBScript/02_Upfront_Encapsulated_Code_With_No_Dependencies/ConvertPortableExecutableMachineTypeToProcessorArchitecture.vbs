Function ConvertPortableExecutableMachineTypeToProcessorArchitecture(ByRef strProcessorArchitecture, ByVal intMachineType)
    'region FunctionMetadata ####################################################
    ' Portable Executable (PE) files and Common Object File Format (COFF) files each have a
    ' header that contains the "machine type" or CPU type that the executable targets. This
    ' targeted machine type is specified as an integer. This function takes these integer
    ' values and converts them to the equivalent processor architecture (i.e., the value
    ' obtained from the environment variable "PROCESSOR_ARCHITECTURE", from the perspective of
    ' the executable if this executable were running).
    '
    ' This function takes two arguments:
    '   - The first argument (strProcessorArchitecture) is populated upon success with the
    '       processor architecture of the executable, i.e., the value that the executable would
    '       see in the environment variable "PROCESSOR_ARCHITECTURE" if it were running.
    '   - The second argument (intMachineType) contains the integrer representation of this
    '       executable's target "machine type", i.e., the numerical code that is part of the
    '       portable executable (PE) file format.
    '
    ' The function returns 0 if the conversion from PE file format / COFF file integer machine
    ' type -> string processor architecture happened successfully; it returns a negative number
    ' if the integer machine type could not be converted.
    '
    ' Example:
    '   intMachineType = &H8664
    '   intReturnCode = ConvertPortableExecutableMachineTypeToProcessorArchitecture(strProcessorArchitecture, intMachineType)
    '   If intReturnCode = 0 Then
    '       ' The conversion of the PE machine type to its string processor architecture was
    '       ' successful. strProcessorArchitecture equals "AMD64".
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
    ' at https://github.com/franklesniak/sysadmin-accelerator
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
    '
    ' Phil Harvey, who write a reference to the tags present in Windows EXEs:
    ' https://exiftool.org/TagNames/EXE.html
    '
    ' Harry Johnston, for confirming that 0x1C4 is an ARM CPU Machine Type
    'endregion Acknowledgements ####################################################

    'region DependsOn ####################################################
    ' TestObjectForData()
    'endregion DependsOn ####################################################
    
    Dim intFunctionReturn
    Dim strWorkingProcessorArchitecture

    Err.Clear

    intFunctionReturn = 0

    If TestObjectForData(intMachineType) = False Then
        intFunctionReturn = -1
    Else
        On Error Resume Next
        Select Case intMachineType
            Case 34404 ' &H8664
                strWorkingProcessorArchitecture = "AMD64"
            Case 332 ' &H14C
                strWorkingProcessorArchitecture = "x86"
            Case 43620 ' &HAA64
                strWorkingProcessorArchitecture = "ARM64"
            Case 452 ' &H1C4
                ' This is the Machine Type observed on Windows RT (Windows 8/8.1 for ARM)
                strWorkingProcessorArchitecture = "ARM"
            Case 448 ' &H1C0
                strWorkingProcessorArchitecture = "ARM"
            Case 512 ' &H200
                strWorkingProcessorArchitecture = "IA64"
            Case 388 ' &H184
                strWorkingProcessorArchitecture = "ALPHA"
            Case 387 ' &H183
                ' &H183 is supposedly the former indicator for Alpha
                strWorkingProcessorArchitecture = "ALPHA"
            Case 644 ' &H284
                ' &H284 is supposedly the indicator for ALPHA64
                strWorkingProcessorArchitecture = "ALPHA64"
            Case 358 ' &H166
                strWorkingProcessorArchitecture = "MIPS"
            Case 870 ' &H366
                ' &H366 is supposedly the indicator for MIPS with FPU
                strWorkingProcessorArchitecture = "MIPS"
            Case 496 ' &H1F0
                strWorkingProcessorArchitecture = "PPC"
            Case 497 ' &H1F1
                ' &H1F1 is supposedly PowerPC with FPU
                strWorkingProcessorArchitecture = "PPC"
            Case Else
                intFunctionReturn = -2
        End Select
    End If

    If intFunctionReturn = 0 Then
        strProcessorArchitecture = strWorkingProcessorArchitecture
    End If

    ConvertPortableExecutableMachineTypeToProcessorArchitecture = intFunctionReturn
End Function
