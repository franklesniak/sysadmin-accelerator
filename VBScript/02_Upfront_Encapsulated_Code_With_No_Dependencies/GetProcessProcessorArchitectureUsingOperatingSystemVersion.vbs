Function GetProcessProcessorArchitectureUsingOperatingSystemVersion(ByRef strProcessProcessorArchitecture, ByVal strOperatingSystemVersion)
    'region FunctionMetadata ####################################################
    ' Safely determines the current VBScript process's processor architecture. The processor
    ' architecture is also known as "bit-width" or "bitness." However, the processor
    ' architecture describes not only the bit-width/bitness of the process, but also the
    ' general design and instruction set that the process uses to communicate with the
    ' processor.
    '
    ' This function takes two arguments:
    '   - The first argument (strProcessProcessorArchitecture) is populated upon success with
    '       the processor architecture used by the current VBScript process.
    '   - The second argument (strOperatingSystemVersion) contains the string representation of
    '       the operating system's version number in the format "major.minor.build.revision".
    '       Minimally, the major, minor, and build portions are required. If
    '       strOperatingSystemVersion is Null, an empty string (""), or otherwise lacks data,
    '       then this function will do its best to obtain this process's processor architecture
    '       without it. However, there are situations where the OS version number is required
    '       and its omission wil result in this function returning a failure code (see below).
    '
    ' The function returns 0 or a positive number if the VBScript process's architecture was
    ' retrieved successfully; it returns a negative number if the process's architecture was
    ' not retrived successfully.
    '
    ' Example:
    '   intReturnCode = GetProcessProcessorArchitectureUsingOperatingSystemVersion(strProcessProcessorArchitecture, "6.1.6701")
    '   If intReturnCode >= 0 Then
    '       ' Process's processor architecture retrieved successfully
    '       ' strProcessProcessorArchitecture equals either "x86" or "AMD64"
    '   End If
    '
    ' Note: the process's processor architecture may be the same as the operating system
    ' architecture or the physical processor architecture, but not always. Using emulation,
    ' Windows can allow processes to use processor architectures different than the operating
    ' system's native processor architecture. For example, Windows 10 version 20H2 running on
    ' ARM64 may run ARM64, Intel x86, or ARM (32-bit) processes.
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
    ' Version: 1.1.20210614.2
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
    ' TestObjectForData()
    ' GetPathToThisProcessExecutable()
    ' GetExecutableProcessorArchitectureFromFile()
    ' ConvertStringVersionNumberToMajorMinorBuildRevisionIntegers()
    ' TestOperatingSystemVersionRestrictsScriptProcessToX86Architecture()
    'endregion DependsOn ####################################################

    Dim intFunctionReturn
    Dim intReturnMultiplier
    Dim objWSHShell
    Dim objEnvironment
    Dim strWorkingProcessProcessorArchitecture
    Dim strPathToThisExecutable
    Dim lngMajor
    Dim lngMinor
    Dim lngBuild
    Dim lngRevision
    Dim boolOnlyX86Architecture

    Err.Clear
    
    intFunctionReturn = 0
    intReturnMultiplier = 1

    On Error Resume Next
    Set objWSHShell = WScript.CreateObject("WScript.Shell")
    If Err Then
        Err.Clear
        On Error Goto 0
        intFunctionReturn = intFunctionReturn + (-1 * intReturnMultiplier)
    Else
        Set objEnvironment = objWSHShell.Environment("Process")
        If Err Then
            Err.Clear
            On Error Goto 0
            intFunctionReturn = intFunctionReturn + (-2 * intReturnMultiplier)
        Else
            strWorkingProcessProcessorArchitecture = objEnvironment("PROCESSOR_ARCHITECTURE")
            If Err Then
                Err.Clear
                On Error Goto 0
                intFunctionReturn = intFunctionReturn + (-3 * intReturnMultiplier)
            Else
                On Error Goto 0
                If TestObjectForData(strWorkingProcessProcessorArchitecture) = False Then
                    intFunctionReturn = intFunctionReturn + (-4 * intReturnMultiplier)
                End If
            End If
        End If
    End If

    intReturnMultiplier = intReturnMultiplier * 8

    If intFunctionReturn < 0 Then
        intReturnCode = GetPathToThisProcessExecutable(strPathToThisExecutable)
        If intReturnCode <> 0 Then
            intFunctionReturn = intFunctionReturn + (-1 * intReturnMultiplier)
        Else
            intReturnCode = GetExecutableProcessorArchitectureFromFile(strWorkingProcessProcessorArchitecture, strPathToThisExecutable)
            If intReturnCode <> 0 Then
                intFunctionReturn = intFunctionReturn + (-2 * intReturnMultiplier)
            Else
                intFunctionReturn = 1
            End If
        End If
    End If

    If intFunctionReturn < 0 Then
        intReturnMultiplier = intReturnMultiplier * 4
        ' An error occurred above; see if we have OS version
        If TestObjectForData(strOperatingSystemVersion) = False Then
            intFunctionReturn = intFunctionReturn + (-1 * intReturnMultiplier)
        Else
            ' We have the OS version
            intReturnCode = ConvertStringVersionNumberToMajorMinorBuildRevisionIntegers(lngMajor, lngMinor, lngBuild, lngRevision, strOperatingSystemVersion)
            If intReturnCode < 0 Then
                intFunctionReturn = intFunctionReturn + (-2 * intReturnMultiplier)
            Else
                If lngMajor <= 0 Or lngMinor < 0 Or lngBuild < 0 Then
                    intFunctionReturn = intFunctionReturn + (-3 * intReturnMultiplier)
                Else
                    ' We have major.minor.build, at least
                    ' Check to see if this is Windows 95, 98, or ME - which would explain the
                    ' above failure to obtain the process' processor architecture
                    intReturnMultiplier = intReturnMultiplier * 4
                    intReturnCode = TestOperatingSystemVersionRestrictsScriptProcessToX86Architecture(boolOnlyX86Architecture, lngMajor, lngMinor, lngBuild)
                    If intReturnCode < 0 Then
                        intFunctionReturn = intFunctionReturn + (intReturnCode * intReturnMultiplier)
                    Else
                        ' Evaluation was successful
                        intReturnMultiplier = intReturnMultiplier * 4
                        If boolOnlyX86Architecture = False Then
                            intFunctionReturn = intFunctionReturn + (-1 * intReturnMultiplier)
                        Else
                            'Must be "x86"
                            intFunctionReturn = 2
                            strWorkingProcessProcessorArchitecture = "x86"
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    If intFunctionReturn >= 0 Then
        strProcessProcessorArchitecture = strWorkingProcessProcessorArchitecture
    End If
    
    GetProcessProcessorArchitectureUsingOperatingSystemVersion = intFunctionReturn
End Function
