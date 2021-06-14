Function TestProcessProcessorArchitectureAvailableOnWindowsVersionAndProcessorArchitecture(ByRef boolProcessorArchitectureIsAvailable, ByVal strProcessProcessorArchitectureToTest, ByVal strOperatingSystemVersionNumber, ByVal strOperatingSystemProcessorArchitecture)
    'region FunctionMetadata ####################################################
    ' Safely determines whether or not a hypothetical process's processor architecture is
    ' available (either natively, or through Windows-on-Windows emulation) on a given Windows
    ' operating system version and operating system processor architecture
    '
    ' This function takes four arguments:
    '   The first argument (boolProcessorArchitectureIsAvailable) is populated upon success
    '       with True if the processor architecture supplied in the second argument
    '       (strProcessProcessorArchitectureToTest) is available on the operating system
    '       version supplied in the third argument (strOperatingSystemVersionNumber) and the
    '       operating system processor architecture supplied in the fourth argument
    '       (strOperatingSystemProcessorArchitecture). It is set to False if the processor
    '       architecture supplied in the second argument
    '       (strProcessProcessorArchitectureToTest) is not available.
    '   The second argument (strProcessProcessorArchitectureToTest) is a string containing the
    '       processor architecture to evaluate as being available on the operating system
    '       version supplied in the third argument (strOperatingSystemVersionNumber) and the
    '       operating system processor architecture supplied in the fourth argument
    '       (strOperatingSystemProcessorArchitecture).
    '   The third argument (strOperatingSystemVersionNumber) provides the operating system
    '       version number in string format for the function to evaluate.
    '   The fourth argument (strOperatingSystemProcessorArchitecture) provides the operating
    '       system processor architecture in string format for the function to evaluate.
    '
    ' The function returns 0 if the processor architecture availability was evaluated
    ' successfully; it returns a negative number if the processor architecture availability
    ' was not evaluated successfully.
    '
    ' Example:
    '   intReturnCode = TestProcessProcessorArchitectureAvailableOnWindowsVersionAndProcessorArchitecture(boolProcessorArchitectureIsAvailable, "x86", "10.0.19042", "ARM64")
    '   If intReturnCode = 0 Then
    '       ' Processor architecture ("x86") availability evaluated successfully
    '       ' boolProcessorArchitectureIsAvailable is: True
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
    ' NOTE: VBScript is not capable of definitively determining whether a processor
    ' architecture is available on a given system because doing so requires calling a Win32
    ' function, which is not possible in VBScript. Instead, this function uses hard-coded
    ' "rules of thumb" to determine whether the processor architecture is available given the
    ' specified operating system version and architecture.
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
    ' Remko Weijnen, for his answer on StackExchange which gave me a pointer for how to make
    ' this determination programmatically from PowerShell:
    ' https://stackoverflow.com/a/66006120/2134110
    'endregion Acknowledgements ####################################################

    'region DependsOn ####################################################
    ' TestObjectIsStringContainingData()
    ' ConvertStringVersionNumberToMajorMinorBuildRevisionIntegers()
    'endregion DependsOn ####################################################

    Dim intFunctionReturn
    Dim boolWorkingProcessorArchitectureIsAvailable
    Dim intReturnCode
    Dim lngMajor
    Dim lngMinor
    Dim lngBuild
    Dim lngRevision

    Err.Clear

    intFunctionReturn = 0

    If TestObjectIsStringContainingData(strProcessProcessorArchitectureToTest) = False Then
        intFunctionReturn = -1
    Else
        If TestObjectIsStringContainingData(strOperatingSystemProcessorArchitecture) = False Then
            intFunctionReturn = -2
        End If
    End If

    If intFunctionReturn = 0 Then
        ' No error occurred
        ' strProcessProcessorArchitectureToTest is a string
        ' strOperatingSystemProcessorArchitecture is a string
        If UCase(strProcessProcessorArchitectureToTest) = UCase(strOperatingSystemProcessorArchitecture) Then
            boolWorkingProcessorArchitectureIsAvailable = True
        Else
            ' strProcessProcessorArchitectureToTest and strOperatingSystemProcessorArchitecture
            ' do not match
            If UCase(strOperatingSystemProcessorArchitecture) = "ALPHA" Or UCase(strOperatingSystemProcessorArchitecture) = "ALPHA64" Or UCase(strOperatingSystemProcessorArchitecture) = "MIPS" Or UCase(strOperatingSystemProcessorArchitecture) = "PPC" Then
                ' OS is ALPHA, ALPHA64, MIPS, or PPC; process processor architecture is
                ' something different.
                '
                ' Operating systems that run these processor architectures and that support
                ' VBScript do not support other Windows-on-Windows processor architectures
                boolWorkingProcessorArchitectureIsAvailable = False
            ElseIf UCase(strOperatingSystemProcessorArchitecture) = "X86" Then
                ' OS is x86; process processor architecture is something other than x86
                '
                ' If OS is x86, no other Windows-on-Windows processor architectures are
                ' available
                boolWorkingProcessorArchitectureIsAvailable = False
            ElseIf UCase(strOperatingSystemProcessorArchitecture) = "ARM" Then
                ' OS is ARM (32-bit); process processor architecture is something other than
                ' ARM
                '
                ' If OS is ARM, no other Windows-on-Windows processor architectures are
                ' available
                boolWorkingProcessorArchitectureIsAvailable = False
            ElseIf UCase(strOperatingSystemProcessorArchitecture) = "AMD64" Then
                ' OS is AMD64; process processor architecture is something other than AMD64
                
                If UCase(strProcessProcessorArchitectureToTest) = "X86" Then
                    ' OS is AMD64; process processor architecture is x86
                    ' x64 versions of Windows generally have x86 WOW support with a notable
                    ' exception that servers can remove the optional WOW64 component. However,
                    ' we are not testing for that here.
                    boolWorkingProcessorArchitectureIsAvailable = True
                Else
                    ' OS is AMD64; process processor architecture is not AMD64 nor x86
                    boolWorkingProcessorArchitectureIsAvailable = False
                End If
            ElseIf UCase(strOperatingSystemProcessorArchitecture) = "IA64" Then
                ' OS is IA64; process processor architecture is something other than IA64
                
                If UCase(strProcessProcessorArchitectureToTest) = "X86" Then
                    ' OS is IA64; process processor architecture is x86
                    ' Itanium versions of Windows have x86 WOW support with a notable
                    ' exception that Windows Server 2008/2008 R2 servers may allow removal of
                    ' the optional WOW64 component. However, we are not testing for that here.
                    boolWorkingProcessorArchitectureIsAvailable = True
                Else
                    ' OS is IA64; process processor architecture is not IA64 nor x86
                    boolWorkingProcessorArchitectureIsAvailable = False
                End If
            ElseIf UCase(strOperatingSystemProcessorArchitecture) = "ARM64" Then
                ' OS is ARM64; process processor architecture is something other than ARM64
                
                If UCase(strProcessProcessorArchitectureToTest) = "X86" Or UCase(strProcessProcessorArchitectureToTest) = "ARM" Then
                    ' OS is ARM64; process processor architecture is x86 or ARM
                    ' ARM64 versions of Windows include x86 and ARM WOW support
                    boolWorkingProcessorArchitectureIsAvailable = True
                ElseIf UCase(strProcessProcessorArchitectureToTest) = "AMD64" Then
                    ' OS is ARM64; process processor architecture is AMD64
                    ' Not supported in Windows 10 as of 2021-02-03, but present in Windows
                    ' Insider dev channel
                    ' OS is ARM64; process processor architecture is not ARM64, x86, nor ARM.
                    If TestObjectIsStringContainingData(strOperatingSystemVersionNumber) = False Then
                        intFunctionReturn = -3
                    Else
                        ' strOperatingSystemVersionNumber is a string
                        intReturnCode = ConvertStringVersionNumberToMajorMinorBuildRevisionIntegers(lngMajor, lngMinor, lngBuild, lngRevision, strOperatingSystemVersionNumber)
                        If intReturnCode <> 0 Then
                            intFunctionReturn = -4
                        Else
                            ' OS version number converted to lngMajor, lngMinor, lngBuild, lngRevision
                            If lngMajor = 10 And lngMinor = 0 And lngBuild >= 21277 Then
                                boolWorkingProcessorArchitectureIsAvailable = True
                            ElseIf lngMajor > 10 Or (lngMajor = 10 and lngMinor > 0) Then
                                boolWorkingProcessorArchitectureIsAvailable = True
                            Else
                                ' TODO: When Windows 10 is updated to support WOW emulation
                                '   of x64 on ARM64, insert a check here for the version
                                '   number. For now, return False because it's not
                                '   supported as of right now on
                                boolWorkingProcessorArchitectureIsAvailable = False
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If

    If intFunctionReturn = 0 Then
        boolProcessorArchitectureIsAvailable = boolWorkingProcessorArchitectureIsAvailable
    End If

    TestProcessProcessorArchitectureAvailableOnWindowsVersionAndProcessorArchitecture = intFunctionReturn
End Function
