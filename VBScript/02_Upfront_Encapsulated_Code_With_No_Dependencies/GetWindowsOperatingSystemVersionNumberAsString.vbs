Function GetWindowsOperatingSystemVersionNumberAsString(ByRef strOperatingSystemVersion, ByVal intRequirementForMajorVersionNumber, ByVal intRequirementForMinorVersionNumber, ByVal intRequirementForBuildVersionNumber, ByVal intRequirementForRevisionVersionNumber, ByVal objSpecificMethods)
    'region FunctionMetadata ####################################################
    ' Safely obtains the operating system version number using a variety of methods.
    '
    ' Function takes six positional arguments:
    '   The first argument (strOperatingSystemVersion) will be populated with the operating
    '       system version in string format upon success. Depending on what was supplied to the
    '       other arguments of this function, and depending on how successful the process was,
    '       the version number will be in one of these formats:
    '       "major.minor.build.revision"
    '       "major.minor.build"
    '       "major.minor"
    '       "build" (only)
    '       "revision" (only)
    '   The second argument (intRequirementForMajorVersionNumber) prescribes whether the major
    '       portion of the operating system's version number is to be retrieved, and to what
    '       degree of required accuracy:
    '       -1 = Do not retrieve the major portion of the version number; only valid if the
    '           third argument (intRequirementForMinorVersionNumber) is also -1, and then
    '           either:
    '               - The fourth argument (intRequirementForBuildVersionNumber) is >= 1 (see
    '                 below), and the fifth argument (intRequirementForRevisionVersionNumber)
    '                 is -1
    '               or
    '               - The fourth argument (intRequirementForBuildVersionNumber) is -1, and the
    '                 fifth argument (intRequirementForRevisionVersionNumber) is >= 1 (see
    '                 below)
    '       Null, 0, or 1 = The major version number is required and every attempt will be made
    '           to retrieve it to the highest degree of accuracy available. If the major
    '           version number cannot be retrieved, an error is returned.
    '       2 = The major portion of the version number is required with an accuracy level of
    '           2/7, with 7 being the highest. If the major portion could not be determined, or
    '           if this level of accuracy is not met, an error is returned.
    '       3 = The major portion of the version number is required with an accuracy level of
    '           3/7, with 7 being the highest. If the major portion could not be determined, or
    '           if this level of accuracy is not met, an error is returned.
    '       4 = The major portion of the version number is required with an accuracy level of
    '           4/7, with 7 being the highest. If the major portion could not be determined, or
    '           if this level of accuracy is not met, an error is returned.
    '       5 = The major portion of the version number is required with an accuracy level of
    '           5/7, with 7 being the highest. If the major portion could not be determined, or
    '           if this level of accuracy is not met, an error is returned.
    '       6 = The major portion of the version number is required with an accuracy level of
    '           6/7, with 7 being the highest. If the major portion could not be determined, or
    '           if this level of accuracy is not met, an error is returned.
    '       7 = The major portion of the version number is required with an accuracy level of
    '           7/7, with 7 being the highest. If the major portion could not be determined, or
    '           if this level of accuracy is not met, an error is returned.
    '   The third argument (intRequirementForMinorVersionNumber) behaves exactly like the
    '       second argument (intRequirementForMajorVersionNumber), except that references to
    '       the major portion of the version number are swapped with references to the minor
    '       portion of the version number, and vice versa.
    '   The fourth argument (intRequirementForBuildVersionNumber) prescribes whether the build
    '       portion of the operating system's version number is to be retrieved, and to what
    '       degree of required accuracy:
    '       -1 = Do not retrieve the build portion of the version number; if -1 is specified,
    '           the function will omit the build number from the version number returned in the
    '           first argument (strOperatingSystemVersion) even if the function was able to
    '           retrieve the build number successfully. -1 is only valid if -1 is also supplied
    '           for the fifth argument (intRequirementForRevisionVersionNumber).
    '       Null or 0 = The function should attempt to retrieve the build portion of the
    '           version number to the highest degree of accuracy available. However, the build
    '           portion is not required. If the function was not able to retrieve the build
    '           portion of the version number, then it will be omitted along with the revision
    '           portion of the version number.
    '       1 = The build portion of the version number is required with an accuracy level of
    '           1/7, with 7 being the highest. If the build number could not be determined, or
    '           if this level of accuracy is not met, an error is returned.
    '       2 = The build portion of the version number is required with an accuracy level of
    '           2/7, with 7 being the highest. If the build number could not be determined, or
    '           if this level of accuracy is not met, an error is returned.
    '       3 = The build portion of the version number is required with an accuracy level of
    '           3/7, with 7 being the highest. If the build number could not be determined, or
    '           if this level of accuracy is not met, an error is returned.
    '       4 = The build portion of the version number is required with an accuracy level of
    '           4/7, with 7 being the highest. If the build number could not be determined, or
    '           if this level of accuracy is not met, an error is returned.
    '       5 = The build portion of the version number is required with an accuracy level of
    '           5/7, with 7 being the highest. If the build number could not be determined, or
    '           if this level of accuracy is not met, an error is returned.
    '       6 = The build portion of the version number is required with an accuracy level of
    '           6/7, with 7 being the highest. If the build number could not be determined, or
    '           if this level of accuracy is not met, an error is returned.
    '       7 = The build portion of the version number is required with an accuracy level of
    '           7/7, with 7 being the highest. If the build number could not be determined, or
    '           if this level of accuracy is not met, an error is returned.
    '   The fifth argument (intRequirementForRevisionVersionNumber) prescribes whether the
    '       revision portion of the operating system's version number is to be retrieved, and
    '       to what degree of required accuracy:
    '       -1 = Do not retrieve the revision portion of the version number; if -1 is specified,
    '           the function will omit the revision number from the version number returned in
    '           the first argument (strOperatingSystemVersion) even if the function was able to
    '           retrieve the revision number successfully.
    '       Null or 0 = The function should attempt to retrieve the revision portion of the
    '           version number to the highest degree of accuracy available. However, the
    '           revision portion is not required. If the function was not able to retrieve the
    '           revision portion of the version number, then it will be omitted.
    '       1 = The revision portion of the version number is required with an accuracy level
    '           of 1/7, with 7 being the highest. If the revision number could not be
    '           determined, or if this level of accuracy is not met, an error is returned.
    '       2 = The revision portion of the version number is required with an accuracy level
    '           of 2/7, with 7 being the highest. If the revision number could not be
    '           determined, or if this level of accuracy is not met, an error is returned.
    '       3 = The revision portion of the version number is required with an accuracy level
    '           of 3/7, with 7 being the highest. If the revision number could not be
    '           determined, or if this level of accuracy is not met, an error is returned.
    '       4 = The revision portion of the version number is required with an accuracy level
    '           of 4/7, with 7 being the highest. If the revision number could not be
    '           determined, or if this level of accuracy is not met, an error is returned.
    '       5 = The revision portion of the version number is required with an accuracy level
    '           of 5/7, with 7 being the highest. If the revision number could not be
    '           determined, or if this level of accuracy is not met, an error is returned.
    '       6 = The revision portion of the version number is required with an accuracy level
    '           of 6/7, with 7 being the highest. If the revision number could not be
    '           determined, or if this level of accuracy is not met, an error is returned.
    '       7 = The revision portion of the version number is required with an accuracy level
    '           of 7/7, with 7 being the highest. If the revision number could not be
    '           determined, or if this level of accuracy is not met, an error is returned.
    '   The sixth argument (objSpecificMethods) prescribes one or more methods by which the
    '       operating system version number is to be retrieved:
    '       Null or 0 = caller does not care; proceed with the default behavior outlined below
    '       One number, 1-9 = use only the corresponding numbered method outlined below
    '       Several numbers in an array (e.g., Array(1, 2, 3)) = use the numbered methods that
    '           correspond with the items in the array
    '
    ' The function returns 0 or a positive number if the operating system version number was
    ' retrieved successfully. A negative number is returned if the operating system version
    ' number was not retrieved successfully.
    '
    ' The function obtains the major.minor.build.revision portions of the OS version number
    ' by going down the following list in order of preference. If any portion of the version
    ' number was not retrieved accurately, it continues down the list as needed to obtain it to
    ' the highest available degree of accuracy:
    ' 1. Win32 -> RtlGetVersion (listed for completion, but not possible in VBScript; does not
    '       include revision number)
    ' 2. WMI -> Win32_OperatingSystem
    '       If successful:
    '       - Major version is supplied, tamper-proof, and accurate (7/7)
    '       - Minor version is supplied, tamper-proof, and accurate (7/7)
    '       - Build should be supplied, and is tamper-proof and accurate (7/7 if supplied)
    '       - Revision may be supplied and would be tamper-proof, but it is NOT accurate.
    '         Reportedly, WMI may return an inaccurate revision number of .0). (0/7 if
    '         supplied)
    ' 3a. "Product version* of the file C:\Windows\System32\ntoskrnl.exe (may include revision
    '       number; Windows NT 4.0, Windows 2000, Windows XP, and newer)
    '       If successful:
    '       - Major version is supplied, tamper-proof, and accurate (7/7)
    '       - Minor version is supplied, tamper-proof, and accurate (7/7)
    '       - Build is supplied and is tamper-proof, but not accurate in all cases.
    '         Specifically, Windows 10 builds that required an "enablement package" keep the
    '         previous build binaries side-by-side with the new build's binaries. Therefore,
    '         this file may inaccurately report the build number of the previous build.
    '         (If Major < 10, 7/7; If Major = 10, Minor = 0, and Build < 18362, 7/7; else 3/7)
    '       - Revision may be supplied; when it is, it is tamper-proof and likely accurate (7/7
    '         if supplied)
    ' 3b. "Product version* of the file C:\Windows\System\krnl386.exe (may include revision
    '       number; Windows 95, Windows 98, and Windows ME only)
    '       If successful:
    '       - Major version is supplied, tamper-proof, and accurate (7/7)
    '       - Minor version is supplied, tamper-proof, and accurate (7/7)
    '       - Build is supplied, tamper-proof, and accurate (7/7)
    '       - Revision may be supplied; when it is, it is tamper-proof and likely accurate (7/7
    '         if supplied)
    ' 4. cmd /c ver > temp file (may include revision number)
    '       If successful:
    '       - Major version is supplied and presumed tamper-proof and accurate (6/7)
    '       - Minor version is supplied and presumed tamper-proof and accurate (6/7)
    '       - Build is usually supplied and presumed tamper-proof and accurate. Notably, it is
    '         absent on Windows NT 4.0 (6/7)
    '       - Revision is sometimes supplied (specifically, on Windows 10 and Windows Server
    '         equivalents). It appears to be sourced from a registry key, so less tamper-proof
    '         and inaccurate if tampered (4/7)
    ' 5a. "File version" of the file C:\Windows\System32\ntoskrnl.exe (may include revision
    '       number; Windows NT 4.0, Windows 2000, Windows XP, and newer)
    '       If successful:
    '       - Major version is supplied, but is NOT tamper-proof, and may be inaccurate if
    '         tampered (3/7)
    '       - Minor version is supplied, but is NOT tamper-proof, and may be inaccurate if
    '         tampered (3/7)
    '       - Build is supplied, but is NOT tamper-proof, and may be inaccurate if tampered,
    '         and not accurate in all cases anyway. Specifically, Windows 10 builds that
    '         required an "enablement package" keep the previous build binaries side-by-side
    '         with the new build's binaries. Therefore, this file may inaccurately report the
    '         build number of the previous build. (If Major < 10, 3/7; If Major = 10,
    '         Minor = 0, and Build < 18362, 3/7; else 2/7)
    '       - Revision may be supplied, but it is NOT tamper-proof, and may be inaccurate if
    '         tampered. (3/7 if supplied)
    ' 5b. "File version" of the file C:\Windows\System\krnl386.exe (may include revision
    '       number; Windows 95, Windows 98, and Windows ME only)
    '       If successful:
    '       - Major version is supplied, but is NOT tamper-proof, and may be inaccurate if
    '         tampered (3/7)
    '       - Minor version is supplied, but is NOT tamper-proof, and may be inaccurate if
    '         tampered (3/7)
    '       - Build is supplied, but is NOT tamper-proof, and may be inaccurate if tampered
    '         (3/7)
    '       - Revision may be supplied, but it is NOT tamper-proof, and may be inaccurate if
    '         tampered. (3/7 if supplied)
    ' 6a. "Product version* of the file C:\Windows\System32\kernel32.dll (may include a
    '       questionably-accurate revision number; Windows NT 4.0, Windows 2000, Windows XP,
    '       and newer)
    '       If successful:
    '       - Major version is supplied, tamper-proof, and accurate (7/7)
    '       - Minor version is supplied, tamper-proof, and accurate (7/7)
    '       - Build is supplied and is tamper-proof, but not accurate in all cases.
    '         Specifically, Windows 10 builds that required an "enablement package" keep the
    '         previous build binaries side-by-side with the new build's binaries. Therefore,
    '         this file may inaccurately report the build number of the previous build.
    '         (If Major < 10, 7/7; If Major = 10, Minor = 0, and Build < 18362, 7/7; else 3/7)
    '       - Revision may be supplied; when it is, it is tamper-proof but often not the same
    '         as the revision number of the OS kernel (4/7 if supplied)
    ' 6b. "Product version* of the file C:\Windows\System32\ntdll.dll (may include a
    '       questionably-accurate revision number; Windows NT 4.0, Windows 2000, Windows XP,
    '       and newer)
    '       If successful:
    '       - Major version is supplied, tamper-proof, and accurate (7/7)
    '       - Minor version is supplied, tamper-proof, and accurate (7/7)
    '       - Build is supplied and is tamper-proof, but not accurate in all cases.
    '         Specifically, Windows 10 builds that required an "enablement package" keep the
    '         previous build binaries side-by-side with the new build's binaries. Therefore,
    '         this file may inaccurately report the build number of the previous build.
    '         (If Major < 10, 7/7; If Major = 10, Minor = 0, and Build < 18362, 7/7; else 3/7)
    '       - Revision may be supplied; when it is, it is tamper-proof but often not the same
    '         as the revision number of the OS kernel (3/7 if supplied)
    ' 6c. "Product version* of the file C:\Windows\System32\hal.dll (may include a
    '       questionably-accurate revision number; Windows NT 4.0, Windows 2000, Windows XP,
    '       and newer)
    '       If successful:
    '       - Major version is supplied, tamper-proof, and accurate (7/7)
    '       - Minor version is supplied, tamper-proof, and accurate (7/7)
    '       - Build is supplied and is tamper-proof, but not accurate in all cases.
    '         Specifically, Windows 10 builds that required an "enablement package" keep the
    '         previous build binaries side-by-side with the new build's binaries. Therefore,
    '         this file may inaccurately report the build number of the previous build.
    '         (If Major < 10, 7/7; If Major = 10, Minor = 0, and Build < 18362, 7/7; else 3/7)
    '       - Revision may be supplied; when it is, it is tamper-proof but often not the same
    '         as the revision number of the OS kernel (2/7 if supplied)
    ' 7a. "File version" of the file C:\Windows\System32\kernel32.dll (may include a
    '       questionably-accurate revision number; Windows NT 4.0, Windows 2000, Windows XP,
    '       and newer)
    '       If successful:
    '       - Major version is supplied, but is NOT tamper-proof, and may be inaccurate if
    '         tampered (3/7)
    '       - Minor version is supplied, but is NOT tamper-proof, and may be inaccurate if
    '         tampered (3/7)
    '       - Build is supplied, but is NOT tamper-proof, and may be inaccurate if tampered,
    '         and not accurate in all cases anyway. Specifically, Windows 10 builds that
    '         required an "enablement package" keep the previous build binaries side-by-side
    '         with the new build's binaries. Therefore, this file may inaccurately report the
    '         build number of the previous build. (If Major < 10, 3/7; If Major = 10,
    '         Minor = 0, and Build < 18362, 3/7; else 2/7)
    '       - Revision may be supplied, but it is NOT tamper-proof, and may be inaccurate if
    '         tampered. (2/7 if supplied)
    ' 7b. "File version" of the file C:\Windows\System32\ntdll.dll (may include a
    '       questionably-accurate revision number; Windows NT 4.0, Windows 2000, Windows XP,
    '       and newer)
    '       If successful:
    '       - Major version is supplied, but is NOT tamper-proof, and may be inaccurate if
    '         tampered (3/7)
    '       - Minor version is supplied, but is NOT tamper-proof, and may be inaccurate if
    '         tampered (3/7)
    '       - Build is supplied, but is NOT tamper-proof, and may be inaccurate if tampered,
    '         and not accurate in all cases anyway. Specifically, Windows 10 builds that
    '         required an "enablement package" keep the previous build binaries side-by-side
    '         with the new build's binaries. Therefore, this file may inaccurately report the
    '         build number of the previous build. (If Major < 10, 3/7; If Major = 10,
    '         Minor = 0, and Build < 18362, 3/7; else 2/7)
    '       - Revision may be supplied, but it is NOT tamper-proof, and may be inaccurate if
    '         tampered. (1/7 if supplied)
    ' 7c. "File version" of the file C:\Windows\System32\hal.dll (may include a
    '       questionably-accurate revision number; Windows NT 4.0, Windows 2000, Windows XP,
    '       and newer)
    '       If successful:
    '       - Major version is supplied, but is NOT tamper-proof, and may be inaccurate if
    '         tampered (3/7)
    '       - Minor version is supplied, but is NOT tamper-proof, and may be inaccurate if
    '         tampered (3/7)
    '       - Build is supplied, but is NOT tamper-proof, and may be inaccurate if tampered,
    '         and not accurate in all cases anyway. Specifically, Windows 10 builds that
    '         required an "enablement package" keep the previous build binaries side-by-side
    '         with the new build's binaries. Therefore, this file may inaccurately report the
    '         build number of the previous build. (If Major < 10, 3/7; If Major = 10,
    '         Minor = 0, and Build < 18362, 3/7; else 2/7)
    '       - Revision may be supplied, but it is NOT tamper-proof, and may be inaccurate if
    '         tampered. (1/7 if supplied)
    ' 8. HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion ->
    '       CurrentMajorVersionNumber (DWORD, not tamper-proof, Windows 10 only)
    '       CurrentMinorVersionNumber (DWORD, not tamper-proof, Windows 10 only)
    '       CurrentBuildNumber (String, not tamper-proof)
    '       If successful:
    '       - Major version is supplied, but is NOT tamper-proof, and may be inaccurate if
    '         tampered (3/7)
    '       - Minor version is supplied, but is NOT tamper-proof, and may be inaccurate if
    '         tampered (3/7)
    '       - Build is supplied, but is NOT tamper-proof, and may be inaccurate if tampered
    '         (3/7)
    ' 9. (Revision number, Windows 10 only):
    '       HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion ->
    '       UBR (DWORD, not tamper-proof)
    '       If successful:
    '       - Revision is NOT tamper-proof, and may be inaccurate if tampered. (3/7)
    '
    ' Example 1: Get the Windows version number without using the registry
    '   objSpecificMethods = Array(1, 2, 3, 4, 5, 6, 7)
    '   intReturnCode = GetWindowsOperatingSystemVersionNumberAsString(strOperatingSystemVersion, Null, Null, Null, Null, objSpecificMethods)
    '   If intReturnCode >= 0 Then
    '       ' strOperatingSystemVersion is populated with the OS version, as expected
    '   End If
    '
    ' Example 2:
    '   'TODO: Fill in
    '
    ' Version: 0.1.20210614.0
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
    ' User "Shem Sargent" on Super User, who provided sample code for augmenting WMI-based
    ' version numbers with their revision number:
    ' https://superuser.com/a/1160428/334370
    '
    ' The NTFSSecurity PowerShell Module team, who provided this promising-looking snapshot of
    ' some C# code that uses RtlGetVersion:
    ' https://git.sern.network/Powershell-Modules/NTFSSecurity-Module/src/commit/d5f552df8cb60f6f3c7b554445da1b7631fa326f/AlphaFS/OperatingSystem.cs
    '
    ' GitHub user "SevenLayerJedi", who posted this Gist that affirmed some approaches:
    ' https://gist.github.com/SevenLayerJedi/c0415c03cab1ff51aa49d2f4d708f265
    'endregion Acknowledgements ####################################################

    'region DependsOn ####################################################
    ' TestObjectForData()
    ' GetWindowsOperatingSystemVersionNumberAsStringUsingWMI()
    ' ConvertStringVersionNumberToMajorMinorBuildRevisionIntegers()
    ' GetWindowsPath()
    ' GetFileProductVersionAsString()
    ' GetWindowsSystemPath()
    ' GetWindowsOperatingSystemVersionNumberAsStringUsingCommandPromptVerCommand()
    ' GetFileVersionAsString()
    'endregion DependsOn ####################################################

    Dim lngFunctionReturn
    Dim lngOSMajor
    Dim intOSMajorCurrentLevelOfAccuracy
    Dim lngOSMinor
    Dim intOSMinorCurrentLevelOfAccuracy
    Dim lngOSBuild
    Dim intOSBuildCurrentLevelOfAccuracy
    Dim lngOSRevision
    Dim intOSRevisionCurrentLevelOfAccuracy
    Dim boolTest
    Dim boolMethod1
    Dim boolMethod2
    Dim boolMethod3
    Dim boolMethod4
    Dim boolMethod5
    Dim boolMethod6
    Dim boolMethod7
    Dim boolMethod8
    Dim boolMethod9
    Dim intUpperBound
    Dim intCounter
    Dim intWorkingAccuracyTargetForMajor
    Dim intWorkingAccuracyTargetForMinor
    Dim intWorkingAccuracyTargetForBuild
    Dim intWorkingAccuracyTargetForRevision
    Dim intThisMethodNumber
    Dim boolMethodSucceeded
    Dim intThisMethodMajorLevelOfAccuracy
    Dim intThisMethodMinorLevelOfAccuracy
    Dim intThisMethodBuildLevelOfAccuracy
    Dim intThisMethodRevisionLevelOfAccuracy
    Dim intReturnCode
    Dim strTempOSVersion
    Dim lngTempMajor
    Dim lngTempMinor
    Dim lngTempBuild
    Dim lngTempRevision
    Dim lngReturnMultiplier
    Dim strOSVersionStaging
    Dim intTempLevelOfAccuracy
    Dim strWindowsPath
    Dim strWindowsSystemPath

    Err.Clear

    lngFunctionReturn = 0
    strOSVersionStaging = ""
    lngOSMajor = CLng(-1)
    intOSMajorCurrentLevelOfAccuracy = -1
    lngOSMinor = CLng(-1)
    intOSMinorCurrentLevelOfAccuracy = -1
    lngOSBuild = CLng(-1)
    intOSBuildCurrentLevelOfAccuracy = -1
    lngOSRevision = CLng(-1)
    intOSRevisionCurrentLevelOfAccuracy = -1

    ' Determine the method numbers to use based on input
    boolMethod1 = True
    boolMethod2 = True
    boolMethod3 = True
    boolMethod4 = True
    boolMethod5 = True
    boolMethod6 = True
    boolMethod7 = True
    boolMethod8 = True
    boolMethod9 = True
    If TestObjectForData(objSpecificMethods) = True Then
        On Error Resume Next
        boolTest = IsArray(objSpecificMethods)
        If Err Then
            ' Data is present but IsArray() failed
            On Error Goto 0
            Err.Clear
            boolMethod1 = False
            boolMethod2 = False
            boolMethod3 = False
            boolMethod4 = False
            boolMethod5 = False
            boolMethod6 = False
            boolMethod7 = False
            boolMethod8 = False
            boolMethod9 = False
            Select Case objSpecificMethods
                Case 1
                    boolMethod1 = True
                Case 2
                    boolMethod2 = True
                Case 3
                    boolMethod3 = True
                Case 4
                    boolMethod4 = True
                Case 5
                    boolMethod5 = True
                Case 6
                    boolMethod6 = True
                Case 7
                    boolMethod7 = True
                Case 8
                    boolMethod8 = True
                Case 9
                    boolMethod9 = True
            End Select
        Else
            ' Data is present; IsArray() did not result in an error
            On Error Goto 0
            If boolTest = True Then
                ' Data is present; objSpecificMethods is an array
                On Error Resume Next
                intUpperBound = UBound(objSpecificMethods)
                If Err Then
                    ' Data is present; objSpecificMethods is an array; but an error occurred
                    ' getting the upper bound of the array
                    On Error Goto 0
                    Err.Clear
                Else
                    ' Data is present; objSpecificMethods is an array; we have the upper bound
                    ' of the array
                    On Error Goto 0
                    boolMethod1 = False
                    boolMethod2 = False
                    boolMethod3 = False
                    boolMethod4 = False
                    boolMethod5 = False
                    boolMethod6 = False
                    boolMethod7 = False
                    boolMethod8 = False
                    boolMethod9 = False
                    For intCounter = 0 To intUpperBound
                        Select Case objSpecificMethods(intCounter)
                            Case 1
                                boolMethod1 = True
                            Case 2
                                boolMethod2 = True
                            Case 3
                                boolMethod3 = True
                            Case 4
                                boolMethod4 = True
                            Case 5
                                boolMethod5 = True
                            Case 6
                                boolMethod6 = True
                            Case 7
                                boolMethod7 = True
                            Case 8
                                boolMethod8 = True
                            Case 9
                                boolMethod9 = True
                        End Select
                    Next
                End If
            Else
                ' Data is present; objSpecificMethods is not an array
                ' Assume objSpecificMethods is an integer
                boolMethod1 = False
                boolMethod2 = False
                boolMethod3 = False
                boolMethod4 = False
                boolMethod5 = False
                boolMethod6 = False
                boolMethod7 = False
                boolMethod8 = False
                boolMethod9 = False
                Select Case objSpecificMethods
                    Case 1
                        boolMethod1 = True
                    Case 2
                        boolMethod2 = True
                    Case 3
                        boolMethod3 = True
                    Case 4
                        boolMethod4 = True
                    Case 5
                        boolMethod5 = True
                    Case 6
                        boolMethod6 = True
                    Case 7
                        boolMethod7 = True
                    Case 8
                        boolMethod8 = True
                    Case 9
                        boolMethod9 = True
                End Select
            End If
        End If
    End If

    If TestObjectForData(intRequirementForMajorVersionNumber) = False Then
        intWorkingAccuracyTargetForMajor = 0
    Else
        intWorkingAccuracyTargetForMajor = intRequirementForMajorVersionNumber
    End If

    If TestObjectForData(intRequirementForMinorVersionNumber) = False Then
        intWorkingAccuracyTargetForMinor = 0
    Else
        intWorkingAccuracyTargetForMinor = intRequirementForMinorVersionNumber
    End If

    If TestObjectForData(intRequirementForBuildVersionNumber) = False Then
        intWorkingAccuracyTargetForBuild = 0
    Else
        intWorkingAccuracyTargetForBuild = intRequirementForBuildVersionNumber
    End If

    If TestObjectForData(intRequirementForRevisionVersionNumber) = False Then
        intWorkingAccuracyTargetForRevision = 0
    Else
        intWorkingAccuracyTargetForRevision = intRequirementForRevisionVersionNumber
    End If

    If intWorkingAccuracyTargetForMajor = -1 And intWorkingAccuracyTargetForMinor >= 0 Then
        ' Can't do minor without major
        lngFunctionReturn = -1
    ElseIf intWorkingAccuracyTargetForMajor >= 0 And intWorkingAccuracyTargetForMinor = -1 Then
        ' Can't do major without minor
        lngFunctionReturn = -1
    ElseIf intWorkingAccuracyTargetForMajor = -1 And intWorkingAccuracyTargetForMinor = -1 And intWorkingAccuracyTargetForBuild < 1 And intWorkingAccuracyTargetForRevision < 1 Then
        ' Can't omit major and minor without making build or revision mandatory
        lngFunctionReturn = -1
    ElseIf intWorkingAccuracyTargetForMajor = -1 And intWorkingAccuracyTargetForMinor = -1 And intWorkingAccuracyTargetForBuild >= 1 And intWorkingAccuracyTargetForRevision >= 1 Then
        ' Can't omit major and minor and make *both* build and revision mandatory
        lngFunctionReturn = -1
    ElseIf intWorkingAccuracyTargetForMajor >= 0 And intWorkingAccuracyTargetForMinor >= 0 And intWorkingAccuracyTargetForBuild = -1 And intWorkingAccuracyTargetForRevision >= 0 Then
        ' If major and minor are in play, can't omit build but allow revision
        lngFunctionReturn = -1
    End If
    
    ' #########################################################################################
    intThisMethodNumber = 1 ' (Win32 -> RtlGetVersion)
    lngReturnMultiplier = &H00000001
    If lngFunctionReturn >= 0 Then
        ' This needs to be set to the boolean variable that matches the method number: ########
        If boolMethod1 Then
            ' Possibly attempt method
            boolMethodSucceeded = False ' Not available on VBScript
        End If
    End If

    ' #########################################################################################
    intThisMethodNumber = 2 ' (WMI -> Win32_OperatingSystem)
    lngReturnMultiplier = &H00000002
    intThisMethodMajorLevelOfAccuracy = 7
    intThisMethodMinorLevelOfAccuracy = 7
    intThisMethodBuildLevelOfAccuracy = 7
    intThisMethodRevisionLevelOfAccuracy = -1
    If lngFunctionReturn >= 0 Then
        ' This needs to be set to the boolean variable that matches the method number: ########
        If boolMethod2 Then
            ' Possibly attempt method
            If intOSMajorCurrentLevelOfAccuracy < intThisMethodMajorLevelOfAccuracy Or intOSMinorCurrentLevelOfAccuracy < intThisMethodMinorLevelOfAccuracy Or intOSBuildCurrentLevelOfAccuracy < intThisMethodBuildLevelOfAccuracy Or intOSRevisionCurrentLevelOfAccuracy < intThisMethodRevisionLevelOfAccuracy Then
                ' It makes sense to attempt this method as this could improve accuracy compared
                ' to the current version information that we have
                boolMethodSucceeded = True

                ' Make the attempt and set boolMethodSucceeded = False if failure
                intReturnCode = GetWindowsOperatingSystemVersionNumberAsStringUsingWMI(strTempOSVersion, True)
                If intReturnCode < 0 Then
                    boolMethodSucceeded = False
                Else
                    intReturnCode =  ConvertStringVersionNumberToMajorMinorBuildRevisionIntegers(lngTempMajor, lngTempMinor, lngTempBuild, lngTempRevision, strTempOSVersion)
                    If intReturnCode < 0 Then
                        boolMethodSucceeded = False
                    Else
                        ' Success!
                    End If
                End If

                If boolMethodSucceeded = True Then
                    lngFunctionReturn = lngFunctionReturn + lngReturnMultiplier
                    If lngTempMajor <> CLng(-1) Then
                        ' Major version obtained
                        If intOSMajorCurrentLevelOfAccuracy < intThisMethodMajorLevelOfAccuracy Then
                            lngOSMajor = lngTempMajor
                            intOSMajorCurrentLevelOfAccuracy = intThisMethodMajorLevelOfAccuracy
                        ElseIf intOSMajorCurrentLevelOfAccuracy = intThisMethodMajorLevelOfAccuracy Then
                            If lngTempMajor > lngOSMajor Then
                                lngOSMajor = lngTempMajor
                            End If
                        End If
                    End If

                    If lngTempMinor <> CLng(-1) Then
                        ' Minor version obtained
                        If intOSMinorCurrentLevelOfAccuracy < intThisMethodMinorLevelOfAccuracy Then
                            lngOSMinor = lngTempMinor
                            intOSMinorCurrentLevelOfAccuracy = intThisMethodMinorLevelOfAccuracy
                        ElseIf intOSMinorCurrentLevelOfAccuracy = intThisMethodMinorLevelOfAccuracy Then
                            If lngTempMinor > lngOSMinor Then
                                lngOSMinor = lngTempMinor
                            End If
                        End If
                    End If

                    If lngTempBuild <> CLng(-1) Then
                        ' Build version obtained
                        If intOSBuildCurrentLevelOfAccuracy < intThisMethodBuildLevelOfAccuracy Then
                            lngOSBuild = lngTempBuild
                            intOSBuildCurrentLevelOfAccuracy = intThisMethodBuildLevelOfAccuracy
                        ElseIf intOSBuildCurrentLevelOfAccuracy = intThisMethodBuildLevelOfAccuracy Then
                            If lngTempBuild > lngOSBuild Then
                                lngOSBuild = lngTempBuild
                            End If
                        End If
                    End If

                    If lngTempRevision <> CLng(-1) Then
                        ' Revision version obtained
                        If intOSRevisionCurrentLevelOfAccuracy < intThisMethodRevisionLevelOfAccuracy Then
                            lngOSRevision = lngTempRevision
                            intOSRevisionCurrentLevelOfAccuracy = intThisMethodRevisionLevelOfAccuracy
                        ElseIf intOSRevisionCurrentLevelOfAccuracy = intThisMethodRevisionLevelOfAccuracy Then
                            If lngTempRevision > lngOSRevision Then
                                lngOSRevision = lngTempRevision
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If

    ' #########################################################################################
    intThisMethodNumber = 3 ' ("Product version" of the file C:\Windows\Sysnative\ntoskrnl.exe,
    '                         C:\Windows\System32\ntoskrnl.exe, or
    '                         C:\Windows\System\krnl386.exe)
    lngReturnMultiplier = &H00000004
    intThisMethodMajorLevelOfAccuracy = 7
    intThisMethodMinorLevelOfAccuracy = 7
    intThisMethodBuildLevelOfAccuracy = 7 ' (If Major < 10, 7/7; If Major = 10, Minor = 0, and
    '                                       Build < 18362, 7/7; else 3/7)
    intThisMethodRevisionLevelOfAccuracy = 7
    If lngFunctionReturn >= 0 Then
        ' This needs to be set to the boolean variable that matches the method number: ########
        If boolMethod3 Then
            ' Possibly attempt method
            If intOSMajorCurrentLevelOfAccuracy < intThisMethodMajorLevelOfAccuracy Or intOSMinorCurrentLevelOfAccuracy < intThisMethodMinorLevelOfAccuracy Or intOSBuildCurrentLevelOfAccuracy < intThisMethodBuildLevelOfAccuracy Or intOSRevisionCurrentLevelOfAccuracy < intThisMethodRevisionLevelOfAccuracy Then
                ' It makes sense to attempt this method as this could improve accuracy compared
                ' to the current version information that we have
                boolMethodSucceeded = True

                ' Make the attempt and set boolMethodSucceeded = False if failure

                ' First, get the Windows path if we don't already have it
                If TestObjectForData(strWindowsPath) = False Then
                    intReturnCode = GetWindowsPath(strWindowsPath)
                    If intReturnCode <> 0 Then
                        boolMethodSucceeded = False
                    End If
                End If

                If boolMethodSucceeded = True Then
                    ' No error occurred yet
                    intReturnCode = GetFileProductVersionAsString(strTempOSVersion, strWindowsPath & "Sysnative\ntoskrnl.exe")
                    If intReturnCode <> 0 Then
                        boolMethodSucceeded = False
                    Else
                        intReturnCode =  ConvertStringVersionNumberToMajorMinorBuildRevisionIntegers(lngTempMajor, lngTempMinor, lngTempBuild, lngTempRevision, strTempOSVersion)
                        If intReturnCode < 0 Then
                            boolMethodSucceeded = False
                        Else
                            ' Success!
                        End If
                    End If
                End If

                If boolMethodSucceeded = False Then
                    ' C:\Windows\Sysnative\ntoskrnl.exe failed.
                    ' This would be expected if we are not using Windows-on-Windows (WOW),
                    ' e.g., the current process is not 32-bit x86 on 64-bit AMD64 Windows, or
                    ' if the operating system is Windows XP or Windows Server 2003/2003 R2 and
                    ' KB942589 is missing.
                    boolMethodSucceeded = True

                    ' Get the system path if we don't already have it
                    If TestObjectForData(strWindowsSystemPath) = False Then
                        intReturnCode = GetWindowsSystemPath(strWindowsSystemPath)
                        If intReturnCode <> 0 Then
                            boolMethodSucceeded = False
                        End If
                    End If

                    If boolMethodSucceeded = True Then
                        ' No error occurred yet
                        intReturnCode = GetFileProductVersionAsString(strTempOSVersion, strWindowsSystemPath & "ntoskrnl.exe")
                        If intReturnCode <> 0 Then
                            boolMethodSucceeded = False
                        Else
                            intReturnCode = ConvertStringVersionNumberToMajorMinorBuildRevisionIntegers(lngTempMajor, lngTempMinor, lngTempBuild, lngTempRevision, strTempOSVersion)
                            If intReturnCode < 0 Then
                                boolMethodSucceeded = False
                            Else
                                ' Success!
                            End If
                        End If
                    End If
                End If

                If boolMethodSucceeded = False Then
                    ' C:\Windows\System32\ntoskrnl.exe failed.
                    ' Try C:\Windows\System\krnl386.exe For Windows 95, 98, or ME.
                    boolMethodSucceeded = True

                    ' Make sure we have the Windows system path
                    If TestObjectForData(strWindowsSystemPath) = True Then
                        intReturnCode = GetFileProductVersionAsString(strTempOSVersion, strWindowsSystemPath & "krnl386.exe")
                        If intReturnCode <> 0 Then
                            boolMethodSucceeded = False
                        Else
                            intReturnCode = ConvertStringVersionNumberToMajorMinorBuildRevisionIntegers(lngTempMajor, lngTempMinor, lngTempBuild, lngTempRevision, strTempOSVersion)
                            If intReturnCode < 0 Then
                                boolMethodSucceeded = False
                            Else
                                ' Success!
                            End If
                        End If
                    Else
                        ' Could not get the Windows System path; still a failure
                        boolMethodSucceeded = False
                    End If
                End If

                If boolMethodSucceeded = True Then
                    lngFunctionReturn = lngFunctionReturn + lngReturnMultiplier
                    If lngTempMajor <> CLng(-1) Then
                        ' Major version obtained
                        If intOSMajorCurrentLevelOfAccuracy < intThisMethodMajorLevelOfAccuracy Then
                            lngOSMajor = lngTempMajor
                            intOSMajorCurrentLevelOfAccuracy = intThisMethodMajorLevelOfAccuracy
                        ElseIf intOSMajorCurrentLevelOfAccuracy = intThisMethodMajorLevelOfAccuracy Then
                            If lngTempMajor > lngOSMajor Then
                                lngOSMajor = lngTempMajor
                            End If
                        End If
                    End If

                    If lngTempMinor <> CLng(-1) Then
                        ' Minor version obtained
                        If intOSMinorCurrentLevelOfAccuracy < intThisMethodMinorLevelOfAccuracy Then
                            lngOSMinor = lngTempMinor
                            intOSMinorCurrentLevelOfAccuracy = intThisMethodMinorLevelOfAccuracy
                        ElseIf intOSMinorCurrentLevelOfAccuracy = intThisMethodMinorLevelOfAccuracy Then
                            If lngTempMinor > lngOSMinor Then
                                lngOSMinor = lngTempMinor
                            End If
                        End If
                    End If

                    If lngTempBuild <> CLng(-1) Then
                        ' Build version obtained

                        ' Adjust accuracy if needed:
                        ' (If Major < 10, 7/7; If Major = 10, Minor = 0, and Build < 18362,
                        ' 7/7; else 3/7)
                        If lngOSMajor = 10 And lngOSMinor = 0 And ((lngOSBuild <> -1 And lngOSBuild >= 18362) Or (lngTempBuild <> -1 And lngTempBuild >= 18362)) Then
                            intThisMethodBuildLevelOfAccuracy = 3
                        ElseIf (lngOSMajor = 10 And lngOSMinor > 0) Or lngOSMajor > 10 Then
                            intThisMethodBuildLevelOfAccuracy = 3
                        End If

                        If intOSBuildCurrentLevelOfAccuracy < intThisMethodBuildLevelOfAccuracy Then
                            lngOSBuild = lngTempBuild
                            intOSBuildCurrentLevelOfAccuracy = intThisMethodBuildLevelOfAccuracy
                        ElseIf intOSBuildCurrentLevelOfAccuracy = intThisMethodBuildLevelOfAccuracy Then
                            If lngTempBuild > lngOSBuild Then
                                lngOSBuild = lngTempBuild
                            End If
                        End If
                    End If

                    If lngTempRevision <> CLng(-1) Then
                        ' Revision version obtained
                        If intOSRevisionCurrentLevelOfAccuracy < intThisMethodRevisionLevelOfAccuracy Then
                            lngOSRevision = lngTempRevision
                            intOSRevisionCurrentLevelOfAccuracy = intThisMethodRevisionLevelOfAccuracy
                        ElseIf intOSRevisionCurrentLevelOfAccuracy = intThisMethodRevisionLevelOfAccuracy Then
                            If lngTempRevision > lngOSRevision Then
                                lngOSRevision = lngTempRevision
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If

    ' #########################################################################################
    intThisMethodNumber = 4 ' cmd /c ver > temp file
    lngReturnMultiplier = &H00000008
    intThisMethodMajorLevelOfAccuracy = 6
    intThisMethodMinorLevelOfAccuracy = 6
    intThisMethodBuildLevelOfAccuracy = 6
    intThisMethodRevisionLevelOfAccuracy = 4
    If lngFunctionReturn >= 0 Then
        ' This needs to be set to the boolean variable that matches the method number: ########
        If boolMethod4 Then
            ' Possibly attempt method
            If intOSMajorCurrentLevelOfAccuracy < intThisMethodMajorLevelOfAccuracy Or intOSMinorCurrentLevelOfAccuracy < intThisMethodMinorLevelOfAccuracy Or intOSBuildCurrentLevelOfAccuracy < intThisMethodBuildLevelOfAccuracy Or intOSRevisionCurrentLevelOfAccuracy < intThisMethodRevisionLevelOfAccuracy Then
                ' It makes sense to attempt this method as this could improve accuracy compared
                ' to the current version information that we have
                boolMethodSucceeded = True

                ' Make the attempt and set boolMethodSucceeded = False if failure

                intReturnCode = GetWindowsOperatingSystemVersionNumberAsStringUsingCommandPromptVerCommand(strTempOSVersion)
                If intReturnCode <> 0 Then
                    boolMethodSucceeded = False
                Else
                    intReturnCode =  ConvertStringVersionNumberToMajorMinorBuildRevisionIntegers(lngTempMajor, lngTempMinor, lngTempBuild, lngTempRevision, strTempOSVersion)
                    If intReturnCode < 0 Then
                        boolMethodSucceeded = False
                    Else
                        ' Success!
                    End If
                End If

                If boolMethodSucceeded = True Then
                    lngFunctionReturn = lngFunctionReturn + lngReturnMultiplier
                    If lngTempMajor <> CLng(-1) Then
                        ' Major version obtained
                        If intOSMajorCurrentLevelOfAccuracy < intThisMethodMajorLevelOfAccuracy Then
                            lngOSMajor = lngTempMajor
                            intOSMajorCurrentLevelOfAccuracy = intThisMethodMajorLevelOfAccuracy
                        ElseIf intOSMajorCurrentLevelOfAccuracy = intThisMethodMajorLevelOfAccuracy Then
                            If lngTempMajor > lngOSMajor Then
                                lngOSMajor = lngTempMajor
                            End If
                        End If
                    End If

                    If lngTempMinor <> CLng(-1) Then
                        ' Minor version obtained
                        If intOSMinorCurrentLevelOfAccuracy < intThisMethodMinorLevelOfAccuracy Then
                            lngOSMinor = lngTempMinor
                            intOSMinorCurrentLevelOfAccuracy = intThisMethodMinorLevelOfAccuracy
                        ElseIf intOSMinorCurrentLevelOfAccuracy = intThisMethodMinorLevelOfAccuracy Then
                            If lngTempMinor > lngOSMinor Then
                                lngOSMinor = lngTempMinor
                            End If
                        End If
                    End If

                    If lngTempBuild <> CLng(-1) Then
                        ' Build version obtained
                        If intOSBuildCurrentLevelOfAccuracy < intThisMethodBuildLevelOfAccuracy Then
                            lngOSBuild = lngTempBuild
                            intOSBuildCurrentLevelOfAccuracy = intThisMethodBuildLevelOfAccuracy
                        ElseIf intOSBuildCurrentLevelOfAccuracy = intThisMethodBuildLevelOfAccuracy Then
                            If lngTempBuild > lngOSBuild Then
                                lngOSBuild = lngTempBuild
                            End If
                        End If
                    End If

                    If lngTempRevision <> CLng(-1) Then
                        ' Revision version obtained
                        If intOSRevisionCurrentLevelOfAccuracy < intThisMethodRevisionLevelOfAccuracy Then
                            lngOSRevision = lngTempRevision
                            intOSRevisionCurrentLevelOfAccuracy = intThisMethodRevisionLevelOfAccuracy
                        ElseIf intOSRevisionCurrentLevelOfAccuracy = intThisMethodRevisionLevelOfAccuracy Then
                            If lngTempRevision > lngOSRevision Then
                                lngOSRevision = lngTempRevision
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If

    ' #########################################################################################
    intThisMethodNumber = 5 ' (Version of the file C:\Windows\Sysnative\ntoskrnl.exe,
    '                         C:\Windows\System32\ntoskrnl.exe, or
    '                         C:\Windows\System\krnl386.exe)
    lngReturnMultiplier = &H00000010
    intThisMethodMajorLevelOfAccuracy = 3
    intThisMethodMinorLevelOfAccuracy = 3
    intThisMethodBuildLevelOfAccuracy = 3 ' (If Major < 10, 3/7; If Major = 10, Minor = 0, and
    '                                       Build < 18362, 3/7; else 2/7)
    intThisMethodRevisionLevelOfAccuracy = 3
    If lngFunctionReturn >= 0 Then
        ' This needs to be set to the boolean variable that matches the method number: ########
        If boolMethod5 Then
            ' Possibly attempt method
            If intOSMajorCurrentLevelOfAccuracy < intThisMethodMajorLevelOfAccuracy Or intOSMinorCurrentLevelOfAccuracy < intThisMethodMinorLevelOfAccuracy Or intOSBuildCurrentLevelOfAccuracy < intThisMethodBuildLevelOfAccuracy Or intOSRevisionCurrentLevelOfAccuracy < intThisMethodRevisionLevelOfAccuracy Then
                ' It makes sense to attempt this method as this could improve accuracy compared
                ' to the current version information that we have
                boolMethodSucceeded = True

                ' Make the attempt and set boolMethodSucceeded = False if failure

                ' First, get the Windows path if we don't already have it
                If TestObjectForData(strWindowsPath) = False Then
                    intReturnCode = GetWindowsPath(strWindowsPath)
                    If intReturnCode <> 0 Then
                        boolMethodSucceeded = False
                    End If
                End If

                If boolMethodSucceeded = True Then
                    ' No error occurred yet
                    intReturnCode = GetFileVersionAsString(strTempOSVersion, strWindowsPath & "Sysnative\ntoskrnl.exe")
                    If intReturnCode <> 0 Then
                        boolMethodSucceeded = False
                    Else
                        intReturnCode =  ConvertStringVersionNumberToMajorMinorBuildRevisionIntegers(lngTempMajor, lngTempMinor, lngTempBuild, lngTempRevision, strTempOSVersion)
                        If intReturnCode < 0 Then
                            boolMethodSucceeded = False
                        Else
                            ' Success!
                        End If
                    End If
                End If

                If boolMethodSucceeded = False Then
                    ' C:\Windows\Sysnative\ntoskrnl.exe failed.
                    ' This would be expected if we are not using Windows-on-Windows (WOW),
                    ' e.g., the current process is not 32-bit x86 on 64-bit AMD64 Windows, or
                    ' if the operating system is Windows XP or Windows Server 2003/2003 R2 and
                    ' KB942589 is missing.
                    boolMethodSucceeded = True

                    ' Get the system path if we don't already have it
                    If TestObjectForData(strWindowsSystemPath) = False Then
                        intReturnCode = GetWindowsSystemPath(strWindowsSystemPath)
                        If intReturnCode <> 0 Then
                            boolMethodSucceeded = False
                        End If
                    End If

                    If boolMethodSucceeded = True Then
                        ' No error occurred yet
                        intReturnCode = GetFileVersionAsString(strTempOSVersion, strWindowsSystemPath & "ntoskrnl.exe")
                        If intReturnCode <> 0 Then
                            boolMethodSucceeded = False
                        Else
                            intReturnCode =  ConvertStringVersionNumberToMajorMinorBuildRevisionIntegers(lngTempMajor, lngTempMinor, lngTempBuild, lngTempRevision, strTempOSVersion)
                            If intReturnCode < 0 Then
                                boolMethodSucceeded = False
                            Else
                                ' Success!
                            End If
                        End If
                    End If
                End If

                If boolMethodSucceeded = False Then
                    ' C:\Windows\System32\ntoskrnl.exe failed.
                    ' Try C:\Windows\System\krnl386.exe For Windows 95, 98, or ME.
                    boolMethodSucceeded = True

                    ' Make sure we have the Windows system path
                    If TestObjectForData(strWindowsSystemPath) = True Then
                        intReturnCode = GetFileVersionAsString(strTempOSVersion, strWindowsSystemPath & "krnl386.exe")
                        If intReturnCode <> 0 Then
                            boolMethodSucceeded = False
                        Else
                            intReturnCode =  ConvertStringVersionNumberToMajorMinorBuildRevisionIntegers(lngTempMajor, lngTempMinor, lngTempBuild, lngTempRevision, strTempOSVersion)
                            If intReturnCode < 0 Then
                                boolMethodSucceeded = False
                            Else
                                ' Success!
                            End If
                        End If
                    Else
                        ' Could not get the Windows System path; still a failure
                        boolMethodSucceeded = False
                    End If
                End If

                If boolMethodSucceeded = True Then
                    lngFunctionReturn = lngFunctionReturn + lngReturnMultiplier
                    If lngTempMajor <> CLng(-1) Then
                        ' Major version obtained
                        If intOSMajorCurrentLevelOfAccuracy < intThisMethodMajorLevelOfAccuracy Then
                            lngOSMajor = lngTempMajor
                            intOSMajorCurrentLevelOfAccuracy = intThisMethodMajorLevelOfAccuracy
                        ElseIf intOSMajorCurrentLevelOfAccuracy = intThisMethodMajorLevelOfAccuracy Then
                            If lngTempMajor > lngOSMajor Then
                                lngOSMajor = lngTempMajor
                            End If
                        End If
                    End If

                    If lngTempMinor <> CLng(-1) Then
                        ' Minor version obtained
                        If intOSMinorCurrentLevelOfAccuracy < intThisMethodMinorLevelOfAccuracy Then
                            lngOSMinor = lngTempMinor
                            intOSMinorCurrentLevelOfAccuracy = intThisMethodMinorLevelOfAccuracy
                        ElseIf intOSMinorCurrentLevelOfAccuracy = intThisMethodMinorLevelOfAccuracy Then
                            If lngTempMinor > lngOSMinor Then
                                lngOSMinor = lngTempMinor
                            End If
                        End If
                    End If

                    If lngTempBuild <> CLng(-1) Then
                        ' Build version obtained

                        ' Adjust accuracy if needed:
                        ' (If Major < 10, 3/7; If Major = 10, Minor = 0, and Build < 18362,
                        ' 3/7; else 2/7)
                        If lngOSMajor = 10 And lngOSMinor = 0 And ((lngOSBuild <> -1 And lngOSBuild >= 18362) Or (lngTempBuild <> -1 And lngTempBuild >= 18362)) Then
                            intThisMethodBuildLevelOfAccuracy = 2
                        ElseIf (lngOSMajor = 10 And lngOSMinor > 0) Or lngOSMajor > 10 Then
                            intThisMethodBuildLevelOfAccuracy = 2
                        End If

                        If intOSBuildCurrentLevelOfAccuracy < intThisMethodBuildLevelOfAccuracy Then
                            lngOSBuild = lngTempBuild
                            intOSBuildCurrentLevelOfAccuracy = intThisMethodBuildLevelOfAccuracy
                        ElseIf intOSBuildCurrentLevelOfAccuracy = intThisMethodBuildLevelOfAccuracy Then
                            If lngTempBuild > lngOSBuild Then
                                lngOSBuild = lngTempBuild
                            End If
                        End If
                    End If

                    If lngTempRevision <> CLng(-1) Then
                        ' Revision version obtained
                        If intOSRevisionCurrentLevelOfAccuracy < intThisMethodRevisionLevelOfAccuracy Then
                            lngOSRevision = lngTempRevision
                            intOSRevisionCurrentLevelOfAccuracy = intThisMethodRevisionLevelOfAccuracy
                        ElseIf intOSRevisionCurrentLevelOfAccuracy = intThisMethodRevisionLevelOfAccuracy Then
                            If lngTempRevision > lngOSRevision Then
                                lngOSRevision = lngTempRevision
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If

    ' #########################################################################################
    intThisMethodNumber = 6 '6a: "Product version" of C:\Windows\Sysnative\kernel32.dll or
    '                            C:\Windows\System32\kernel32.dll
    lngReturnMultiplier = &H00000020
    intThisMethodMajorLevelOfAccuracy = 7
    intThisMethodMinorLevelOfAccuracy = 7
    intThisMethodBuildLevelOfAccuracy = 7 ' (If Major < 10, 7/7; If Major = 10, Minor = 0, and
    '                                       Build < 18362, 7/7; else 3/7)
    intThisMethodRevisionLevelOfAccuracy = 4
    If lngFunctionReturn >= 0 Then
        ' This needs to be set to the boolean variable that matches the method number: ########
        If boolMethod6 Then
            ' Possibly attempt method
            If intOSMajorCurrentLevelOfAccuracy < intThisMethodMajorLevelOfAccuracy Or intOSMinorCurrentLevelOfAccuracy < intThisMethodMinorLevelOfAccuracy Or intOSBuildCurrentLevelOfAccuracy < intThisMethodBuildLevelOfAccuracy Or intOSRevisionCurrentLevelOfAccuracy < intThisMethodRevisionLevelOfAccuracy Then
                ' It makes sense to attempt this method as this could improve accuracy compared
                ' to the current version information that we have
                boolMethodSucceeded = True

                ' Make the attempt and set boolMethodSucceeded = False if failure

                ' First, get the Windows path if we don't already have it
                If TestObjectForData(strWindowsPath) = False Then
                    intReturnCode = GetWindowsPath(strWindowsPath)
                    If intReturnCode <> 0 Then
                        boolMethodSucceeded = False
                    End If
                End If

                If boolMethodSucceeded = True Then
                    ' No error occurred yet
                    intReturnCode = GetFileProductVersionAsString(strTempOSVersion, strWindowsPath & "Sysnative\kernel32.dll")
                    If intReturnCode <> 0 Then
                        boolMethodSucceeded = False
                    Else
                        intReturnCode =  ConvertStringVersionNumberToMajorMinorBuildRevisionIntegers(lngTempMajor, lngTempMinor, lngTempBuild, lngTempRevision, strTempOSVersion)
                        If intReturnCode < 0 Then
                            boolMethodSucceeded = False
                        Else
                            ' Success!
                        End If
                    End If
                End If

                If boolMethodSucceeded = False Then
                    ' C:\Windows\Sysnative\kernel32.dll failed.
                    ' This would be expected if we are not using Windows-on-Windows (WOW),
                    ' e.g., the current process is not 32-bit x86 on 64-bit AMD64 Windows, or
                    ' if the operating system is Windows XP or Windows Server 2003/2003 R2 and
                    ' KB942589 is missing.
                    boolMethodSucceeded = True

                    ' Get the system path if we don't already have it
                    If TestObjectForData(strWindowsSystemPath) = False Then
                        intReturnCode = GetWindowsSystemPath(strWindowsSystemPath)
                        If intReturnCode <> 0 Then
                            boolMethodSucceeded = False
                        End If
                    End If

                    If boolMethodSucceeded = True Then
                        ' No error occurred yet
                        intReturnCode = GetFileProductVersionAsString(strTempOSVersion, strWindowsSystemPath & "kernel32.dll")
                        If intReturnCode <> 0 Then
                            boolMethodSucceeded = False
                        Else
                            intReturnCode =  ConvertStringVersionNumberToMajorMinorBuildRevisionIntegers(lngTempMajor, lngTempMinor, lngTempBuild, lngTempRevision, strTempOSVersion)
                            If intReturnCode < 0 Then
                                boolMethodSucceeded = False
                            Else
                                ' Success!
                            End If
                        End If
                    End If
                End If

                If boolMethodSucceeded = True Then
                    lngFunctionReturn = lngFunctionReturn + lngReturnMultiplier
                    If lngTempMajor <> CLng(-1) Then
                        ' Major version obtained
                        If intOSMajorCurrentLevelOfAccuracy < intThisMethodMajorLevelOfAccuracy Then
                            lngOSMajor = lngTempMajor
                            intOSMajorCurrentLevelOfAccuracy = intThisMethodMajorLevelOfAccuracy
                        ElseIf intOSMajorCurrentLevelOfAccuracy = intThisMethodMajorLevelOfAccuracy Then
                            If lngTempMajor > lngOSMajor Then
                                lngOSMajor = lngTempMajor
                            End If
                        End If
                    End If

                    If lngTempMinor <> CLng(-1) Then
                        ' Minor version obtained
                        If intOSMinorCurrentLevelOfAccuracy < intThisMethodMinorLevelOfAccuracy Then
                            lngOSMinor = lngTempMinor
                            intOSMinorCurrentLevelOfAccuracy = intThisMethodMinorLevelOfAccuracy
                        ElseIf intOSMinorCurrentLevelOfAccuracy = intThisMethodMinorLevelOfAccuracy Then
                            If lngTempMinor > lngOSMinor Then
                                lngOSMinor = lngTempMinor
                            End If
                        End If
                    End If

                    If lngTempBuild <> CLng(-1) Then
                        ' Build version obtained

                        ' Adjust accuracy if needed:
                        ' (If Major < 10, 7/7; If Major = 10, Minor = 0, and Build < 18362,
                        ' 7/7; else 3/7)
                        If lngOSMajor = 10 And lngOSMinor = 0 And ((lngOSBuild <> -1 And lngOSBuild >= 18362) Or (lngTempBuild <> -1 And lngTempBuild >= 18362)) Then
                            intThisMethodBuildLevelOfAccuracy = 3
                        ElseIf (lngOSMajor = 10 And lngOSMinor > 0) Or lngOSMajor > 10 Then
                            intThisMethodBuildLevelOfAccuracy = 3
                        End If

                        If intOSBuildCurrentLevelOfAccuracy < intThisMethodBuildLevelOfAccuracy Then
                            lngOSBuild = lngTempBuild
                            intOSBuildCurrentLevelOfAccuracy = intThisMethodBuildLevelOfAccuracy
                        ElseIf intOSBuildCurrentLevelOfAccuracy = intThisMethodBuildLevelOfAccuracy Then
                            If lngTempBuild > lngOSBuild Then
                                lngOSBuild = lngTempBuild
                            End If
                        End If
                    End If

                    If lngTempRevision <> CLng(-1) Then
                        ' Revision version obtained
                        If intOSRevisionCurrentLevelOfAccuracy < intThisMethodRevisionLevelOfAccuracy Then
                            lngOSRevision = lngTempRevision
                            intOSRevisionCurrentLevelOfAccuracy = intThisMethodRevisionLevelOfAccuracy
                        ElseIf intOSRevisionCurrentLevelOfAccuracy = intThisMethodRevisionLevelOfAccuracy Then
                            If lngTempRevision > lngOSRevision Then
                                lngOSRevision = lngTempRevision
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If

    ' #########################################################################################
    intThisMethodNumber = 6 '6b: "Product version" of C:\Windows\Sysnative\ntdll.dll or
    '                            C:\Windows\System32\ntdll.dll
    lngReturnMultiplier = &H00000020
    intThisMethodMajorLevelOfAccuracy = 7
    intThisMethodMinorLevelOfAccuracy = 7
    intThisMethodBuildLevelOfAccuracy = 7 ' (If Major < 10, 7/7; If Major = 10, Minor = 0, and
    '                                       Build < 18362, 7/7; else 3/7)
    intThisMethodRevisionLevelOfAccuracy = 3
    If lngFunctionReturn >= 0 Then
        ' This needs to be set to the boolean variable that matches the method number: ########
        If boolMethod6 Then
            ' Possibly attempt method
            If intOSMajorCurrentLevelOfAccuracy < intThisMethodMajorLevelOfAccuracy Or intOSMinorCurrentLevelOfAccuracy < intThisMethodMinorLevelOfAccuracy Or intOSBuildCurrentLevelOfAccuracy < intThisMethodBuildLevelOfAccuracy Or intOSRevisionCurrentLevelOfAccuracy < intThisMethodRevisionLevelOfAccuracy Then
                ' It makes sense to attempt this method as this could improve accuracy compared
                ' to the current version information that we have
                boolMethodSucceeded = True

                ' Make the attempt and set boolMethodSucceeded = False if failure

                ' First, get the Windows path if we don't already have it
                If TestObjectForData(strWindowsPath) = False Then
                    intReturnCode = GetWindowsPath(strWindowsPath)
                    If intReturnCode <> 0 Then
                        boolMethodSucceeded = False
                    End If
                End If

                If boolMethodSucceeded = True Then
                    ' No error occurred yet
                    intReturnCode = GetFileProductVersionAsString(strTempOSVersion, strWindowsPath & "Sysnative\ntdll.dll")
                    If intReturnCode <> 0 Then
                        boolMethodSucceeded = False
                    Else
                        intReturnCode =  ConvertStringVersionNumberToMajorMinorBuildRevisionIntegers(lngTempMajor, lngTempMinor, lngTempBuild, lngTempRevision, strTempOSVersion)
                        If intReturnCode < 0 Then
                            boolMethodSucceeded = False
                        Else
                            ' Success!
                        End If
                    End If
                End If

                If boolMethodSucceeded = False Then
                    ' C:\Windows\Sysnative\ntdll.dll failed.
                    ' This would be expected if we are not using Windows-on-Windows (WOW),
                    ' e.g., the current process is not 32-bit x86 on 64-bit AMD64 Windows, or
                    ' if the operating system is Windows XP or Windows Server 2003/2003 R2 and
                    ' KB942589 is missing.
                    boolMethodSucceeded = True

                    ' Get the system path if we don't already have it
                    If TestObjectForData(strWindowsSystemPath) = False Then
                        intReturnCode = GetWindowsSystemPath(strWindowsSystemPath)
                        If intReturnCode <> 0 Then
                            boolMethodSucceeded = False
                        End If
                    End If

                    If boolMethodSucceeded = True Then
                        ' No error occurred yet
                        intReturnCode = GetFileProductVersionAsString(strTempOSVersion, strWindowsSystemPath & "ntdll.dll")
                        If intReturnCode <> 0 Then
                            boolMethodSucceeded = False
                        Else
                            intReturnCode =  ConvertStringVersionNumberToMajorMinorBuildRevisionIntegers(lngTempMajor, lngTempMinor, lngTempBuild, lngTempRevision, strTempOSVersion)
                            If intReturnCode < 0 Then
                                boolMethodSucceeded = False
                            Else
                                ' Success!
                            End If
                        End If
                    End If
                End If

                If boolMethodSucceeded = True Then
                    lngFunctionReturn = lngFunctionReturn + lngReturnMultiplier
                    If lngTempMajor <> CLng(-1) Then
                        ' Major version obtained
                        If intOSMajorCurrentLevelOfAccuracy < intThisMethodMajorLevelOfAccuracy Then
                            lngOSMajor = lngTempMajor
                            intOSMajorCurrentLevelOfAccuracy = intThisMethodMajorLevelOfAccuracy
                        ElseIf intOSMajorCurrentLevelOfAccuracy = intThisMethodMajorLevelOfAccuracy Then
                            If lngTempMajor > lngOSMajor Then
                                lngOSMajor = lngTempMajor
                            End If
                        End If
                    End If

                    If lngTempMinor <> CLng(-1) Then
                        ' Minor version obtained
                        If intOSMinorCurrentLevelOfAccuracy < intThisMethodMinorLevelOfAccuracy Then
                            lngOSMinor = lngTempMinor
                            intOSMinorCurrentLevelOfAccuracy = intThisMethodMinorLevelOfAccuracy
                        ElseIf intOSMinorCurrentLevelOfAccuracy = intThisMethodMinorLevelOfAccuracy Then
                            If lngTempMinor > lngOSMinor Then
                                lngOSMinor = lngTempMinor
                            End If
                        End If
                    End If

                    If lngTempBuild <> CLng(-1) Then
                        ' Build version obtained

                        ' Adjust accuracy if needed:
                        ' (If Major < 10, 7/7; If Major = 10, Minor = 0, and Build < 18362,
                        ' 7/7; else 3/7)
                        If lngOSMajor = 10 And lngOSMinor = 0 And ((lngOSBuild <> -1 And lngOSBuild >= 18362) Or (lngTempBuild <> -1 And lngTempBuild >= 18362)) Then
                            intThisMethodBuildLevelOfAccuracy = 3
                        ElseIf (lngOSMajor = 10 And lngOSMinor > 0) Or lngOSMajor > 10 Then
                            intThisMethodBuildLevelOfAccuracy = 3
                        End If

                        If intOSBuildCurrentLevelOfAccuracy < intThisMethodBuildLevelOfAccuracy Then
                            lngOSBuild = lngTempBuild
                            intOSBuildCurrentLevelOfAccuracy = intThisMethodBuildLevelOfAccuracy
                        ElseIf intOSBuildCurrentLevelOfAccuracy = intThisMethodBuildLevelOfAccuracy Then
                            If lngTempBuild > lngOSBuild Then
                                lngOSBuild = lngTempBuild
                            End If
                        End If
                    End If

                    If lngTempRevision <> CLng(-1) Then
                        ' Revision version obtained
                        If intOSRevisionCurrentLevelOfAccuracy < intThisMethodRevisionLevelOfAccuracy Then
                            lngOSRevision = lngTempRevision
                            intOSRevisionCurrentLevelOfAccuracy = intThisMethodRevisionLevelOfAccuracy
                        ElseIf intOSRevisionCurrentLevelOfAccuracy = intThisMethodRevisionLevelOfAccuracy Then
                            If lngTempRevision > lngOSRevision Then
                                lngOSRevision = lngTempRevision
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If

    ' #########################################################################################
    intThisMethodNumber = 6 '6c: "Product version" of C:\Windows\Sysnative\hal.dll or
    '                            C:\Windows\System32\hal.dll
    lngReturnMultiplier = &H00000020
    intThisMethodMajorLevelOfAccuracy = 7
    intThisMethodMinorLevelOfAccuracy = 7
    intThisMethodBuildLevelOfAccuracy = 7 ' (If Major < 10, 7/7; If Major = 10, Minor = 0, and
    '                                       Build < 18362, 7/7; else 3/7)
    intThisMethodRevisionLevelOfAccuracy = 2
    If lngFunctionReturn >= 0 Then
        ' This needs to be set to the boolean variable that matches the method number: ########
        If boolMethod6 Then
            ' Possibly attempt method
            If intOSMajorCurrentLevelOfAccuracy < intThisMethodMajorLevelOfAccuracy Or intOSMinorCurrentLevelOfAccuracy < intThisMethodMinorLevelOfAccuracy Or intOSBuildCurrentLevelOfAccuracy < intThisMethodBuildLevelOfAccuracy Or intOSRevisionCurrentLevelOfAccuracy < intThisMethodRevisionLevelOfAccuracy Then
                ' It makes sense to attempt this method as this could improve accuracy compared
                ' to the current version information that we have
                boolMethodSucceeded = True

                ' Make the attempt and set boolMethodSucceeded = False if failure

                ' First, get the Windows path if we don't already have it
                If TestObjectForData(strWindowsPath) = False Then
                    intReturnCode = GetWindowsPath(strWindowsPath)
                    If intReturnCode <> 0 Then
                        boolMethodSucceeded = False
                    End If
                End If

                If boolMethodSucceeded = True Then
                    ' No error occurred yet
                    intReturnCode = GetFileProductVersionAsString(strTempOSVersion, strWindowsPath & "Sysnative\hal.dll")
                    If intReturnCode <> 0 Then
                        boolMethodSucceeded = False
                    Else
                        intReturnCode =  ConvertStringVersionNumberToMajorMinorBuildRevisionIntegers(lngTempMajor, lngTempMinor, lngTempBuild, lngTempRevision, strTempOSVersion)
                        If intReturnCode < 0 Then
                            boolMethodSucceeded = False
                        Else
                            ' Success!
                        End If
                    End If
                End If

                If boolMethodSucceeded = False Then
                    ' C:\Windows\Sysnative\hal.dll failed.
                    ' This would be expected if we are not using Windows-on-Windows (WOW),
                    ' e.g., the current process is not 32-bit x86 on 64-bit AMD64 Windows, or
                    ' if the operating system is Windows XP or Windows Server 2003/2003 R2 and
                    ' KB942589 is missing.
                    boolMethodSucceeded = True

                    ' Get the system path if we don't already have it
                    If TestObjectForData(strWindowsSystemPath) = False Then
                        intReturnCode = GetWindowsSystemPath(strWindowsSystemPath)
                        If intReturnCode <> 0 Then
                            boolMethodSucceeded = False
                        End If
                    End If

                    If boolMethodSucceeded = True Then
                        ' No error occurred yet
                        intReturnCode = GetFileProductVersionAsString(strTempOSVersion, strWindowsSystemPath & "hal.dll")
                        If intReturnCode <> 0 Then
                            boolMethodSucceeded = False
                        Else
                            intReturnCode =  ConvertStringVersionNumberToMajorMinorBuildRevisionIntegers(lngTempMajor, lngTempMinor, lngTempBuild, lngTempRevision, strTempOSVersion)
                            If intReturnCode < 0 Then
                                boolMethodSucceeded = False
                            Else
                                ' Success!
                            End If
                        End If
                    End If
                End If

                If boolMethodSucceeded = True Then
                    lngFunctionReturn = lngFunctionReturn + lngReturnMultiplier
                    If lngTempMajor <> CLng(-1) Then
                        ' Major version obtained
                        If intOSMajorCurrentLevelOfAccuracy < intThisMethodMajorLevelOfAccuracy Then
                            lngOSMajor = lngTempMajor
                            intOSMajorCurrentLevelOfAccuracy = intThisMethodMajorLevelOfAccuracy
                        ElseIf intOSMajorCurrentLevelOfAccuracy = intThisMethodMajorLevelOfAccuracy Then
                            If lngTempMajor > lngOSMajor Then
                                lngOSMajor = lngTempMajor
                            End If
                        End If
                    End If

                    If lngTempMinor <> CLng(-1) Then
                        ' Minor version obtained
                        If intOSMinorCurrentLevelOfAccuracy < intThisMethodMinorLevelOfAccuracy Then
                            lngOSMinor = lngTempMinor
                            intOSMinorCurrentLevelOfAccuracy = intThisMethodMinorLevelOfAccuracy
                        ElseIf intOSMinorCurrentLevelOfAccuracy = intThisMethodMinorLevelOfAccuracy Then
                            If lngTempMinor > lngOSMinor Then
                                lngOSMinor = lngTempMinor
                            End If
                        End If
                    End If

                    If lngTempBuild <> CLng(-1) Then
                        ' Build version obtained

                        ' Adjust accuracy if needed:
                        ' (If Major < 10, 7/7; If Major = 10, Minor = 0, and Build < 18362,
                        ' 7/7; else 3/7)
                        If lngOSMajor = 10 And lngOSMinor = 0 And ((lngOSBuild <> -1 And lngOSBuild >= 18362) Or (lngTempBuild <> -1 And lngTempBuild >= 18362)) Then
                            intThisMethodBuildLevelOfAccuracy = 3
                        ElseIf (lngOSMajor = 10 And lngOSMinor > 0) Or lngOSMajor > 10 Then
                            intThisMethodBuildLevelOfAccuracy = 3
                        End If

                        If intOSBuildCurrentLevelOfAccuracy < intThisMethodBuildLevelOfAccuracy Then
                            lngOSBuild = lngTempBuild
                            intOSBuildCurrentLevelOfAccuracy = intThisMethodBuildLevelOfAccuracy
                        ElseIf intOSBuildCurrentLevelOfAccuracy = intThisMethodBuildLevelOfAccuracy Then
                            If lngTempBuild > lngOSBuild Then
                                lngOSBuild = lngTempBuild
                            End If
                        End If
                    End If

                    If lngTempRevision <> CLng(-1) Then
                        ' Revision version obtained
                        If intOSRevisionCurrentLevelOfAccuracy < intThisMethodRevisionLevelOfAccuracy Then
                            lngOSRevision = lngTempRevision
                            intOSRevisionCurrentLevelOfAccuracy = intThisMethodRevisionLevelOfAccuracy
                        ElseIf intOSRevisionCurrentLevelOfAccuracy = intThisMethodRevisionLevelOfAccuracy Then
                            If lngTempRevision > lngOSRevision Then
                                lngOSRevision = lngTempRevision
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If

    ' #########################################################################################
    intThisMethodNumber = 7 '7a: File version of C:\Windows\Sysnative\kernel32.dll or
    '                            C:\Windows\System32\kernel32.dll
    lngReturnMultiplier = &H00000040
    intThisMethodMajorLevelOfAccuracy = 3
    intThisMethodMinorLevelOfAccuracy = 3
    intThisMethodBuildLevelOfAccuracy = 3 ' (If Major < 10, 3/7; If Major = 10, Minor = 0, and
    '                                       Build < 18362, 3/7; else 2/7)
    intThisMethodRevisionLevelOfAccuracy = 2
    If lngFunctionReturn >= 0 Then
        ' This needs to be set to the boolean variable that matches the method number: ########
        If boolMethod7 Then
            ' Possibly attempt method
            If intOSMajorCurrentLevelOfAccuracy < intThisMethodMajorLevelOfAccuracy Or intOSMinorCurrentLevelOfAccuracy < intThisMethodMinorLevelOfAccuracy Or intOSBuildCurrentLevelOfAccuracy < intThisMethodBuildLevelOfAccuracy Or intOSRevisionCurrentLevelOfAccuracy < intThisMethodRevisionLevelOfAccuracy Then
                ' It makes sense to attempt this method as this could improve accuracy compared
                ' to the current version information that we have
                boolMethodSucceeded = True

                ' Make the attempt and set boolMethodSucceeded = False if failure

                ' First, get the Windows path if we don't already have it
                If TestObjectForData(strWindowsPath) = False Then
                    intReturnCode = GetWindowsPath(strWindowsPath)
                    If intReturnCode <> 0 Then
                        boolMethodSucceeded = False
                    End If
                End If

                If boolMethodSucceeded = True Then
                    ' No error occurred yet
                    intReturnCode = GetFileVersionAsString(strTempOSVersion, strWindowsPath & "Sysnative\kernel32.dll")
                    If intReturnCode <> 0 Then
                        boolMethodSucceeded = False
                    Else
                        intReturnCode =  ConvertStringVersionNumberToMajorMinorBuildRevisionIntegers(lngTempMajor, lngTempMinor, lngTempBuild, lngTempRevision, strTempOSVersion)
                        If intReturnCode < 0 Then
                            boolMethodSucceeded = False
                        Else
                            ' Success!
                        End If
                    End If
                End If

                If boolMethodSucceeded = False Then
                    ' C:\Windows\Sysnative\kernel32.dll failed.
                    ' This would be expected if we are not using Windows-on-Windows (WOW),
                    ' e.g., the current process is not 32-bit x86 on 64-bit AMD64 Windows, or
                    ' if the operating system is Windows XP or Windows Server 2003/2003 R2 and
                    ' KB942589 is missing.
                    boolMethodSucceeded = True

                    ' Get the system path if we don't already have it
                    If TestObjectForData(strWindowsSystemPath) = False Then
                        intReturnCode = GetWindowsSystemPath(strWindowsSystemPath)
                        If intReturnCode <> 0 Then
                            boolMethodSucceeded = False
                        End If
                    End If

                    If boolMethodSucceeded = True Then
                        ' No error occurred yet
                        intReturnCode = GetFileVersionAsString(strTempOSVersion, strWindowsSystemPath & "kernel32.dll")
                        If intReturnCode <> 0 Then
                            boolMethodSucceeded = False
                        Else
                            intReturnCode =  ConvertStringVersionNumberToMajorMinorBuildRevisionIntegers(lngTempMajor, lngTempMinor, lngTempBuild, lngTempRevision, strTempOSVersion)
                            If intReturnCode < 0 Then
                                boolMethodSucceeded = False
                            Else
                                ' Success!
                            End If
                        End If
                    End If
                End If

                If boolMethodSucceeded = True Then
                    lngFunctionReturn = lngFunctionReturn + lngReturnMultiplier
                    If lngTempMajor <> CLng(-1) Then
                        ' Major version obtained
                        If intOSMajorCurrentLevelOfAccuracy < intThisMethodMajorLevelOfAccuracy Then
                            lngOSMajor = lngTempMajor
                            intOSMajorCurrentLevelOfAccuracy = intThisMethodMajorLevelOfAccuracy
                        ElseIf intOSMajorCurrentLevelOfAccuracy = intThisMethodMajorLevelOfAccuracy Then
                            If lngTempMajor > lngOSMajor Then
                                lngOSMajor = lngTempMajor
                            End If
                        End If
                    End If

                    If lngTempMinor <> CLng(-1) Then
                        ' Minor version obtained
                        If intOSMinorCurrentLevelOfAccuracy < intThisMethodMinorLevelOfAccuracy Then
                            lngOSMinor = lngTempMinor
                            intOSMinorCurrentLevelOfAccuracy = intThisMethodMinorLevelOfAccuracy
                        ElseIf intOSMinorCurrentLevelOfAccuracy = intThisMethodMinorLevelOfAccuracy Then
                            If lngTempMinor > lngOSMinor Then
                                lngOSMinor = lngTempMinor
                            End If
                        End If
                    End If

                    If lngTempBuild <> CLng(-1) Then
                        ' Build version obtained

                        ' Adjust accuracy if needed:
                        ' (If Major < 10, 3/7; If Major = 10, Minor = 0, and Build < 18362,
                        ' 3/7; else 2/7)
                        If lngOSMajor = 10 And lngOSMinor = 0 And ((lngOSBuild <> -1 And lngOSBuild >= 18362) Or (lngTempBuild <> -1 And lngTempBuild >= 18362)) Then
                            intThisMethodBuildLevelOfAccuracy = 2
                        ElseIf (lngOSMajor = 10 And lngOSMinor > 0) Or lngOSMajor > 10 Then
                            intThisMethodBuildLevelOfAccuracy = 2
                        End If

                        If intOSBuildCurrentLevelOfAccuracy < intThisMethodBuildLevelOfAccuracy Then
                            lngOSBuild = lngTempBuild
                            intOSBuildCurrentLevelOfAccuracy = intThisMethodBuildLevelOfAccuracy
                        ElseIf intOSBuildCurrentLevelOfAccuracy = intThisMethodBuildLevelOfAccuracy Then
                            If lngTempBuild > lngOSBuild Then
                                lngOSBuild = lngTempBuild
                            End If
                        End If
                    End If

                    If lngTempRevision <> CLng(-1) Then
                        ' Revision version obtained
                        If intOSRevisionCurrentLevelOfAccuracy < intThisMethodRevisionLevelOfAccuracy Then
                            lngOSRevision = lngTempRevision
                            intOSRevisionCurrentLevelOfAccuracy = intThisMethodRevisionLevelOfAccuracy
                        ElseIf intOSRevisionCurrentLevelOfAccuracy = intThisMethodRevisionLevelOfAccuracy Then
                            If lngTempRevision > lngOSRevision Then
                                lngOSRevision = lngTempRevision
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If

    ' #########################################################################################
    intThisMethodNumber = 7 '7b: File version of C:\Windows\Sysnative\ntdll.dll or
    '                            C:\Windows\System32\ntdll.dll
    lngReturnMultiplier = &H00000040
    intThisMethodMajorLevelOfAccuracy = 3
    intThisMethodMinorLevelOfAccuracy = 3
    intThisMethodBuildLevelOfAccuracy = 3 ' (If Major < 10, 3/7; If Major = 10, Minor = 0, and
    '                                       Build < 18362, 3/7; else 2/7)
    intThisMethodRevisionLevelOfAccuracy = 1
    If lngFunctionReturn >= 0 Then
        ' This needs to be set to the boolean variable that matches the method number: ########
        If boolMethod7 Then
            ' Possibly attempt method
            If intOSMajorCurrentLevelOfAccuracy < intThisMethodMajorLevelOfAccuracy Or intOSMinorCurrentLevelOfAccuracy < intThisMethodMinorLevelOfAccuracy Or intOSBuildCurrentLevelOfAccuracy < intThisMethodBuildLevelOfAccuracy Or intOSRevisionCurrentLevelOfAccuracy < intThisMethodRevisionLevelOfAccuracy Then
                ' It makes sense to attempt this method as this could improve accuracy compared
                ' to the current version information that we have
                boolMethodSucceeded = True

                ' Make the attempt and set boolMethodSucceeded = False if failure

                ' First, get the Windows path if we don't already have it
                If TestObjectForData(strWindowsPath) = False Then
                    intReturnCode = GetWindowsPath(strWindowsPath)
                    If intReturnCode <> 0 Then
                        boolMethodSucceeded = False
                    End If
                End If

                If boolMethodSucceeded = True Then
                    ' No error occurred yet
                    intReturnCode = GetFileVersionAsString(strTempOSVersion, strWindowsPath & "Sysnative\ntdll.dll")
                    If intReturnCode <> 0 Then
                        boolMethodSucceeded = False
                    Else
                        intReturnCode =  ConvertStringVersionNumberToMajorMinorBuildRevisionIntegers(lngTempMajor, lngTempMinor, lngTempBuild, lngTempRevision, strTempOSVersion)
                        If intReturnCode < 0 Then
                            boolMethodSucceeded = False
                        Else
                            ' Success!
                        End If
                    End If
                End If

                If boolMethodSucceeded = False Then
                    ' C:\Windows\Sysnative\ntdll.dll failed.
                    ' This would be expected if we are not using Windows-on-Windows (WOW),
                    ' e.g., the current process is not 32-bit x86 on 64-bit AMD64 Windows, or
                    ' if the operating system is Windows XP or Windows Server 2003/2003 R2 and
                    ' KB942589 is missing.
                    boolMethodSucceeded = True

                    ' Get the system path if we don't already have it
                    If TestObjectForData(strWindowsSystemPath) = False Then
                        intReturnCode = GetWindowsSystemPath(strWindowsSystemPath)
                        If intReturnCode <> 0 Then
                            boolMethodSucceeded = False
                        End If
                    End If

                    If boolMethodSucceeded = True Then
                        ' No error occurred yet
                        intReturnCode = GetFileVersionAsString(strTempOSVersion, strWindowsSystemPath & "ntdll.dll")
                        If intReturnCode <> 0 Then
                            boolMethodSucceeded = False
                        Else
                            intReturnCode =  ConvertStringVersionNumberToMajorMinorBuildRevisionIntegers(lngTempMajor, lngTempMinor, lngTempBuild, lngTempRevision, strTempOSVersion)
                            If intReturnCode < 0 Then
                                boolMethodSucceeded = False
                            Else
                                ' Success!
                            End If
                        End If
                    End If
                End If

                If boolMethodSucceeded = True Then
                    lngFunctionReturn = lngFunctionReturn + lngReturnMultiplier
                    If lngTempMajor <> CLng(-1) Then
                        ' Major version obtained
                        If intOSMajorCurrentLevelOfAccuracy < intThisMethodMajorLevelOfAccuracy Then
                            lngOSMajor = lngTempMajor
                            intOSMajorCurrentLevelOfAccuracy = intThisMethodMajorLevelOfAccuracy
                        ElseIf intOSMajorCurrentLevelOfAccuracy = intThisMethodMajorLevelOfAccuracy Then
                            If lngTempMajor > lngOSMajor Then
                                lngOSMajor = lngTempMajor
                            End If
                        End If
                    End If

                    If lngTempMinor <> CLng(-1) Then
                        ' Minor version obtained
                        If intOSMinorCurrentLevelOfAccuracy < intThisMethodMinorLevelOfAccuracy Then
                            lngOSMinor = lngTempMinor
                            intOSMinorCurrentLevelOfAccuracy = intThisMethodMinorLevelOfAccuracy
                        ElseIf intOSMinorCurrentLevelOfAccuracy = intThisMethodMinorLevelOfAccuracy Then
                            If lngTempMinor > lngOSMinor Then
                                lngOSMinor = lngTempMinor
                            End If
                        End If
                    End If

                    If lngTempBuild <> CLng(-1) Then
                        ' Build version obtained

                        ' Adjust accuracy if needed:
                        ' (If Major < 10, 3/7; If Major = 10, Minor = 0, and Build < 18362,
                        ' 3/7; else 2/7)
                        If lngOSMajor = 10 And lngOSMinor = 0 And ((lngOSBuild <> -1 And lngOSBuild >= 18362) Or (lngTempBuild <> -1 And lngTempBuild >= 18362)) Then
                            intThisMethodBuildLevelOfAccuracy = 2
                        ElseIf (lngOSMajor = 10 And lngOSMinor > 0) Or lngOSMajor > 10 Then
                            intThisMethodBuildLevelOfAccuracy = 2
                        End If

                        If intOSBuildCurrentLevelOfAccuracy < intThisMethodBuildLevelOfAccuracy Then
                            lngOSBuild = lngTempBuild
                            intOSBuildCurrentLevelOfAccuracy = intThisMethodBuildLevelOfAccuracy
                        ElseIf intOSBuildCurrentLevelOfAccuracy = intThisMethodBuildLevelOfAccuracy Then
                            If lngTempBuild > lngOSBuild Then
                                lngOSBuild = lngTempBuild
                            End If
                        End If
                    End If

                    If lngTempRevision <> CLng(-1) Then
                        ' Revision version obtained
                        If intOSRevisionCurrentLevelOfAccuracy < intThisMethodRevisionLevelOfAccuracy Then
                            lngOSRevision = lngTempRevision
                            intOSRevisionCurrentLevelOfAccuracy = intThisMethodRevisionLevelOfAccuracy
                        ElseIf intOSRevisionCurrentLevelOfAccuracy = intThisMethodRevisionLevelOfAccuracy Then
                            If lngTempRevision > lngOSRevision Then
                                lngOSRevision = lngTempRevision
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If

    ' #########################################################################################
    intThisMethodNumber = 7 '7c: File version of C:\Windows\Sysnative\hal.dll or
    '                            C:\Windows\System32\hal.dll
    lngReturnMultiplier = &H00000040
    intThisMethodMajorLevelOfAccuracy = 3
    intThisMethodMinorLevelOfAccuracy = 3
    intThisMethodBuildLevelOfAccuracy = 3 ' (If Major < 10, 3/7; If Major = 10, Minor = 0, and
    '                                       Build < 18362, 3/7; else 2/7)
    intThisMethodRevisionLevelOfAccuracy = 1
    If lngFunctionReturn >= 0 Then
        ' This needs to be set to the boolean variable that matches the method number: ########
        If boolMethod7 Then
            ' Possibly attempt method
            If intOSMajorCurrentLevelOfAccuracy < intThisMethodMajorLevelOfAccuracy Or intOSMinorCurrentLevelOfAccuracy < intThisMethodMinorLevelOfAccuracy Or intOSBuildCurrentLevelOfAccuracy < intThisMethodBuildLevelOfAccuracy Or intOSRevisionCurrentLevelOfAccuracy < intThisMethodRevisionLevelOfAccuracy Then
                ' It makes sense to attempt this method as this could improve accuracy compared
                ' to the current version information that we have
                boolMethodSucceeded = True

                ' Make the attempt and set boolMethodSucceeded = False if failure

                ' First, get the Windows path if we don't already have it
                If TestObjectForData(strWindowsPath) = False Then
                    intReturnCode = GetWindowsPath(strWindowsPath)
                    If intReturnCode <> 0 Then
                        boolMethodSucceeded = False
                    End If
                End If

                If boolMethodSucceeded = True Then
                    ' No error occurred yet
                    intReturnCode = GetFileVersionAsString(strTempOSVersion, strWindowsPath & "Sysnative\hal.dll")
                    If intReturnCode <> 0 Then
                        boolMethodSucceeded = False
                    Else
                        intReturnCode =  ConvertStringVersionNumberToMajorMinorBuildRevisionIntegers(lngTempMajor, lngTempMinor, lngTempBuild, lngTempRevision, strTempOSVersion)
                        If intReturnCode < 0 Then
                            boolMethodSucceeded = False
                        Else
                            ' Success!
                        End If
                    End If
                End If

                If boolMethodSucceeded = False Then
                    ' C:\Windows\Sysnative\hal.dll failed.
                    ' This would be expected if we are not using Windows-on-Windows (WOW),
                    ' e.g., the current process is not 32-bit x86 on 64-bit AMD64 Windows, or
                    ' if the operating system is Windows XP or Windows Server 2003/2003 R2 and
                    ' KB942589 is missing.
                    boolMethodSucceeded = True

                    ' Get the system path if we don't already have it
                    If TestObjectForData(strWindowsSystemPath) = False Then
                        intReturnCode = GetWindowsSystemPath(strWindowsSystemPath)
                        If intReturnCode <> 0 Then
                            boolMethodSucceeded = False
                        End If
                    End If

                    If boolMethodSucceeded = True Then
                        ' No error occurred yet
                        intReturnCode = GetFileVersionAsString(strTempOSVersion, strWindowsSystemPath & "hal.dll")
                        If intReturnCode <> 0 Then
                            boolMethodSucceeded = False
                        Else
                            intReturnCode =  ConvertStringVersionNumberToMajorMinorBuildRevisionIntegers(lngTempMajor, lngTempMinor, lngTempBuild, lngTempRevision, strTempOSVersion)
                            If intReturnCode < 0 Then
                                boolMethodSucceeded = False
                            Else
                                ' Success!
                            End If
                        End If
                    End If
                End If

                If boolMethodSucceeded = True Then
                    lngFunctionReturn = lngFunctionReturn + lngReturnMultiplier ' TODO: This approach needs to be updated, I think, if multiple 7x methods took place?
                    If lngTempMajor <> CLng(-1) Then
                        ' Major version obtained
                        If intOSMajorCurrentLevelOfAccuracy < intThisMethodMajorLevelOfAccuracy Then
                            lngOSMajor = lngTempMajor
                            intOSMajorCurrentLevelOfAccuracy = intThisMethodMajorLevelOfAccuracy
                        ElseIf intOSMajorCurrentLevelOfAccuracy = intThisMethodMajorLevelOfAccuracy Then
                            If lngTempMajor > lngOSMajor Then
                                lngOSMajor = lngTempMajor
                            End If
                        End If
                    End If

                    If lngTempMinor <> CLng(-1) Then
                        ' Minor version obtained
                        If intOSMinorCurrentLevelOfAccuracy < intThisMethodMinorLevelOfAccuracy Then
                            lngOSMinor = lngTempMinor
                            intOSMinorCurrentLevelOfAccuracy = intThisMethodMinorLevelOfAccuracy
                        ElseIf intOSMinorCurrentLevelOfAccuracy = intThisMethodMinorLevelOfAccuracy Then
                            If lngTempMinor > lngOSMinor Then
                                lngOSMinor = lngTempMinor
                            End If
                        End If
                    End If

                    If lngTempBuild <> CLng(-1) Then
                        ' Build version obtained

                        ' Adjust accuracy if needed:
                        ' (If Major < 10, 3/7; If Major = 10, Minor = 0, and Build < 18362,
                        ' 3/7; else 2/7)
                        If lngOSMajor = 10 And lngOSMinor = 0 And ((lngOSBuild <> -1 And lngOSBuild >= 18362) Or (lngTempBuild <> -1 And lngTempBuild >= 18362)) Then
                            intThisMethodBuildLevelOfAccuracy = 2
                        ElseIf (lngOSMajor = 10 And lngOSMinor > 0) Or lngOSMajor > 10 Then
                            intThisMethodBuildLevelOfAccuracy = 2
                        End If

                        If intOSBuildCurrentLevelOfAccuracy < intThisMethodBuildLevelOfAccuracy Then
                            lngOSBuild = lngTempBuild
                            intOSBuildCurrentLevelOfAccuracy = intThisMethodBuildLevelOfAccuracy
                        ElseIf intOSBuildCurrentLevelOfAccuracy = intThisMethodBuildLevelOfAccuracy Then
                            If lngTempBuild > lngOSBuild Then
                                lngOSBuild = lngTempBuild
                            End If
                        End If
                    End If

                    If lngTempRevision <> CLng(-1) Then
                        ' Revision version obtained
                        If intOSRevisionCurrentLevelOfAccuracy < intThisMethodRevisionLevelOfAccuracy Then
                            lngOSRevision = lngTempRevision
                            intOSRevisionCurrentLevelOfAccuracy = intThisMethodRevisionLevelOfAccuracy
                        ElseIf intOSRevisionCurrentLevelOfAccuracy = intThisMethodRevisionLevelOfAccuracy Then
                            If lngTempRevision > lngOSRevision Then
                                lngOSRevision = lngTempRevision
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If

    ' ' Windows NT 4.0, Windows 2000 and Newer OS Kernel:
    ' strFilePath = "C:\Windows\System32\ntoskrnl.exe"
    ' intReturnCode = GetFileProductVersionAsString(strFileProductVersion, strFilePath)
    ' WScript.Echo intReturnCode
    ' WScript.Echo strFileProductVersion

    ' ' Windows 9x OS Kernel:
    ' strFilePath = "C:\Windows\System\krnl386.exe"
    ' intReturnCode = GetFileProductVersionAsString(strFileProductVersion, strFilePath)
    ' WScript.Echo intReturnCode
    ' WScript.Echo strFileProductVersion

    ' strFilePath = "C:\Windows\System32\kernel32.dll"
    ' intReturnCode = GetFileProductVersionAsString(strFileProductVersion, strFilePath)
    ' WScript.Echo intReturnCode
    ' WScript.Echo strFileProductVersion

    ' strFilePath = "C:\Windows\System32\ntdll.dll"
    ' intReturnCode = GetFileProductVersionAsString(strFileProductVersion, strFilePath)
    ' WScript.Echo intReturnCode
    ' WScript.Echo strFileProductVersion

    ' strFilePath = "C:\Windows\System32\hal.dll"
    ' intReturnCode = GetFileProductVersionAsString(strFileProductVersion, strFilePath)
    ' WScript.Echo intReturnCode
    ' WScript.Echo strFileProductVersion
    '
    ' Fifth choice: file version of the above files (not guaranteed accurate)
    ' See https://devblogs.microsoft.com/scripting/how-can-i-determine-the-version-number-of-a-file/
    ' strFilePath = "C:\Windows\System32\ntoskrnl.exe"
    ' intReturnCode = GetFileVersionAsString(strFileVersion, strFilePath)
    ' WScript.Echo intReturnCode
    ' WScript.Echo strFileVersion

    ' ' Windows 9x OS Kernel:
    ' strFilePath = "C:\Windows\System\krnl386.exe"
    ' intReturnCode = GetFileVersionAsString(strFileVersion, strFilePath)
    ' WScript.Echo intReturnCode
    ' WScript.Echo strFileVersion

    ' strFilePath = "C:\Windows\System32\kernel32.dll"
    ' intReturnCode = GetFileVersionAsString(strFileVersion, strFilePath)
    ' WScript.Echo intReturnCode
    ' WScript.Echo strFileVersion

    ' strFilePath = "C:\Windows\System32\ntdll.dll"
    ' intReturnCode = GetFileVersionAsString(strFileVersion, strFilePath)
    ' WScript.Echo intReturnCode
    ' WScript.Echo strFileVersion

    ' strFilePath = "C:\Windows\System32\hal.dll"
    ' intReturnCode = GetFileVersionAsString(strFileVersion, strFilePath)
    ' WScript.Echo intReturnCode
    ' WScript.Echo strFileVersion
    '
    ' Sixth choice: registry: https://superuser.com/a/1160428/334370
    '
    ' Other reference:
    '   https://gist.github.com/SevenLayerJedi/c0415c03cab1ff51aa49d2f4d708f265

    If lngFunctionReturn >= 0 Then
        ' No error has occurred
        ' Check for failures to deliver required accuracy levels
        If intWorkingAccuracyTargetForMajor >= 0 Then
            lngReturnMultiplier = &H00000800
            If intOSMajorCurrentLevelOfAccuracy = -1 Or intOSMajorCurrentLevelOfAccuracy < intWorkingAccuracyTargetForMajor Then
                If lngFunctionReturn > 0 Then
                    lngFunctionReturn = 0
                End If
                lngFunctionReturn = lngFunctionReturn + (-1 * lngReturnMultiplier)
            End If
        End If

        If intWorkingAccuracyTargetForMinor >= 0 Then
            lngReturnMultiplier = &H00000400
            If intOSMinorCurrentLevelOfAccuracy = -1 Or intOSMinorCurrentLevelOfAccuracy < intWorkingAccuracyTargetForMinor Then
                If lngFunctionReturn > 0 Then
                    lngFunctionReturn = 0
                End If
                lngFunctionReturn = lngFunctionReturn + (-1 * lngReturnMultiplier)
            End If
        End If

        If intWorkingAccuracyTargetForBuild >= 1 Then
            lngReturnMultiplier = &H00000200
            If intOSBuildCurrentLevelOfAccuracy = -1 Or intOSBuildCurrentLevelOfAccuracy < intWorkingAccuracyTargetForBuild Then
                If lngFunctionReturn > 0 Then
                    lngFunctionReturn = 0
                End If
                lngFunctionReturn = lngFunctionReturn + (-1 * lngReturnMultiplier)
            End If
        End If

        If intWorkingAccuracyTargetForRevision >= 1 Then
            lngReturnMultiplier = &H00000100
            If intOSRevisionCurrentLevelOfAccuracy = -1 Or intOSRevisionCurrentLevelOfAccuracy < intWorkingAccuracyTargetForRevision Then
                If lngFunctionReturn > 0 Then
                    lngFunctionReturn = 0
                End If
                lngFunctionReturn = lngFunctionReturn + (-1 * lngReturnMultiplier)
            End If
        End If
    End If

    If lngFunctionReturn >= 0 Then
        ' No error has occurred
        ' Build positive return code that conveys accuracy levels
        lngReturnMultiplier = &H01000000
        If intOSMajorCurrentLevelOfAccuracy = -1 Then
            intTempLevelOfAccuracy = 0
        Else
            intTempLevelOfAccuracy = intOSMajorCurrentLevelOfAccuracy
        End If
        lngFunctionReturn = lngFunctionReturn + (lngReturnMultiplier * intTempLevelOfAccuracy)

        lngReturnMultiplier = &H00100000
        If intOSMinorCurrentLevelOfAccuracy = -1 Then
            intTempLevelOfAccuracy = 0
        Else
            intTempLevelOfAccuracy = intOSMinorCurrentLevelOfAccuracy
        End If
        lngFunctionReturn = lngFunctionReturn + (lngReturnMultiplier * intTempLevelOfAccuracy)

        lngReturnMultiplier = &H00010000
        If intOSBuildCurrentLevelOfAccuracy = -1 Then
            intTempLevelOfAccuracy = 0
        Else
            intTempLevelOfAccuracy = intOSBuildCurrentLevelOfAccuracy
        End If
        lngFunctionReturn = lngFunctionReturn + (lngReturnMultiplier * intTempLevelOfAccuracy)
        
        lngReturnMultiplier = &H00001000
        If intOSRevisionCurrentLevelOfAccuracy = -1 Then
            intTempLevelOfAccuracy = 0
        Else
            intTempLevelOfAccuracy = intOSRevisionCurrentLevelOfAccuracy
        End If
        lngFunctionReturn = lngFunctionReturn + (lngReturnMultiplier * intTempLevelOfAccuracy)
    End If

    If lngFunctionReturn >= 0 Then
        ' No error has occurred
        ' Build version string to return
        If intWorkingAccuracyTargetForMajor = -1 Then
            ' Build or Revision, only
            If intWorkingAccuracyTargetForBuild >= 1 Then
                strOSVersionStaging = CStr(lngOSBuild)
            Else
                strOSVersionStaging = CStr(lngOSRevision)
            End If
        Else
            strOSVersionStaging = CStr(lngOSMajor) & "." & CStr(lngOSMinor)
            If intOSBuildCurrentLevelOfAccuracy >= 1 Then
                strOSVersionStaging = strOSVersionStaging & "." & CStr(lngOSBuild)
                If intOSRevisionCurrentLevelOfAccuracy >= 1 Then
                    strOSVersionStaging = strOSVersionStaging & "." & CStr(lngOSRevision)
                End If
            End If
        End If
    End If

    If lngFunctionReturn >= 0 Then
        strOperatingSystemVersion = strOSVersionStaging
    End If

    GetWindowsOperatingSystemVersionNumberAsString = lngFunctionReturn
End Function
