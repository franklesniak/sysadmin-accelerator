' Versions of VBScript prior to 5.0 do not support the class keyword; suppress errors:
Err.Clear
On Error Resume Next

Class Version
    'region VersionClassMetadata ####################################################
    ' Implements a VBScript version of the .NET System.Version class. Useful because .NET
    ' objects are not readily accessible in VBScript, and version-processing/comparison is a
    ' common systems administration activity.
    '
    ' Version: 1.1.20210613.1
    '
    ' Public Methods:
    '   Clone(ByRef objTargetVersionObject)
    '   CompareTo(ByVal objOtherVersionObject)
    '   CompareToString(ByVal strOtherVersion)
    '   Equals(ByVal objOtherVersionObject)
    '   GreaterThan(ByVal objOtherVersionObject)
    '   GreaterThanOrEqual(ByVal objOtherVersionObject)
    '   InitFromMajorMinor(ByVal lngMajor, ByVal lngMinor)
    '   InitFromMajorMinorBuild(ByVal lngMajor, ByVal lngMinor, ByVal lngBuild)
    '   InitFromMajorMinorBuildRevision(ByVal lngMajor, ByVal lngMinor, ByVal lngBuild,
    '       ByVal lngRevision)
    '   InitFromString(ByVal strVersion)
    '   LessThan(ByVal objOtherVersionObject)
    '   LessThanOrEqual(ByVal objOtherVersionObject)
    '   NotEquals(ByVal objOtherVersionObject)
    '   ToString()
    '
    ' Public Properties:
    '   Major (get)
    '   Minor (get)
    '   Build (get)
    '   Revision (get)
    '   MajorRevision (get)
    '   MinorRevision (get)
    '
    ' Not implemented:
    '   GetHashCode
    '   Parse (see InitFromString method)
    '   TryFormat (see ToString method)
    '   TryParse (see InitFromString method)
    '
    ' Note: the creation of a class such as this one requires VBScript 5.0, which is included
    ' in Internet Explorer 5.0 and was made available as a standalone download. One can also
    ' install Windows Scripting Host 2.0, which includes VBScript 5.1 and is compatible.
    ' Previous versions of VBScript (e.g., VBScript 3.0, included in Internet Explorer 4, IIS
    ' 4, Outlook 98, and Windows Scripting Host 1.0) are not compatible.
    '
    ' Example 1:
    ' Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
    ' Set colItems = objWMI.ExecQuery("Select Version from Win32_OperatingSystem")
    ' For Each objItem in colItems
    '   strOSString = objItem.Version
    ' Next
    ' Set versionOperatingSystem = New Version
    ' intReturnCode = versionOperatingSystem.InitFromString(strOSString)
    ' If intReturnCode = 0 Then
    '   ' Success
    '   If versionOperatingSystem.CompareToString("10.0") >= 0 Then
    '       WScript.Echo("Windows 10, Windows Server 2016, or newer!")
    '   Else
    '       WScript.Echo("Windows 8.1, Windows Server 2012 R2, or older!")
    '   End If
    ' Else
    '   WScript.Echo("An error occurred reading the OS version.")
    ' End If
    '
    ' Example 2:
    ' Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
    ' Set colItems = objWMI.ExecQuery("Select Version from Win32_OperatingSystem")
    ' For Each objItem in colItems
    '   strOSString = objItem.Version
    ' Next
    ' Set versionCurrentOperatingSystem = New Version
    ' intReturnCode = versionCurrentOperatingSystem.InitFromString(strOSString)
    ' If intReturnCode <> 0 Then
    '   WScript.Echo("Failed to get the current operating system version!")
    ' End If
    ' Set versionWindows98 = New Version
    ' intReturnCode = versionWindows98.InitFromMajorMinorBuild(4,10,1998)
    ' Set versionWindows98SE = New Version
    ' intReturnCode = versionWindows98SE.InitFromMajorMinorBuild(4,10,2222)
    ' Set versionWindowsME = New Version
    ' intReturnCode = versionWindowsME.InitFromMajorMinor(4,90)
    ' bool9x = False
    ' If versionCurrentOperatingSystem.GreaterThanOrEqual(versionWindows98) And versionCurrentOperatingSystem.LessThanOrEqual(versionWindows98SE) Then
    '   bool9x = True
    ' ElseIf (versionCurrentOperatingSystem.Major = versionWindowsME.Major) And (versionCurrentOperatingSystem.Minor = versionWindowsME.Minor) Then
    '   bool9x = True
    ' End If
    ' If bool9x Then
    '   WScript.Echo("Current OS is Windows 9x. It's 2020 (or later). What are you thinking?")
    ' Else
    '   WScript.Echo("Thank the maker! This OS is not Windows 9x.")
    ' End If
    'endregion VersionClassMetadata ####################################################

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
    ' at https://github.com/franklesniak/dotNet_System_Version_on_VBScript
    'endregion DownloadLocationNotice ####################################################

    'region Acknowledgements ####################################################
    ' Andrew Clinick, for writing the MSDN article "Clinick's Clinic on Scripting: Take Five
    ' What's New in the Version 5.0 Script Engines" - which confirmed that a VBScript class
    ' requires 5.0 of the script engine.
    '
    ' Jerry Lee Ford, Jr., for providing a history of VBScript and Windows Scripting Host in
    ' his book, "Microsoft WSH and VBScript Programming for the Absolute Beginner".
    '
    ' Gunter Born, for providing a history of Windows Scripting Host in his book "Microsoft
    ' Windows Script Host 2.0 Developer's Guide" that corrected some points.
    'endregion Acknowledgements ####################################################

    'region DependsOn ####################################################
    ' None - this class is entirely self-contained. However, this class contains a private
    ' function TestObjectForData() that should be identical to the public TestObjectForData()
    'endregion DependsOn ####################################################

    Private lngPrivateMajor
    Private lngPrivateMinor
    Private lngPrivateBuild
    Private lngPrivateRevision
    
    Private Sub Class_Initialize()
        lngPrivateMajor = CLng(0)
        lngPrivateMinor = CLng(0)
        lngPrivateBuild = CLng(-1)
        lngPrivateRevision = CLng(-1)
    End Sub

    Private Function TestObjectForData(ByVal objToCheck)
        'region FunctionMetadata ####################################################
        ' Checks an object or variable to see if it "has data".
        ' If any of the following are true, then objToCheck is regarded as NOT having data:
        '   VarType(objToCheck) = 0
        '   VarType(objToCheck) = 1
        '   objToCheck Is Nothing
        '   IsEmpty(objToCheck)
        '   IsNull(objToCheck)
        '   objToCheck = vbNullString (or "")
        '   IsArray(objToCheck) = True And UBound(objToCheck) throws an error
        '   IsArray(objToCheck) = True And UBound(objToCheck) < 0
        ' In any of these cases, the function returns False. Otherwise, it returns True.
        '
        ' Version: 1.1.20210115.0
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
        ' at https://github.com/franklesniak/Test_Object_For_Data
        'endregion DownloadLocationNotice ####################################################
    
        'region Acknowledgements ####################################################
        ' Thanks to Scott Dexter for writing the article "Empty Nothing And Null How Do You Feel
        ' Today", which inspired me to create this function. https://evolt.org/node/346
        '
        ' Thanks also to "RhinoScript" for the article "Testing for Empty Arrays" for providing
        ' guidance for how to test for the empty array condition in VBScript.
        ' https://wiki.mcneel.com/developer/scriptsamples/emptyarray
        '
        ' Thanks also "iamresearcher" who posted this and inspired the test case for vbNullString:
        ' https://www.vbforums.com/showthread.php?684799-The-Differences-among-Empty-Nothing-vbNull-vbNullChar-vbNullString-and-the-Zero-L
        'endregion Acknowledgements ####################################################
    
        Dim boolTestResult
        Dim boolFunctionReturn
        Dim intArrayUBound
    
        Err.Clear
    
        boolFunctionReturn = True
    
        'Check VarType(objToCheck) = 0
        On Error Resume Next
        boolTestResult = (VarType(objToCheck) = 0)
        If Err Then
            'Error occurred
            Err.Clear
            On Error Goto 0
        Else
            'No Error
            On Error Goto 0
            If boolTestResult = True Then
                'vbEmpty
                boolFunctionReturn = False
            End If
        End If
    
        'Check VarType(objToCheck) = 1
        On Error Resume Next
        boolTestResult = (VarType(objToCheck) = 1)
        If Err Then
            'Error occurred
            Err.Clear
            On Error Goto 0
        Else
            'No Error
            On Error Goto 0
            If boolTestResult = True Then
                'vbNull
                boolFunctionReturn = False
            End If
        End If
    
        'Check to see if objToCheck Is Nothing
        If boolFunctionReturn = True Then
            On Error Resume Next
            boolTestResult = (objToCheck Is Nothing)
            If Err Then
                'Error occurred
                Err.Clear
                On Error Goto 0
            Else
                'No Error
                On Error Goto 0
                If boolTestResult = True Then
                    'No data
                    boolFunctionReturn = False
                End If
            End If
        End If
    
        'Check IsEmpty(objToCheck)
        If boolFunctionReturn = True Then
            On Error Resume Next
            boolTestResult = IsEmpty(objToCheck)
            If Err Then
                'Error occurred
                Err.Clear
                On Error Goto 0
            Else
                'No Error
                On Error Goto 0
                If boolTestResult = True Then
                    'No data
                    boolFunctionReturn = False
                End If
            End If
        End If
    
        'Check IsNull(objToCheck)
        If boolFunctionReturn = True Then
            On Error Resume Next
            boolTestResult = IsNull(objToCheck)
            If Err Then
                'Error occurred
                Err.Clear
                On Error Goto 0
            Else
                'No Error
                On Error Goto 0
                If boolTestResult = True Then
                    'No data
                    boolFunctionReturn = False
                End If
            End If
        End If
        
        'Check objToCheck = vbNullString
        If boolFunctionReturn = True Then
            On Error Resume Next
            boolTestResult = (objToCheck = vbNullString)
            If Err Then
                'Error occurred
                Err.Clear
                On Error Goto 0
            Else
                'No Error
                On Error Goto 0
                If boolTestResult = True Then
                    'No data
                    boolFunctionReturn = False
                End If
            End If
        End If
    
        If boolFunctionReturn = True Then
            On Error Resume Next
            boolTestResult = IsArray(objToCheck)
            If Err Then
                'Error occurred
                Err.Clear
                On Error Goto 0
                boolTestResult = False
            Else
                'No Error
                On Error Goto 0
            End If
            If boolTestResult = True Then
                ' objToCheck is an array
                On Error Resume Next
                intArrayUBound = UBound(objToCheck)
                If Err Then
                    'Undimensioned array
                    Err.Clear
                    On Error Goto 0
                    intArrayUBound = -1
                Else
                    On Error Goto 0
                End If
                If intArrayUBound < 0 Then
                    boolFunctionReturn = False
                End If
            End If
        End If
    
        TestObjectForData = boolFunctionReturn
    End Function        

    Public Function Clone(ByRef objTargetVersionObject)
        ' Creates a copy of the current version object and stores it in the first (and only)
        ' argument supplied to this function

        ' Returns 0 if successful; non-zero otherwise

        Dim intReturnCode
        intReturnCode = 0
        Set objTargetVersionObject = New Version
        If lngPrivateRevision = CLng(-1) Then
            If lngPrivateBuild = CLng(-1) Then
                ' Initialize with Major/Minor only
                intReturnCode = objTargetVersionObject.InitFromMajorMinor(lngPrivateMajor, lngPrivateMinor)
            Else
                ' Initialize with Major/Minor/Build only
                intReturnCode = objTargetVersionObject.InitFromMajorMinorBuild(lngPrivateMajor, lngPrivateMinor, lngPrivateBuild)
            End If
        Else
            ' Initialize with Major/Minor/Build/Revision
            intReturnCode = objTargetVersionObject.InitFromMajorMinorBuildRevision(lngPrivateMajor, lngPrivateMinor, lngPrivateBuild, lngPrivateRevision)
        End If

        Clone = intReturnCode
    End Function

    Public Function CompareTo(ByVal objOtherVersionObject)
        ' Compares this version object to the version object supplied as an argument.
        ' Returns 1 if this version object is subsequent/later than the version supplied in
        '   the argument. Also returns 1 if the object supplied as an argument was null/
        '   nothing, or if the object supplied as an argument was not a valid version object
        ' Returns 0 if this version object is equal to the version supplied in the argument
        ' Returns -1 if this version object is before/earlier than the version supplied in the
        '   argument.
        Dim intResult
        Dim lngComparedMajor
        Dim lngComparedMinor
        Dim lngComparedBuild
        Dim lngComparedRevision

        Err.Clear

        intResult = 0
        If TestObjectForData(objOtherVersionObject) = False Then
            intResult = 1
        Else
            On Error Resume Next
            lngComparedMajor = CLng(objOtherVersionObject.Major)
            If Err Then
                Err.Clear
                On Error Goto 0
                intResult = 1
            Else
                lngComparedMinor = CLng(objOtherVersionObject.Minor)
                If Err Then
                    Err.Clear
                    On Error Goto 0
                    intResult = 1
                Else
                    lngComparedBuild = CLng(objOtherVersionObject.Build)
                    If Err Then
                        Err.Clear
                        On Error Goto 0
                        intResult = 1
                    Else
                        lngComparedRevision = CLng(objOtherVersionObject.Revision)
                        If Err Then
                            Err.Clear
                            On Error Goto 0
                            intResult = 1
                        Else
                            On Error Goto 0
                        End If
                    End If
                End If
            End If
        End If

        If intResult = 0 Then
            If lngPrivateMajor <> lngComparedMajor Then
                If lngPrivateMajor < lngComparedMajor Then
                    intResult = -1
                Else
                    intResult = 1
                End If
            ElseIf lngPrivateMinor <> lngComparedMinor Then
                If lngPrivateMinor < lngComparedMinor Then
                    intResult = -1
                Else
                    intResult = 1
                End If
            ElseIf lngPrivateBuild <> lngComparedBuild Then
                If lngPrivateBuild < lngComparedBuild Then
                    intResult = -1
                Else
                    intResult = 1
                End If
            ElseIf lngPrivateRevision <> lngComparedRevision Then
                If lngPrivateRevision < lngComparedRevision Then
                    intResult = -1
                Else
                    intResult = 1
                End If
            End If
        End If
        CompareTo = intResult
    End Function

    Public Function CompareToString(ByVal strOtherVersion)
        ' Compares this version object to the string representation of a version number
        ' supplied as an argument.
        ' Returns 1 if this version object is subsequent/later than the version supplied in
        '   the argument. Also returns 1 if the string supplied as an argument was null/
        '   nothing/empty string, or if the string supplied as an argument was not a valid
        '   version object
        ' Returns 0 if this version object is equal to the version supplied in the argument
        ' Returns -1 if this version object is before/earlier than the version supplied in the
        '   argument.
        Dim objOtherVersionObject
        Dim intReturnCode
        Dim intResult
        Dim lngComparedMajor
        Dim lngComparedMinor
        Dim lngComparedBuild
        Dim lngComparedRevision

        Err.Clear

        intResult = 0
        If TestObjectForData(strOtherVersion) = False Then
            intResult = 1
        Else
            Set objOtherVersionObject = New Version
            intReturnCode = objOtherVersionObject.InitFromString(strOtherVersion)
            If intReturnCode <> 0 Then
                intResult = 1
            End If
        End If

        If intResult = 0 Then
            If TestObjectForData(objOtherVersionObject) = False Then
                intResult = 1
            Else
                On Error Resume Next
                lngComparedMajor = CLng(objOtherVersionObject.Major)
                If Err Then
                    Err.Clear
                    On Error Goto 0
                    intResult = 1
                Else
                    lngComparedMinor = CLng(objOtherVersionObject.Minor)
                    If Err Then
                        Err.Clear
                        On Error Goto 0
                        intResult = 1
                    Else
                        lngComparedBuild = CLng(objOtherVersionObject.Build)
                        If Err Then
                            Err.Clear
                            On Error Goto 0
                            intResult = 1
                        Else
                            lngComparedRevision = CLng(objOtherVersionObject.Revision)
                            If Err Then
                                Err.Clear
                                On Error Goto 0
                                intResult = 1
                            Else
                                On Error Goto 0
                            End If
                        End If
                    End If
                End If
            End If
        End If

        If intResult = 0 Then
            If lngPrivateMajor <> lngComparedMajor Then
                If lngPrivateMajor < lngComparedMajor Then
                    intResult = -1
                Else
                    intResult = 1
                End If
            ElseIf lngPrivateMinor <> lngComparedMinor Then
                If lngPrivateMinor < lngComparedMinor Then
                    intResult = -1
                Else
                    intResult = 1
                End If
            ElseIf lngPrivateBuild <> lngComparedBuild Then
                If lngPrivateBuild < lngComparedBuild Then
                    intResult = -1
                Else
                    intResult = 1
                End If
            ElseIf lngPrivateRevision <> lngComparedRevision Then
                If lngPrivateRevision < lngComparedRevision Then
                    intResult = -1
                Else
                    intResult = 1
                End If
            End If
        End If
        CompareToString = intResult
    End Function

    Public Function Equals(ByVal objOtherVersionObject)
        ' Compares the current version to the version object supplied as an argument.
        ' Returns True if the two versions are equal
        ' Returns False otherwise. Also returns False if the object supplied as an argument is
        '   not a valid Version object
        Dim boolResult
        Dim lngComparedMajor
        Dim lngComparedMinor
        Dim lngComparedBuild
        Dim lngComparedRevision

        Err.Clear

        boolResult = True
        If TestObjectForData(objOtherVersionObject) = False Then
            boolResult = False
        Else
            On Error Resume Next
            lngComparedMajor = CLng(objOtherVersionObject.Major)
            If Err Then
                Err.Clear
                On Error Goto 0
                boolResult = False
            Else
                lngComparedMinor = CLng(objOtherVersionObject.Minor)
                If Err Then
                    Err.Clear
                    On Error Goto 0
                    boolResult = False
                Else
                    lngComparedBuild = CLng(objOtherVersionObject.Build)
                    If Err Then
                        Err.Clear
                        On Error Goto 0
                        boolResult = False
                    Else
                        lngComparedRevision = CLng(objOtherVersionObject.Revision)
                        If Err Then
                            Err.Clear
                            On Error Goto 0
                            boolResult = False
                        Else
                            On Error Goto 0
                        End If
                    End If
                End If
            End If
        End If

        If boolResult = True Then
            If lngPrivateMajor <> lngComparedMajor Then
                boolResult = False
            ElseIf lngPrivateMinor <> lngComparedMinor Then
                boolResult = False
            ElseIf lngPrivateBuild <> lngComparedBuild Then
                boolResult = False
            ElseIf lngPrivateRevision <> lngComparedRevision Then
                boolResult = False
            End If
        End If
        Equals = boolResult
    End Function

    Public Function GreaterThan(ByVal objOtherVersionObject)
        ' Compares the current version to the version object supplied as an argument.
        ' Returns True if this version object is greater than the version supplied as an
        '   argument
        ' Returns False otherwise. Also returns False if the object supplied as an argument is
        '   not a valid Version object
        Dim boolResult
        Dim lngComparedMajor
        Dim lngComparedMinor
        Dim lngComparedBuild
        Dim lngComparedRevision

        Err.Clear

        boolResult = True
        If TestObjectForData(objOtherVersionObject) = False Then
            boolResult = False
        Else
            On Error Resume Next
            lngComparedMajor = CLng(objOtherVersionObject.Major)
            If Err Then
                Err.Clear
                On Error Goto 0
                boolResult = False
            Else
                lngComparedMinor = CLng(objOtherVersionObject.Minor)
                If Err Then
                    Err.Clear
                    On Error Goto 0
                    boolResult = False
                Else
                    lngComparedBuild = CLng(objOtherVersionObject.Build)
                    If Err Then
                        Err.Clear
                        On Error Goto 0
                        boolResult = False
                    Else
                        lngComparedRevision = CLng(objOtherVersionObject.Revision)
                        If Err Then
                            Err.Clear
                            On Error Goto 0
                            boolResult = False
                        Else
                            On Error Goto 0
                        End If
                    End If
                End If
            End If
        End If

        If boolResult = True Then
            If lngPrivateMajor < lngComparedMajor Then
                boolResult = False
            ElseIf lngPrivateMajor = lngComparedMajor Then
                If lngPrivateMinor < lngComparedMinor Then
                    boolResult = False
                ElseIf lngPrivateMinor = lngComparedMinor Then
                    If lngPrivateBuild < lngComparedBuild Then
                        boolResult = False
                    ElseIf lngPrivateBuild = lngComparedBuild Then
                        If lngPrivateRevision <= lngComparedRevision Then
                            boolResult = False
                        End If
                    End If
                End If
            End If
        End If
        GreaterThan = boolResult
    End Function

    Public Function GreaterThanOrEqual(ByVal objOtherVersionObject)
        ' Compares the current version to the version object supplied as an argument.
        ' Returns True if this version object is greater than or equal to the version supplied
        '   as an argument
        ' Returns False otherwise. Also returns False if the object supplied as an argument is
        '   not a valid Version object
        Dim boolResult
        Dim lngComparedMajor
        Dim lngComparedMinor
        Dim lngComparedBuild
        Dim lngComparedRevision

        Err.Clear

        boolResult = True
        If TestObjectForData(objOtherVersionObject) = False Then
            boolResult = False
        Else
            On Error Resume Next
            lngComparedMajor = CLng(objOtherVersionObject.Major)
            If Err Then
                Err.Clear
                On Error Goto 0
                boolResult = False
            Else
                lngComparedMinor = CLng(objOtherVersionObject.Minor)
                If Err Then
                    Err.Clear
                    On Error Goto 0
                    boolResult = False
                Else
                    lngComparedBuild = CLng(objOtherVersionObject.Build)
                    If Err Then
                        Err.Clear
                        On Error Goto 0
                        boolResult = False
                    Else
                        lngComparedRevision = CLng(objOtherVersionObject.Revision)
                        If Err Then
                            Err.Clear
                            On Error Goto 0
                            boolResult = False
                        Else
                            On Error Goto 0
                        End If
                    End If
                End If
            End If
        End If

        If boolResult = True Then
            If lngPrivateMajor < lngComparedMajor Then
                boolResult = False
            ElseIf lngPrivateMinor < lngComparedMinor Then
                boolResult = False
            ElseIf lngPrivateBuild < lngComparedBuild Then
                boolResult = False
            ElseIf lngPrivateRevision < lngComparedRevision Then
                boolResult = False
            End If
        End If
        GreaterThanOrEqual = boolResult
    End Function

    Public Function InitFromMajorMinor(ByVal lngMajor, ByVal lngMinor)
        ' Initalizes a Version object from a pair of long integers supplied by two arguments.
        ' The first argument is the major version number
        ' The second argument is the minor version number
        ' For example: major.minor
        ' This method returns 0 if successful; non-zero otherwise.
        Dim intFunctionReturn

        intFunctionReturn = InitFromMajorMinorBuildRevision(lngMajor, lngMinor, 0, 0)

        If intFunctionReturn = 0 Then
            lngPrivateBuild = CLng(-1)
            lngPrivateRevision = CLng(-1)
        End If

        InitFromMajorMinor = intFunctionReturn
    End Function

    Public Function InitFromMajorMinorBuild(ByVal lngMajor, ByVal lngMinor, ByVal lngBuild)
        ' Initalizes a Version object from three long integers supplied by three arguments.
        ' The first argument is the major version number
        ' The second argument is the minor version number
        ' The third argument is the build number
        ' For example: major.minor.build
        ' This method returns 0 if successful; non-zero otherwise.
        Dim intFunctionReturn

        intFunctionReturn = InitFromMajorMinorBuildRevision(lngMajor, lngMinor, lngBuild, 0)

        If intFunctionReturn = 0 Then
            lngPrivateRevision = CLng(-1)
        End If

        InitFromMajorMinorBuild = intFunctionReturn
    End Function

    Public Function InitFromMajorMinorBuildRevision(ByVal lngMajor, ByVal lngMinor, ByVal lngBuild, ByVal lngRevision)
        ' Initalizes a Version object from four long integers supplied by four arguments.
        ' The first argument is the major version number
        ' The second argument is the minor version number
        ' The third argument is the build number
        ' The fourth argument is the revision number
        ' For example: major.minor.build.revision
        ' This method returns 0 if successful; non-zero otherwise.
        Dim intFunctionReturn
        Dim lngTempMajor
        Dim lngTempMinor
        Dim lngTempBuild
        Dim lngTempRevision

        Err.Clear

        intFunctionReturn = 0

        If TestObjectForData(lngMajor) = False Then
            ' Blank sections of the version number are allowed here
            lngTempMajor = CLng(0)
        Else
            On Error Resume Next
            lngTempMajor = CLng(lngMajor)
            If Err Then
                Err.Clear
                On Error Goto 0
                ' The "major" portion of the version number was not a valid long integer
                intFunctionReturn = -1
            Else
                On Error Goto 0
                If lngTempMajor < CLng(0) Then
                    ' Cannot have negative version numbers
                    intFunctionReturn = -2
                Else
                    lngTempMinor = CLng(0)
                    lngTempBuild = CLng(0)
                    lngTempRevision = CLng(0)
                End If
            End If
        End If

        If intFunctionReturn = 0 Then
            ' No error occurred
            If TestObjectForData(lngMinor) = False Then
                ' Blank sections of the version number are allowed here
                ' Already set; nothing more to do
            Else
                On Error Resume Next
                lngTempMinor = CLng(lngMinor)
                If Err Then
                    Err.Clear
                    On Error Goto 0
                    ' The "minor" portion of the version number was not a valid long integer
                    intFunctionReturn = -3
                Else
                    On Error Goto 0
                    If lngTempMinor < CLng(0) Then
                        ' Cannot have negative version numbers
                        intFunctionReturn = -4
                    End If
                End If
            End If
        End If

        If intFunctionReturn = 0 Then
            ' No error occurred
            If TestObjectForData(lngBuild) = False Then
                ' Blank sections of the version number are allowed here
                ' Already set; nothing more to do
            Else
                On Error Resume Next
                lngTempBuild = CLng(lngBuild)
                If Err Then
                    Err.Clear
                    On Error Goto 0
                    ' The "build" portion of the version number was not a valid long integer
                    intFunctionReturn = -5
                Else
                    On Error Goto 0
                    If lngTempBuild < CLng(0) Then
                        ' Cannot have negative version numbers
                        intFunctionReturn = -6
                    End If
                End If
            End If
        End If

        If intFunctionReturn = 0 Then
            ' No error occurred
            If TestObjectForData(lngRevision) = False Then
                ' Blank sections of the version number are allowed here
                ' Already set; nothing more to do
            Else
                On Error Resume Next
                lngTempRevision = CLng(lngRevision)
                If Err Then
                    Err.Clear
                    On Error Goto 0
                    ' The "Revision" portion of the version number was not a valid long integer
                    intFunctionReturn = -7
                Else
                    On Error Goto 0
                    If lngTempRevision < CLng(0) Then
                        ' Cannot have negative version numbers
                        intFunctionReturn = -8
                    End If
                End If
            End If
        End If

        If intFunctionReturn = 0 Then
            ' No error occurred
            lngPrivateMajor = lngTempMajor
            lngPrivateMinor = lngTempMinor
            lngPrivateBuild = lngTempBuild
            lngPrivateRevision = lngTempRevision
        End If

        InitFromMajorMinorBuildRevision = intFunctionReturn
    End Function

    Public Default Function InitFromString(ByVal strVersion)
        ' Initalizes a Version object from a version-formatted string supplied as an argument.
        ' Valid strings look like the following:
        ' major.minor
        ' major.minor.build
        ' or
        ' major.minor.build.revision
        ' Each part of the version string must be in decimal and convertable to a long integer
        ' This method returns 0 if successful; non-zero otherwise.
        Dim intFunctionReturn
        Dim arrVersion
        Dim intCountOfVersionSections
        Dim boolVersionSectionCountTest
        Dim lngTempMajor
        Dim lngTempMinor
        Dim lngTempBuild
        Dim lngTempRevision

        Err.Clear

        intFunctionReturn = 0

        If TestObjectForData(strVersion) = False Then
            ' No data was passed to function
            intFunctionReturn = -1
        Else
            On Error Resume Next
            arrVersion = Split(strVersion, ".")
            If Err Then
                Err.Clear
                On Error Goto 0
                ' Object passed to function was not a string, or an error occurred splitting
                ' the string
                intFunctionReturn = -2
            Else
                intCountOfVersionSections = UBound(arrVersion)
                If Err Then
                    Err.Clear
                    On Error Goto 0
                    ' Something went wrong reading the upper boundary of the array resulting
                    ' from the Split() function
                    intFunctionReturn = -3
                Else
                    boolVersionSectionCountTest = (intCountOfVersionSections > 3) Or (intCountOfVersionSections < 1)
                    If Err Then
                        Err.Clear
                        On Error Goto 0
                        ' Something went wrong comparing the upper boundary to an interger
                        intFunctionReturn = -4
                    Else
                        On Error Goto 0
                    End If
                End If
            End If
        End If

        If intFunctionReturn = 0 Then
            ' No error occurred
            If boolVersionSectionCountTest = True Then
                ' Less than two parts of the version string were passed (e.g., "1")
                ' or
                ' More than four parts of the version string were passed (e.g., "1.2.3.4.5")
                ' Neither is allowed here, nor the System.Version .NET analog
                intFunctionReturn = -5
            Else
                ' String appears valid so far and has 2-4 parts, e.g.:
                ' 1.2
                ' 1.2.3
                ' 1.2.3.4
                If TestObjectForData(arrVersion(0)) = False Then
                    ' Blank sections of the version number are not allowed during conversion
                    ' from string
                    intFunctionReturn = -6
                Else
                    On Error Resume Next
                    lngTempMajor = CLng(arrVersion(0))
                    If Err Then
                        Err.Clear
                        On Error Goto 0
                        ' The "major" portion of the version number was not a valid long
                        ' integer
                        intFunctionReturn = -7
                    Else
                        On Error Goto 0
                        If lngTempMajor < CLng(0) Then
                            ' Cannot have negative version numbers
                            intFunctionReturn = -8
                        Else
                            lngTempMinor = CLng(0)
                            lngTempBuild = CLng(0)
                            lngTempRevision = CLng(0)
                        End If
                    End If
                End If
            End If
        End If

        If intFunctionReturn = 0 Then
            ' No error occurred
            If TestObjectForData(arrVersion(1)) = False Then
                ' Blank sections of the version number are not allowed during conversion
                ' from string
                intFunctionReturn = -9
            Else
                On Error Resume Next
                lngTempMinor = CLng(arrVersion(1))
                If Err Then
                    Err.Clear
                    On Error Goto 0
                    ' The "minor" portion of the version number was not a valid long integer
                    intFunctionReturn = -10
                Else
                    On Error Goto 0
                    If lngTempMinor < CLng(0) Then
                        ' Cannot have negative version numbers
                        intFunctionReturn = -11
                    End If
                End If
            End If
        End If

        If intFunctionReturn = 0 Then
            ' No error occurred
            If intCountOfVersionSections >= 2 Then
                ' Build portion of version should be present
                If TestObjectForData(arrVersion(2)) = False Then
                    ' Blank sections of the version number are not allowed during conversion
                    ' from string
                    intFunctionReturn = -12
                Else
                    On Error Resume Next
                    lngTempBuild = CLng(arrVersion(2))
                    If Err Then
                        Err.Clear
                        On Error Goto 0
                        ' The "build" portion of the version number was not a valid long integer
                        intFunctionReturn = -13
                    Else
                        On Error Goto 0
                        If lngTempBuild < CLng(0) Then
                            ' Cannot have negative version numbers
                            intFunctionReturn = -14
                        End If
                    End If
                End If
            Else
                lngTempBuild = CLng(-1)
            End If
        End If

        If intFunctionReturn = 0 Then
            ' No error occurred
            If intCountOfVersionSections = 3 Then
                ' Revision portion of version should be present
                If TestObjectForData(arrVersion(3)) = False Then
                    ' Blank sections of the version number are not allowed during conversion
                    ' from string
                    intFunctionReturn = -15
                Else
                    On Error Resume Next
                    lngTempRevision = CLng(arrVersion(3))
                    If Err Then
                        Err.Clear
                        On Error Goto 0
                        ' The "revision" portion of the version number was not a valid long integer
                        intFunctionReturn = -16
                    Else
                        On Error Goto 0
                        If lngTempRevision < CLng(0) Then
                            ' Cannot have negative version numbers
                            intFunctionReturn = -17
                        End If
                    End If
                End If
            Else
                lngTempRevision = CLng(-1)
            End If
        End If

        If intFunctionReturn = 0 Then
            ' No error occurred
            lngPrivateMajor = lngTempMajor
            lngPrivateMinor = lngTempMinor
            lngPrivateBuild = lngTempBuild
            lngPrivateRevision = lngTempRevision
        End If

        InitFromString = intFunctionReturn
    End Function

    Public Function LessThan(ByVal objOtherVersionObject)
        ' Compares the current version to the version object supplied as an argument.
        ' Returns True if this version object is less than the version supplied as an argument
        ' Returns False otherwise. Also returns False if the object supplied as an argument is
        '   not a valid Version object
        Dim boolResult
        Dim lngComparedMajor
        Dim lngComparedMinor
        Dim lngComparedBuild
        Dim lngComparedRevision

        Err.Clear

        boolResult = True
        If TestObjectForData(objOtherVersionObject) = False Then
            boolResult = False
        Else
            On Error Resume Next
            lngComparedMajor = CLng(objOtherVersionObject.Major)
            If Err Then
                Err.Clear
                On Error Goto 0
                boolResult = False
            Else
                lngComparedMinor = CLng(objOtherVersionObject.Minor)
                If Err Then
                    Err.Clear
                    On Error Goto 0
                    boolResult = False
                Else
                    lngComparedBuild = CLng(objOtherVersionObject.Build)
                    If Err Then
                        Err.Clear
                        On Error Goto 0
                        boolResult = False
                    Else
                        lngComparedRevision = CLng(objOtherVersionObject.Revision)
                        If Err Then
                            Err.Clear
                            On Error Goto 0
                            boolResult = False
                        Else
                            On Error Goto 0
                        End If
                    End If
                End If
            End If
        End If

        If boolResult = True Then
            If lngPrivateMajor > lngComparedMajor Then
                boolResult = False
            ElseIf lngPrivateMajor = lngComparedMajor Then
                If lngPrivateMinor > lngComparedMinor Then
                    boolResult = False
                ElseIf lngPrivateMinor = lngComparedMinor Then
                    If lngPrivateBuild > lngComparedBuild Then
                        boolResult = False
                    ElseIf lngPrivateBuild = lngComparedBuild Then
                        If lngPrivateRevision >= lngComparedRevision Then
                            boolResult = False
                        End If
                    End If
                End If
            End If
        End If
        LessThan = boolResult
    End Function

    Public Function LessThanOrEqual(ByVal objOtherVersionObject)
        ' Compares the current version to the version object supplied as an argument.
        ' Returns True if this version object is less than or equal to the version supplied as
        '   an argument
        ' Returns False otherwise. Also returns False if the object supplied as an argument is
        '   not a valid Version object
        Dim boolResult
        Dim lngComparedMajor
        Dim lngComparedMinor
        Dim lngComparedBuild
        Dim lngComparedRevision

        Err.Clear

        boolResult = True
        If TestObjectForData(objOtherVersionObject) = False Then
            boolResult = False
        Else
            On Error Resume Next
            lngComparedMajor = CLng(objOtherVersionObject.Major)
            If Err Then
                Err.Clear
                On Error Goto 0
                boolResult = False
            Else
                lngComparedMinor = CLng(objOtherVersionObject.Minor)
                If Err Then
                    Err.Clear
                    On Error Goto 0
                    boolResult = False
                Else
                    lngComparedBuild = CLng(objOtherVersionObject.Build)
                    If Err Then
                        Err.Clear
                        On Error Goto 0
                        boolResult = False
                    Else
                        lngComparedRevision = CLng(objOtherVersionObject.Revision)
                        If Err Then
                            Err.Clear
                            On Error Goto 0
                            boolResult = False
                        Else
                            On Error Goto 0
                        End If
                    End If
                End If
            End If
        End If

        If boolResult = True Then
            If lngPrivateMajor > lngComparedMajor Then
                boolResult = False
            ElseIf lngPrivateMinor > lngComparedMinor Then
                boolResult = False
            ElseIf lngPrivateBuild > lngComparedBuild Then
                boolResult = False
            ElseIf lngPrivateRevision > lngComparedRevision Then
                boolResult = False
            End If
        End If
        LessThanOrEqual = boolResult
    End Function

    Public Function NotEquals(ByVal objOtherVersionObject)
        ' Compares the current version to the version object supplied as an argument.
        ' Returns True if the two versions are not equal. Also returns True if the object
        '   supplied as an argument is not a valid Version object
        ' Returns False otherwise. 
        Dim boolResult
        Dim lngComparedMajor
        Dim lngComparedMinor
        Dim lngComparedBuild
        Dim lngComparedRevision

        Err.Clear

        boolResult = False
        If TestObjectForData(objOtherVersionObject) = False Then
            boolResult = True
        Else
            On Error Resume Next
            lngComparedMajor = CLng(objOtherVersionObject.Major)
            If Err Then
                Err.Clear
                On Error Goto 0
                boolResult = True
            Else
                lngComparedMinor = CLng(objOtherVersionObject.Minor)
                If Err Then
                    Err.Clear
                    On Error Goto 0
                    boolResult = True
                Else
                    lngComparedBuild = CLng(objOtherVersionObject.Build)
                    If Err Then
                        Err.Clear
                        On Error Goto 0
                        boolResult = True
                    Else
                        lngComparedRevision = CLng(objOtherVersionObject.Revision)
                        If Err Then
                            Err.Clear
                            On Error Goto 0
                            boolResult = True
                        Else
                            On Error Goto 0
                        End If
                    End If
                End If
            End If
        End If

        If boolResult = False Then
            If lngPrivateMajor <> lngComparedMajor Then
                boolResult = True
            ElseIf lngPrivateMinor <> lngComparedMinor Then
                boolResult = True
            ElseIf lngPrivateBuild <> lngComparedBuild Then
                boolResult = True
            ElseIf lngPrivateRevision <> lngComparedRevision Then
                boolResult = True
            End If
        End If
        NotEquals = boolResult
    End Function

    Public Function ToString()
        ' Returns a dot-separated representation of the version number as a string.
        ' Valid strings look like the following:
        ' major.minor
        ' major.minor.build
        ' or
        ' major.minor.build.revision
        ' (where each part of the string is a long integer converted to string format)
        Dim strToReturn
        If lngPrivateRevision = CLng(-1) Then
            If lngPrivateBuild = CLng(-1) Then
                ' Output Major/Minor only
                strToReturn = CStr(lngPrivateMajor) & "." & CStr(lngPrivateMinor)
            Else
                ' Output Major/Minor/Build only
                strToReturn = CStr(lngPrivateMajor) & "." & CStr(lngPrivateMinor) & "." & CStr(lngPrivateBuild)
            End If
        Else
            ' Output Major/Minor/Build/Revision
            strToReturn = CStr(lngPrivateMajor) & "." & CStr(lngPrivateMinor) & "." & CStr(lngPrivateBuild) & "." & CStr(lngPrivateRevision)
        End If
        ToString = strToReturn
    End Function

    Public Property Get Major()
        Major = lngPrivateMajor
    End Property

    Public Property Get Minor()
        Minor = lngPrivateMinor
    End Property

    Public Property Get Build()
        Build = lngPrivateBuild
    End Property

    Public Property Get Revision()
        Revision = lngPrivateRevision
    End Property

    Public Property Get MajorRevision()
        ' Returns the "upper" 16-bits of the revision number. The upper 16-bits are down-
        ' shifted by 16 bits and returned as a 16-bit integer.
        ' If the revision was uninitialized (-1), then -1 is returned.
        Dim lngBitMask
        Dim lngShiftRightDivisor
        If lngPrivateRevision = CLng(-1) Then
            MajorRevision = CInt(-1)
        Else
            lngBitMask = &H7FFF0000
            lngShiftRightDivisor = &H10000
            MajorRevision = CInt((lngPrivateRevision And lngBitMask) / lngShiftRightDivisor)
        End If
    End Property

    Public Property Get MinorRevision()
        ' Returns the "lower" 16-bits of the revision number, returned as a 16-bit integer.
        ' If the revision was uninitialized (-1), then -1 is returned.
        Dim lngBitMask
        If lngPrivateRevision = CLng(-1) Then
            MinorRevision = CInt(-1)
        Else
            ' 65535 is FFFF in hex; can't use hex because it's interpreted as -1
            lngBitMask = CLng(65535)
            MinorRevision = CInt(lngPrivateRevision And lngBitMask)
        End If
    End Property
End Class

' Restore error handling for VBScript 1.0 - 4.0:
Err.Clear
On Error Goto 0
