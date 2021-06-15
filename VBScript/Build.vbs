Option Explicit

Function TestObjectForData(ByVal objToCheck)
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

    'region DependsOn ####################################################
    ' None!
    'endregion DependsOn ####################################################

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

Function TestObjectIsStringContainingData(ByRef objToTest)
    'region FunctionMetadata ####################################################
    ' Safely determines if the specified object is a string or not
    '
    ' Function takes one positional argument (objToTest), which is the object to be tested to
    '   determine if it is a string.
    '
    ' The function returns boolean True if the specified object is a string that contains data,
    ' boolean False otherwise
    '
    ' Example 1:
    '   objToTest = "12345"
    '   boolResult = TestObjectIsStringContainingData(objToTest)
    '   ' boolResult is equal to True
    '
    ' Example 2:
    '   objToTest = ""
    '   boolResult = TestObjectIsStringContainingData(objToTest)
    '   ' boolResult is equal to False
    '
    ' Example 3:
    '   objToTest = 12345
    '   boolResult = TestObjectIsStringContainingData(objToTest)
    '   ' boolResult is equal to False
    '
    ' Version: 1.0.20210613.0
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
    'endregion DependsOn ####################################################

    Dim boolFunctionReturn
    Dim boolTest
    Dim intVarType

    If TestObjectForData(objToTest) = False Then
        boolFunctionReturn = False
    Else
        ' objToTest has data
        On Error Resume Next
        intVarType = VarType(objToTest)
        If Err Then
            On Error Goto 0
            Err.Clear
            boolFunctionReturn = False
        Else
            boolTest = (intVarType <> 8)
            If Err Then
                On Error Goto 0
                Err.Clear
                boolFunctionReturn = False
            Else
                On Error Goto 0
                If boolTest = True Then
                    ' VarType(objToTest) <> 8
                    boolFunctionReturn = False
                Else
                    ' VarType(objToTest) = 8
                    boolFunctionReturn = True
                End If
            End If
        End If
    End If

    TestObjectIsStringContainingData = boolFunctionReturn
End Function

' Thanks to Fionnuala who reminded me that ADODB.RecordSets can sort in-memory in VBScript:
' https://stackoverflow.com/a/308735/2134110

Dim arrSubfolderNames
Dim strSubfolderName
Dim objFileSystemObject
Dim strScriptFullName
Dim strScriptDir
Dim objFolder
Dim arrFiles
Dim objFile
Dim strOutput
Dim objTextStreamFile
Dim strTextFile
Dim strOutputFileName
Dim objADODBRecordSet

Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
Const Unicode = -1
Const ASCII = 0
Const UseSystemDefault = -2
Const adVarChar = 200
Const adOpenStatic = 3

arrSubfolderNames = Array("01_Overall_Script_Header", "02_Upfront_Encapsualed_Code_With_No_Dependencies", "03_Main_Section_Code_Executed_Every_Time", "04_Later_Encapsulated_Code_With_Dependencies_on_Main", "05_Script_Footer")
strOutputFileName = "Accelerator.vbs"

strOutput = ""

On Error Resume Next
Set objFileSystemObject = CreateObject("Scripting.FileSystemObject")
If Err Then
    On Error Goto 0
    Err.Clear
Else
    strScriptFullName = WScript.ScriptFullName
    If Err Then
        On Error Goto 0
        Err.Clear
    Else
        strScriptDir = objFileSystemObject.GetParentFolderName(strScriptFullName)
        If Err Then
            On Error Goto 0
            Err.Clear
        Else
            On Error Goto 0
            If TestObjectIsStringContainingData(strScriptDir) = True Then
                If Right(strScriptDir,1) <> "\" Then strScriptDir = strScriptDir & "\"
                For Each strSubfolderName in arrSubfolderNames
                    If objFileSystemObject.FolderExists(strScriptDir & strSubfolderName & "\") Then
                        Set objFolder = objFileSystemObject.GetFolder(strScriptDir & strSubfolderName & "\")
                        Set objADODBRecordSet = CreateObject("ADODB.RecordSet")
                        objADODBRecordSet.Fields.Append "FilePath", adVarChar, 255
                        objADODBRecordSet.CursorType = adOpenStatic
                        objADODBRecordSet.Open
                        Set arrFiles = objFolder.Files
                        For Each objFile in arrFiles
                            objADODBRecordSet.AddNew "FilePath", objFile.Path
                            'objADODBRecordSet.Fields(0) = objFile.Path
                            objADODBRecordSet.Update
                        Next
                        objADODBRecordSet.Sort = "FilePath"
                        If objADODBRecordSet.BOF = False Then objADODBRecordSet.MoveFirst()
                        Do Until objADODBRecordSet.EOF
                            strTextFile = ""
                            ' WScript.Echo objADODBRecordSet.Fields("FilePath")
                            Set objTextStreamFile = objFileSystemObject.OpenTextFile(objADODBRecordSet.Fields("FilePath"), ForReading, False, ASCII)
                            strTextFile = objTextStreamFile.ReadAll
                            objTextStreamFile.Close
                            If strOutput = "" Then
                                strOutput = strTextFile
                            Else
                                strOutput = strOutput & vbCrLf & strTextFile
                            End If
                            objADODBRecordSet.MoveNext
                        Loop
                        Set objADODBRecordSet = Nothing
                    End If
                Next
                Set objTextStreamFile =  objFileSystemObject.OpenTextFile(strScriptDir & strOutputFileName, ForWriting, True, ASCII)
                objTextStreamFile.Write strOutput
                objTextStreamFile.Close
            End If
        End If
    End If
End If
