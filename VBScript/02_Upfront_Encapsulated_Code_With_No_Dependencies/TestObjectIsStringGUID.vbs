Function TestObjectIsStringGUID(ByRef boolResult, ByVal objToTest)
    'region FunctionMetadata #######################################################
    ' Safely determines if the specified object is a string representation of a
    ' globally-unique identifier (GUID)
    '
    ' This function takes two positional arguments:
    '   If the test was successful, the first argument (boolResult) contains a boolean
    '   value: True when the specified object is a string representation of a GUID,
    '   False otherwise
    '
    '   The second argument (objToTest) is the object to be tested to determine if it is
    '   a string representation of a GUID
    '
    ' The function returns 0 if the test was successful, a negative integer otherwise.
    '
    ' Note: this function requires VBScript 5.0 or later, which is included with
    ' Internet Explorer 5.0 or later, Windows Scripting Host 2.0 or later, and Windows
    ' 2000 or later.
    '
    ' Example 1:
    '   objToTest = "8589D070-619D-41FB-9D93-D6F50F6AD99D"
    '   intReturnCode = TestObjectIsStringGUID(boolResult, objToTest)
    '   If intReturnCode = 0 Then
    '       ' boolResult is equal to True
    '   End If
    '
    ' Example 2:
    '   objToTest = "{8589D070-619D-41FB-9D93-D6F50F6AD99D}"
    '   boolResult = TestObjectIsStringGUID(objToTest)
    '   intReturnCode = TestObjectIsStringGUID(boolResult, objToTest)
    '   If intReturnCode = 0 Then
    '       ' boolResult is equal to True
    '   End If
    '
    ' Example 3:
    '   objToTest = "(8589D070-619D-41FB-9D93-D6F50F6AD99D)"
    '   boolResult = TestObjectIsStringGUID(objToTest)
    '   intReturnCode = TestObjectIsStringGUID(boolResult, objToTest)
    '   If intReturnCode = 0 Then
    '       ' boolResult is equal to True
    '   End If
    '
    ' Example 4:
    '   objToTest = "{8589D070-619D-41FB-9D93-D6F50F6AD99D)"
    '   boolResult = TestObjectIsStringGUID(objToTest)
    '   intReturnCode = TestObjectIsStringGUID(boolResult, objToTest)
    '   If intReturnCode = 0 Then
    '       ' boolResult is equal to False (mismatched brackets around GUID)
    '   End If
    '
    ' Example 5:
    '   objToTest = "8589D070619D41FB9D93D6F50F6AD99D"
    '   intReturnCode = TestObjectIsStringGUID(boolResult, objToTest)
    '   If intReturnCode = 0 Then
    '       ' boolResult is equal to True
    '   End If
    '
    ' Example 6:
    '   objToTest = "8589D070619D41FB9D93D6F50F6AD99DE"
    '   intReturnCode = TestObjectIsStringGUID(boolResult, objToTest)
    '   If intReturnCode = 0 Then
    '       ' boolResult is equal to False (too many characters)
    '   End If
    '
    ' Example 7:
    '   objToTest = "8589D070619D41FB9D93G6F50F6AD99D"
    '   intReturnCode = TestObjectIsStringGUID(boolResult, objToTest)
    '   If intReturnCode = 0 Then
    '       ' boolResult is equal to False (invalid character)
    '   End If
    '
    ' Example 8:
    '   objToTest = "a4a43a35-d139-480d-a9fd-6f032257c7aa"
    '   intReturnCode = TestObjectIsStringGUID(boolResult, objToTest)
    '   If intReturnCode = 0 Then
    '       ' boolResult is equal to True (lowercase letters are valid)
    '   End If
    '
    ' Version: 1.0.20230423.0
    'endregion FunctionMetadata #######################################################

    'region License ################################################################
    ' Copyright 2023 Frank Lesniak
    '
    ' Permission is hereby granted, free of charge, to any person obtaining a copy of
    ' this software and associated documentation files (the "Software"), to deal in the
    ' Software without restriction, including without limitation the rights to use,
    ' copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the
    ' Software, and to permit persons to whom the Software is furnished to do so,
    ' subject to the following conditions:
    '
    ' The above copyright notice and this permission notice shall be included in all
    ' copies or substantial portions of the Software.
    '
    ' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    ' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
    ' FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR
    ' COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN
    ' AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
    ' WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
    'endregion License ################################################################

    'region DownloadLocationNotice #################################################
    ' The most up-to-date version of this script can be found on the author's GitHub
    ' repository at https://github.com/franklesniak/sysadmin-accelerator
    'endregion DownloadLocationNotice #################################################

    'region Acknowledgements #######################################################
    ' Thanks to Gunter Born for covering RegExp in "Microsoft Windows Scripting Host
    ' 2.0 Developer's Guide", which indicated that RegExp requries VBScript 5.0 or
    ' newer and saved me some time.
    'endregion Acknowledgements #######################################################

    'region DependsOn ##############################################################
    ' TestObjectIsStringContainingData()
    'endregion DependsOn ##############################################################

    Dim intFunctionReturn
    Dim objRegExp
    Dim boolTestResult
    Dim boolTestResultToReturn

    intFunctionReturn = 0

    If TestObjectIsStringContainingData(objToTest) = False Then
        intFunctionReturn = -1
    Else
        ' objToTest is a string containing data
        On Error Resume Next
        Set objRegExp = CreateObject("VBScript.RegExp")
        If Err Then
            ' RegExp object could not be created
            Err.Clear
            On Error Goto 0
            intFunctionReturn = -2
        Else
            ' RegExp object was created successfully
            ' Check to see if the string is a GUID wrapped in curly braces
            objRegExp.Pattern = "^{[A-Fa-f0-9]{8}[-]?(?:[A-Fa-f0-9]{4}[-]?){3}[A-Fa-f0-9]{12}}$"
            boolTestResult = objRegExp.Test(objToTest)
            If Err Then
                ' RegExp object failed to test objToTest
                Err.Clear
                On Error Goto 0
                intFunctionReturn = -3
            Else
                ' RegExp object was used to test objToTest successfully
                On Error Goto 0
                If boolTestResult = True Then
                    ' objToTest is a string representation of a GUID
                    boolTestResultToReturn = True
                Else
                    ' String was not a GUID wrapped in curly braces
                    ' Check to see if string is a GUID wrapped in parentheses
                    On Error Resume Next
                    objRegExp.Pattern = "^\([A-Fa-f0-9]{8}[-]?(?:[A-Fa-f0-9]{4}[-]?){3}[A-Fa-f0-9]{12}\)$"
                    boolTestResult = objRegExp.Test(objToTest)
                    If Err Then
                        ' RegExp object failed to test objToTest
                        Err.Clear
                        On Error Goto 0
                        intFunctionReturn = -4
                    Else
                        ' RegExp object was used to test objToTest successfully
                        On Error Goto 0
                        If boolTestResult = True Then
                            ' objToTest is a string representation of a GUID
                            boolTestResultToReturn = True
                        Else
                            ' String was not a GUID wrapped in curly braces or
                            ' parentheses
                            ' Check to see if string is a GUID without any
                            ' wrapping
                            On Error Resume Next
                            objRegExp.Pattern = "^[A-Fa-f0-9]{8}[-]?(?:[A-Fa-f0-9]{4}[-]?){3}[A-Fa-f0-9]{12}$"
                            boolTestResult = objRegExp.Test(objToTest)
                            If Err Then
                                ' RegExp object failed to test objToTest
                                Err.Clear
                                On Error Goto 0
                                intFunctionReturn = -5
                            Else
                                ' RegExp object was used to test objToTest successfully
                                On Error Goto 0
                                If boolTestResult = True Then
                                    ' objToTest is a string representation of a GUID
                                    boolTestResultToReturn = True
                                Else
                                    ' String was not a GUID wrapped in curly braces,
                                    ' parentheses, or without any wrapping
                                    boolTestResultToReturn = False
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If

    If intFunctionReturn = 0 Then
        boolResult = boolTestResultToReturn
    End If

    TestObjectIsStringGUID = intFunctionReturn
End Function
