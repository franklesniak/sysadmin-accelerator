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
