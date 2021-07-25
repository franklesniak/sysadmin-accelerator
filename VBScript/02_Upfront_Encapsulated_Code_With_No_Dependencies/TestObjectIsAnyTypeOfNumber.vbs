Function TestObjectIsAnyTypeOfNumber(ByRef objToTest)
    'region FunctionMetadata ####################################################
    ' Safely determines if the specified object is a number (of any kind). In other words, this
    ' function checks to see if the specified object is an integer or floating point number
    '
    ' Function takes one positional argument (objToTest), which is the object to be tested to
    ' determine if it is an integer or floating point number
    '
    ' The function returns boolean True if the specified object is an integer or floating point
    ' number, boolean False otherwise
    '
    ' Example 1:
    '   objToTest = "12345"
    '   boolResult = TestObjectIsAnyTypeOfNumber(objToTest)
    '   ' boolResult is equal to False
    '
    ' Example 2:
    '   objToTest = 0
    '   boolResult = TestObjectIsAnyTypeOfNumber(objToTest)
    '   ' boolResult is equal to True
    '
    ' Example 3:
    '   objToTest = 12345
    '   boolResult = TestObjectIsAnyTypeOfNumber(objToTest)
    '   ' boolResult is equal to True
    '
    ' Example 4:
    '   objToTest = 12345.678
    '   boolResult = TestObjectIsAnyTypeOfNumber(objToTest)
    '   ' boolResult is equal to True
    '
    ' Example 5:
    '   objToTest = True
    '   boolResult = TestObjectIsAnyTypeOfNumber(objToTest)
    '   ' boolResult is equal to False
    '
    ' Version: 1.0.20210724.0
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
            boolTest = (intVarType <> 2 And intVarType <> 3 And intVarType <> 4 And intVarType <> 5 And intVarType <> 17 And intVarType <> 20)
            If Err Then
                On Error Goto 0
                Err.Clear
                boolFunctionReturn = False
            Else
                On Error Goto 0
                If boolTest = True Then
                    ' VarType(objToTest) <> 2 And VarType(objToTest) <> 3 And VarType(objToTest) <> 4 And VarType(objToTest) <> 5 And VarType(objToTest) <> 17 And VarType(objToTest) <> 20
                    boolFunctionReturn = False
                Else
                    ' VarType(objToTest) = 2 Or VarType(objToTest) = 3 Or VarType(objToTest) = 4 Or VarType(objToTest) = 5 Or VarType(objToTest) = 17 Or VarType(objToTest) = 20
                    boolFunctionReturn = True
                End If
            End If
        End If
    End If

    TestObjectIsAnyTypeOfNumber = boolFunctionReturn
End Function
