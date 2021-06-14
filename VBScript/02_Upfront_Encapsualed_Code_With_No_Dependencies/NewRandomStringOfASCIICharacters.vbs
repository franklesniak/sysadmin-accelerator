Function NewRandomStringOfASCIICharacters(ByRef strRandom, ByVal intNumberOfCharactersToGenerate, ByVal boolIncludeLowerCase, ByVal boolIncludeUpperCase, ByVal boolIncludeNumbers, ByVal boolIncludePunctuation, ByVal boolIncludeSymbols)
    'region FunctionMetadata ####################################################
    ' Generates a string of the specified number and type of ASCII characters
    '
    ' Function takes seven positional arguments:
    '   The first argument (strRandom) will be populated upon success with the randomly-
    '       generated string of ASCII characters.
    '   The second argument (intNumberOfCharactersToGenerate) is an integer specifying the
    '       number of ASCII characters to generate
    '   The third argument (boolIncludeLowerCase) is an boolean True or False specifying
    '       whether the generated string should include lowercase letters from the ASCII
    '       character set.
    '   The fourth argument (boolIncludeUpperCase) is a boolean True or False specifing
    '       whether the generated string should include uppercase (capital) letters from the
    '       ASCII character set.
    '   The fifth argument (boolIncludeNumbers) is a boolean True or False specifying whether
    '       the generated string should include numbers from the ASCII character set.
    '   The sixth argument (boolIncludePunctuation) is a boolean True or False specifying
    '       whether the generated string should include punctuation from the ASCII character
    '       set. Punctuation includes the following characters:
    '       !"#%&'()*,-./:;?@[\]_{}
    '   The seventh argument (boolIncludeSymbols) is a boolean True or False specifying whether
    '       the generated string should include symbols from the ASCII character set. Symbols
    '       include the following characters:
    '       $+<=>^`|~
    '
    ' The function returns 0 if the random string was generated successfully. A negative number
    ' is returned if the random string was not generated successfully.
    '
    ' Example:
    '   intReturnCode = NewRandomStringOfASCIICharacters(strRandom, 100, True, False, True, False, False)
    '   If intReturnCode = 0 Then
    '       ' The string was generated successfully.
    '       ' strRandom contains a random string like the following:
    '       ' 8lx5gg06sff4h2jtkkxbeagr6rp95qh8lknwe063pvl6vcdeqcwey3lyv10y83pqmutlubbghw703c9y8aofzr8vmanyls3zq537
    '   End If
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
    ' Stack Overflow user ub3rst4r, who posted the following answer that got me on the right
    ' track: https://stackoverflow.com/a/30116847/2134110
    'endregion Acknowledgements ####################################################

    'region DependsOn ####################################################
    ' TestObjectForData()
    ' TestObjectIsAnyTypeOfInteger()
    'endregion DependsOn ####################################################

    Dim intFunctionReturn
    Dim strWorkingRandom
    Dim strCharSet
    Dim intCounter

    Const LOWERCASE = "abcdefghijklmnopqrstuvwxyz"
    Const UPPERCASE = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    Const NUMBERS = "0123456789"
    Const PUNCTUATION = "!""#%&'()*,-./:;?@[\]_{}"
    Const SYMBOLS = "$+<=>^`|~"

    Err.Clear

    intFunctionReturn = 0
    strCharSet = ""

    On Error Resume Next
    If TestObjectForData(boolIncludeLowerCase) = True Then
        If boolIncludeLowerCase = True Then
            strCharSet = strCharSet & LOWERCASE
        End If
    End If
    If TestObjectForData(boolIncludeUpperCase) = True Then
        If boolIncludeUpperCase = True Then
            strCharSet = strCharSet & UPPERCASE
        End If
    End If
    If TestObjectForData(boolIncludeNumbers) = True Then
        If boolIncludeNumbers = True Then
            strCharSet = strCharSet & NUMBERS
        End If
    End If
    If TestObjectForData(boolIncludePunctuation) = True Then
        If boolIncludePunctuation = True Then
            strCharSet = strCharSet & PUNCTUATION
        End If
    End If
    If TestObjectForData(boolIncludeSymbols) = True Then
        If boolIncludeSymbols = True Then
            strCharSet = strCharSet & SYMBOLS
        End If
    End If
    If Err Then
        On Error Goto 0
        Err.Clear
        intFunctionReturn = -1
    Else
        On Error Goto 0
    End If

    If intFunctionReturn = 0 Then
        ' No error occurred yet
        If Len(strCharSet) <= 0 Then
            intFunctionReturn = -2
        Else
            If TestObjectIsAnyTypeOfInteger(intNumberOfCharactersToGenerate) = False Then
                intFunctionReturn = -3
            Else
                If intNumberOfCharactersToGenerate < 0 Then
                    intFunctionReturn = -4
                End If
            End If
        End If
    End If

    If intFunctionReturn = 0 Then
        ' No error occurred yet
        strWorkingRandom = ""
        If intNumberOfCharactersToGenerate > 0 Then
            Randomize
        End If
        For intCounter = 1 To intNumberOfCharactersToGenerate
            ' Int((Max - Min + 1) * Rnd + Min)
            strWorkingRandom = strWorkingRandom & Mid(strCharSet, Int((Len(strCharSet) - 1 + 1) * Rnd + 1), 1)
        Next
    End If

    If intFunctionReturn = 0 Then
        strRandom = strWorkingRandom
    End If
    NewRandomStringOfASCIICharacters = intFunctionReturn
End Function
