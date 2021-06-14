Function ReadCharactersFromFile(ByRef strData, ByVal strPathToFile, ByVal lngMaxNumberOfCharactersToRead, ByVal boolContinueOnError)
    'region FunctionMetadata ####################################################
    ' Safely reads-in characters from a file at path strPathToFile and stores them in a string
    ' (strData)
    '
    ' This function takes four arguments:
    '   - The first argument strData) is populated upon success with a string containing all
    '       of the characters that were read-in from the file.
    '   - The second argument (strPathToFile) is a string containing the path to the file to be
    '       read-in by this function.
    '   - The third argument (lngMaxNumberOfCharactersToRead) allows the caller to set an upper
    '       boundary on the number of characters read-in from the file. It can be set to an
    '       integrer, or set to Null if there is no limit.
    '   - The fourth argument (boolContinueOnError) allows the caller to specify whether the
    '       function should continue reading-in characters if the operation to read one
    '       character fails. If set to True, the process continues and drops the character that
    '       resulted in a read error. If set to False or Null, the process would stop on error
    '       and return a numerical code indicating failure (see below).
    '
    ' The function returns 0 if the characters were read-in from the specified file
    ' successfully; it returns a negative number if the characters were not able to be read
    '
    ' Example:
    '   intReturnCode = ReadCharactersFromFile(strData, "C:\Users\flesniak\Desktop\TestFile.txt", 60000, Null)
    '   If intReturnCode = 0 Then
    '       ' The file was read successfully and was capped at a maximum of 60,000 characters
    '       ' strData contains the characters read from the file
    '   End If
    '
    ' Version: 1.2.20210614.0
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
    ' StackExchange user "jumpjack", for the post that inspired the creation of this function:
    ' https://superuser.com/a/1027161/334370
    'endregion Acknowledgements ####################################################

    'region DependsOn ####################################################
    ' TestObjectForData()
    'endregion DependsOn ####################################################

    Dim intFunctionReturn
    Dim boolWorkingContinueOnError
    Dim fileSystemObject
    Dim boolTest
    Dim fileObjectSource
    Dim textStreamObjectSource
    Dim strWorkingOutput
    Dim lngCounter
    Dim boolBreakOut

    Err.Clear
    
    intFunctionReturn = 0

    If TestObjectForData(boolContinueOnError) = False Then
        boolWorkingContinueOnError = False
    Else
        On Error Resume Next
        boolTest = (boolContinueOnError = True)
        If Err Then
            On Error Goto 0
            Err.Clear
            intFunctionReturn = -1
        Else
            On Error Goto 0
            If boolTest Then
                boolWorkingContinueOnError = True
            Else
                boolWorkingContinueOnError = False
            End If
        End If
    End If

    On Error Resume Next
    Set fileSystemObject = CreateObject("Scripting.FileSystemObject")
    If Err Then
        On Error Goto 0
        Err.Clear
        intFunctionReturn = -2
    Else
        boolTest = fileSystemObject.FileExists(strPathToFile)
        If Err Then
            On Error Goto 0
            Err.Clear
            intFunctionReturn = -3
        Else
            On Error Goto 0
            If boolTest = False Then
                ' File specified by strPathToFile did not exist
                intFunctionReturn = -4
            Else
                On Error Resume Next
                Set fileObjectSource = fileSystemObject.GetFile(strPathToFile)
                If Err Then
                    On Error Goto 0
                    Err.Clear
                    intFunctionReturn = -5
                Else
                    On Error Goto 0
                End If
            End If
        End If
    End If

    If intFunctionReturn = 0 Then
        ' No error occurred yet
        ' fileObjectSource is a FileObject consisting of the source file
        On Error Resume Next
        Set textStreamObjectSource = fileObjectSource.OpenAsTextStream()
        If Err Then
            On Error Goto 0
            Err.Clear
            intFunctionReturn = -6
        Else
            On Error Goto 0
            strWorkingOutput = ""
            lngCounter = CLng(0)
            boolBreakOut = False
            On Error Resume Next
            boolTest = ((textStreamObjectSource.AtEndOfStream = False) And (lngCounter < lngMaxNumberOfCharactersToRead) And (boolBreakOut = False))
            If Err Then
                On Error Goto 0
                Err.Clear
                intFunctionReturn = -7
            Else
                While boolTest = True
                    strWorkingOutput = strWorkingOutput + textStreamObjectSource.Read(1)
                    If Err Then
                        Err.Clear
                        If boolWorkingContinueOnError = False Then
                            On Error Goto 0
                            intFunctionReturn = -8
                            boolBreakOut = True
                            boolTest = False
                        End If
                    End If
                    lngCounter = lngCounter + 1
                    boolTest = ((textStreamObjectSource.AtEndOfStream = False) And (lngCounter < lngMaxNumberOfCharactersToRead) And (boolBreakOut = False))
                    If Err Then
                        On Error Goto 0
                        Err.Clear
                        intFunctionReturn = -9
                        boolBreakOut = True
                        boolTest = False
                    End If
                Wend
                If Err Then
                    Err.Clear
                End If
                On Error Resume Next
                textStreamObjectSource.Close
                If Err Then
                    On Error Goto 0
                    Err.Clear
                Else
                    On Error Goto 0
                End If
            End If
        End If
    End If

    If intFunctionReturn = 0 Then
        strData = strWorkingOutput
    End If

    ReadCharactersFromFile = intFunctionReturn
End Function
