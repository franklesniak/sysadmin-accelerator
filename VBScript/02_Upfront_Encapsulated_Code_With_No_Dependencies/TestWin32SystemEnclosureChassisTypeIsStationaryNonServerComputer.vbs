Function TestWin32SystemEnclosureChassisTypeIsStationaryNonServerComputer(ByVal intChassisType)
    'region FunctionMetadata ####################################################
    ' This function tests the supplied parameter (intChassisType) to determine if it represents
    ' the chassis type of a stationary, non-server computer (desktop or similar).
    '
    ' The function takes one positional argument (intChassisType), which must be an integer
    ' from the ChassisTypes property of an instance of a Win32_SystemEnclosure.
    '
    ' The function returns a boolean True if the chassis type is a stationary, non-server
    ' computer (desktop). It returns a boolean False if the chassis type is not a stationary,
    ' non-server computer (i.e., it is a portable computer, a server, or an unknown type), or
    ' if the supplied parameter was not a valid integer.
    '
    ' Example:
    '   intReturnCode = GetSystemEnclosureInstances(arrSystemEnclosureInstances)
    '   If intReturnCode > 0 Then
    '       ' At least one system enclosure instance was retrieved successfully
    '       For Each objSystemEnclosureInstance in arrSystemEnclosureInstances
    '           arrChassisTypes = objSystemEnclosureInstance.ChassisTypes
    '           For Each intChassisType in arrChassisTypes
    '               If TestWin32SystemEnclosureChassisTypeIsStationaryNonServerComputer(intChassisType) = True Then
    '                   ' The enclosure type is a stationary, non-server computer
    '               Else
    '                   ' The enclosure type is not a stationary, non-server computer
    '               End If
    '           Next
    '       Next
    '   End If
    '
    ' Version: 1.0.20210625.1
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
    ' Microsoft, for (intentionally or not) making the Microsoft Deployment Toolkit (MDT) with
    ' its source code viewable based on it being written in VBS/WSH. MDT has a function that
    ' determines whether a given system is a desktop/laptop/server/VM, which was useful in
    ' determining how to approach this function
    '
    ' DMTF, for publishing the SMBIOS standard, which defines what each of the chassis type
    ' integers mean, and publishing documentation on their website
    'endregion Acknowledgements ####################################################

    'region DependsOn ####################################################
    ' TestObjectIsAnyTypeOfInteger()
    'endregion DependsOn ####################################################

    Dim boolInterimResult

    Const WIN32_SYSTEMENCLOSURE_CHASSISTYPE_DESKTOP = 3
    Const WIN32_SYSTEMENCLOSURE_CHASSISTYPE_LOWPROFILEDESKTOP = 4
    ' Note: a "pizza box" is a desktop chassis, not a server chassis e.g., the SPARCstation 10:
    Const WIN32_SYSTEMENCLOSURE_CHASSISTYPE_PIZZABOX = 5
    Const WIN32_SYSTEMENCLOSURE_CHASSISTYPE_MINITOWER = 6
    Const WIN32_SYSTEMENCLOSURE_CHASSISTYPE_TOWER = 7
    Const WIN32_SYSTEMENCLOSURE_CHASSISTYPE_ALLINONE = 13
    Const WIN32_SYSTEMENCLOSURE_CHASSISTYPE_SPACESAVING = 15
    Const WIN32_SYSTEMENCLOSURE_CHASSISTYPE_LUNCHBOX = 16
    Const WIN32_SYSTEMENCLOSURE_CHASSISTYPE_MINIPC = 35
    Const WIN32_SYSTEMENCLOSURE_CHASSISTYPE_STICKPC = 36

    boolInterimResult = False

    If TestObjectIsAnyTypeOfInteger(intChassisType) <> True Then
        boolInterimResult = False
    Else
        ' intChassisType was an integer
        Select Case intChassisType
            Case WIN32_SYSTEMENCLOSURE_CHASSISTYPE_DESKTOP
                boolInterimResult = True
            Case WIN32_SYSTEMENCLOSURE_CHASSISTYPE_LOWPROFILEDESKTOP
                boolInterimResult = True
            Case WIN32_SYSTEMENCLOSURE_CHASSISTYPE_PIZZABOX
                boolInterimResult = True
            Case WIN32_SYSTEMENCLOSURE_CHASSISTYPE_MINITOWER
                boolInterimResult = True
            Case WIN32_SYSTEMENCLOSURE_CHASSISTYPE_TOWER
                boolInterimResult = True
            Case WIN32_SYSTEMENCLOSURE_CHASSISTYPE_ALLINONE
                boolInterimResult = True
            Case WIN32_SYSTEMENCLOSURE_CHASSISTYPE_SPACESAVING
                boolInterimResult = True
            Case WIN32_SYSTEMENCLOSURE_CHASSISTYPE_LUNCHBOX
                boolInterimResult = True
            Case WIN32_SYSTEMENCLOSURE_CHASSISTYPE_MINIPC
                boolInterimResult = True
            Case WIN32_SYSTEMENCLOSURE_CHASSISTYPE_STICKPC
                boolInterimResult = True
            Case Else
                boolInterimResult = False
        End Select
    End If
    
    TestWin32SystemEnclosureChassisTypeIsStationaryNonServerComputer = boolInterimResult
End Function
