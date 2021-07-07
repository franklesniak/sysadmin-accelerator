function Test-ObjectForData {
    #region TestObjectForDataFunctionMetadata
    ###########################################################################################
    # Checks an object or variable to see if it "has data".
    # If any of the following are true, then objToCheck is regarded as NOT having data:
    #   objToCheck has no type (i.e., $objToCheck.GetType() would throw an error)
    #   $null -eq objToCheck
    #   function Get-Nothing {}; $objToCheck = Get-Nothing
    #   $objToCheck -eq ''
    #   $objToCheck is an empty array ($objToCheck is an array and $objToCheck.Count -eq 0)
    #   $objToCheck is an array of empty strings and/or $nulls
    #   $objToCheck is an array of arrays of empty strings/$nulls, etc.
    #   $objToCheck is of type System.DBNull
    #   $objToCheck is of type System.Management.Automation.Language.NullString
    # In any of these cases, the function returns False. Otherwise, it returns True.
    #
    # The function takes at least one and up to five total positional arguments:
    #   The first positional argument is a reference to the object to be tested for the
    #       presence of data. See the example below for syntax.
    #   The second positional argument is optional. If supplied, it is a boolean value that
    #       states whether warning messages should sent to the warning stream by this function
    #       when applicable. For example, a warning would be generated if an array has more
    #       elements than this function is allowed to test. If not supplied, the function
    #       defaults to showing warning messages.
    #   The third positional argument is optional. If supplied, it is an integer value that
    #       states the maximum number of recursive calls allowed (i.e., the maximum number of
    #       nested arrays). When the maximum number of nested arrays is reached during
    #       evaluation, the function uses [string]::IsNullOrEmpty() on the remainder to
    #       determine if data is present. If not supplied, the function defaults to a maximum
    #       recursive depth of 3.
    #   The fourth positional argument is optional. If supplied, it is an integer value that
    #       states the maximum number of array elements to check. If not supplied, the function
    #       defaults to 500. If the maximum number of array elements is evaluated and no data
    #       has been found, the function displays a warning if the second positional argument
    #       is set to $true, and then returns $false.
    #   The fifth positional argument is for internal use by the function only. It is an
    #       integer that indicates how deep in the recursive call chain the current function
    #       is. Users of this function should not supply this fifth argument. It defaults to 0.
    #
    #   The function returns $true if data is found as per the introductory paragraph; $false
    #       otherwise.
    #
    #   Example:
    #   $result = Do-Something
    #   $boolDataPresent = Test-ObjectForData ([ref]$result)
    #   if ($boolDataPresent) {
    #       # $result contains data
    #   } else {
    #       # $result did not contain any data
    #   }
    #
    # Version: 1.0.20200105.1
    ###########################################################################################
    #endregion TestObjectForDataFunctionMetadata

    #region License
    ###########################################################################################
    # Copyright 2021 Frank Lesniak
    #
    # Permission is hereby granted, free of charge, to any person obtaining a copy of this
    # software and associated documentation files (the "Software"), to deal in the Software
    # without restriction, including without limitation the rights to use, copy, modify, merge,
    # publish, distribute, sublicense, and/or sell copies of the Software, and to permit
    # persons to whom the Software is furnished to do so, subject to the following conditions:
    #
    # The above copyright notice and this permission notice shall be included in all copies or
    # substantial portions of the Software.
    #
    # THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
    # INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
    # PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
    # FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
    # OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
    # DEALINGS IN THE SOFTWARE.
    ###########################################################################################
    #endregion License

    #region DownloadLocationNotice
    ###########################################################################################
    # The most up-to-date version of this script can be found on the author's GitHub repository
    # at https://github.com/franklesniak/Test_Object_For_Data
    ###########################################################################################
    #endregion DownloadLocationNotice

    #region Acknowledgements ####################################################
    # Thanks to Scott Dexter for writing the article "Empty Nothing And Null How Do You Feel
    # Today", which inspired me to create this function originally in VBScript:
    # https://evolt.org/node/346
    #
    # Thanks to Kevin Marquette for providing guidance on how PowerShell handles $null and
    # "empty $null". You also helped me to define test cases to validate that this function
    # works as intended, and pointed me toward [string]::IsNullOrEmpty()
    # https://powershellexplained.com/2018-12-23-Powershell-null-everything-you-wanted-to-know/
    #
    # Finally, thanks to Cody Konior for his article that pointed out DBNull and NullString as
    # alternative forms of "null" in PowerShell:
    # https://www.codykonior.com/2013/10/17/checking-for-null-in-powershell/
    #endregion Acknowledgements ####################################################

    $refObjectToTest = $args[0]
    if ($args.Count -gt 1) {
        $boolDisplayWarnings = $args[1]
    } else {
        $boolDisplayWarnings = $true
    }
    if ($args.Count -gt 2) {
        $intMaxNestedCalls = $args[2]
    } else {
        $intMaxNestedCalls = 3
    }
    if ($args.Count -gt 3) {
        $intMaxArrayItemsToCheck = $args[3]
    } else {
        $intMaxArrayItemsToCheck = 500
    }
    if ($args.Count -gt 4) {
        $intNestedCount = $args[4]
    } else {
        $intNestedCount = 0
    }

    # Get the object type name without throwing an error in the case of $null or "nothing"
    $strTypeName = Get-ObjectType $refObjectToTest

    $boolResult = $true

    if ([string]::IsNullOrEmpty($strTypeName)) {
        # A blank string ('') returned from Get-ObjectType indicates no type information
        $boolResult = $false
    }

    if ($boolResult -eq $true) {
        if ($strTypeName.Contains('[]')) {
            # Array
            if ($null -eq ($refObjectToTest.Value).Count) {
                # An array should never have a null .Count, but just in case, we handle it
                $boolResult = $false
            } else {
                if (($refObjectToTest.Value).Count -eq 0) {
                    # Empty array
                    $boolResult = $false
                } else {
                    if ($intNestedCount -ge $intMaxNestedCalls) {
                        # too many recursive calls have been made
                        # must be arrays in side of arrays inside of arrays...
                        if ([string]::IsNullOrEmpty($refObjectToTest.Value)) {
                            $boolResult = $false
                        } # else $true
                    } else {
                        $boolResult = $false
                        for ($intCounterA = 0; ($intCounterA -lt ($refObjectToTest.Value).Count) -and ($intCounterA -lt $intMaxArrayItemsToCheck); $intCounterA++) {
                            $refArrayMember = [ref](($refObjectToTest.Value)[$intCounterA])
                            $boolInterimResult = Test-ObjectForData $refArrayMember $boolDisplayWarnings $intMaxNestedCalls $intMaxArrayItemsToCheck ($intNestedCount + 1)
                            if ($boolInterimResult) {
                                $boolResult = $true
                                break
                            }
                        }

                        if (($boolResult -eq $false) -and (($refObjectToTest.Value).Count -gt $intMaxArrayItemsToCheck)) {
                            if ($boolDisplayWarnings) {
                                Write-Warning ('Test-ObjectForData only checked the first ' + [string]$intMaxArrayItemsToCheck + ' items in the array for data, but the function found none. Function is returning $false')
                            }
                        }
                    }
                }
            }
        } else {
            # Not an array
            $strTypeName = $strTypeName.ToLower()
            if ($strTypeName -eq 'string') {
                if ([string]::IsNullOrEmpty($refObjectToTest.Value)) {
                    $boolResult = $false
                } # else $true
            } elseif (($strTypeName -eq 'dbnull') -or ($strTypeName -eq 'system.dbnull')) {
                # Type System.DBNull, e.g.: [System.DBNull]::Value
                $boolResult = $false
            } elseif (($strTypeName -eq 'nullstring') -or ($strTypeName -eq 'system.management.automation.language.nullstring')) {
                # Type System.Management.Automation.Language.NullString, e.g.:
                # [System.Management.Automation.Language.NullString]::Value
                $boolResult = $false
            }
        }
    }

    $boolResult
}