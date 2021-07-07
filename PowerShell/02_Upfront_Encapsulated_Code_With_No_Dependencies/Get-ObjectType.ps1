    #region License ####################################################
    # Copyright (c) 2021 Frank Lesniak
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
    #endregion License ####################################################

    #region DownloadLocationNotice ####################################################
    # The most up-to-date version of this script can be found on the author's GitHub repository
    # at https://github.com/franklesniak/sysadmin-accelerator
    #endregion DownloadLocationNotice ####################################################

function Get-ObjectType {
    #region FunctionHeader ####################################################
    # This function is a "safer" way to call .GetType() on an object; if the object is $null
    # or void/nothing/uninitialized and has no type, this function returns an empty string ('')
    # and avoids generating an error.
    #
    # For example, the following would normally generate an error, which we can avoid by using
    # this function in place of .GetType:
    # function Get-Nothing {}
    # (Get-Nothing).GetType()
    #
    # This function takes one positional parameter: a reference to an object.
    #
    # Example usage:
    # $objectSomething = Do-SomeFunction
    # $strObjectType = Get-ObjectType ([ref]$objectSomething)
    # if ($strObjectType -eq '') {
    #   # $strObjectType has no type
    # } else {
    #   # $strObjectType contains the result of $objectSomething.GetType().TypeName
    # }
    #
    # Version: 1.0.20210105.0
    #endregion FunctionHeader ####################################################

    trap {
        # Intentionally left empty to prevent terminating errors from halting processing
    }

    $refObjectToTest = $args[0]

    # Retrieve the newest error on the stack prior to doing work
    $refLastKnownError = Get-ReferenceToLastError

    # Store current error preference; we will restore it after we do our work
    $actionPreferenceFormerErrorPreference = $global:ErrorActionPreference

    # Set ErrorActionPreference to SilentlyContinue; this will suppress error output.
    # Terminating errors will not output anything, kick to the empty trap statement and then
    # continue on. Likewise, non-terminating errors will also not output anything, but they
    # do not kick to the trap statement; they simply continue on.
    $global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue

    $PSTypeInfo = $null
    $PSTypeInfo = ($refObjectToTest.Value).GetType()

    # Restore the former error preference
    $global:ErrorActionPreference = $actionPreferenceFormerErrorPreference

    # Retrieve the newest error on the error stack
    $refNewestCurrentError = Get-ReferenceToLastError

    if (Test-ErrorOccurred $refLastKnownError $refNewestCurrentError) {
        # Error occurred
        '' # Return empty string
    } else {
        if ($null -eq $PSTypeInfo) {
            '' # Return empty string
        } else {
            $refLastKnownError = $null
            # Retrieve the newest error on the stack prior to doing work
            $refLastKnownError = Get-ReferenceToLastError

            # Store current error preference; we will restore it after we do our work
            $actionPreferenceFormerErrorPreference = $global:ErrorActionPreference

            # Set ErrorActionPreference to SilentlyContinue; this will suppress error output.
            # Terminating errors will not output anything, kick to the empty trap statement and then
            # continue on. Likewise, non-terminating errors will also not output anything, but they
            # do not kick to the trap statement; they simply continue on.
            $global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue

            $strTypeName = $PSTypeInfo.Name

            # Restore the former error preference
            $global:ErrorActionPreference = $actionPreferenceFormerErrorPreference

            $refNewestCurrentError = $null
            # Retrieve the newest error on the error stack
            $refNewestCurrentError = Get-ReferenceToLastError

            if (Test-ErrorOccurred $refLastKnownError $refNewestCurrentError) {
                # Error occurred
                '' # Return empty string
            } else {
                if ([string]::IsNullOrEmpty($strTypeName)) {
                    '' # Return empty string
                } else {
                    $strTypeName
                }
            }
        }
    }
}