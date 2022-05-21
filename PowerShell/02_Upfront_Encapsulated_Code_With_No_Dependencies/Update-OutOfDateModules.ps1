function Update-OutOfDateModules {
    #region FunctionMetadata ################################################
    # After identifying the PowerShell modules that are out of date (via
    # Get-OutOfDateModules), this function proceeds with updating those that are out
    # of date
    #
    # This function takes no input and generates no output. Informational messages are
    # written to the verbose pipeline
    #
    # Example usage:
    # $actionPreferenceFormerVerbose = $VerbosePreference
    # $VerbosePreference = [Management.Automation.ActionPreference]::Continue
    # Update-OutOfDateModules
    # $VerbosePreference = $actionPreferenceFormerVerbose
    #
    # Version: 1.0.20220521.0
    #endregion FunctionMetadata ################################################

    #region License ################################################
    # Copyright (c) 2022 Frank Lesniak
    #
    # Permission is hereby granted, free of charge, to any person obtaining a copy of
    # this software and associated documentation files (the "Software"), to deal in the
    # Software without restriction, including without limitation the rights to use,
    # copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the
    # Software, and to permit persons to whom the Software is furnished to do so,
    # subject to the following conditions:
    #
    # The above copyright notice and this permission notice shall be included in all
    # copies or substantial portions of the Software.
    #
    # THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    # IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
    # FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR
    # COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN
    # AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
    # WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
    #endregion License ################################################

    #region DownloadLocationNotice ################################################
    # The most up-to-date version of this script can be found on the author's GitHub repository
    # at https://github.com/franklesniak/sysadmin-accelerator
    #endregion DownloadLocationNotice ################################################

    #region Acknowledgements ################################################
    # None!
    #endregion Acknowledgements ################################################

    #region DependsOn ################################################
    # Get-OutOfDateModules
    #endregion DependsOn ################################################

    # TODO: validate that the current version of PowerShell is 5.0 or greater
    # TODO: handle errors with Install-Module
    # TODO: handle the error from Install-Module: Install-Package: The following
    #       commands are already available on this system:'Add-MetadataConverter,
    #       ConvertFrom-Metadata,ConvertTo-Metadata,Export-Metadata,Get-ManifestValue,
    #       Import-Metadata,Update-Manifest'. This module 'Metadata' may override the
    #       existing commands. If you still want to install this module 'Metadata', use
    #       -AllowClobber parameter.

    $arrModuleUpdatesAvailable = Get-OutOfDateModules

    # First, determine if we need to update the Az module or any of its subcomponents
    $boolAzRootModuleFound = $false
    $intCountOfAzModules = 0
    $arrModuleUpdatesAvailable | ForEach-Object {
        if ($_.Name -eq 'Az') {
            $boolAzRootModuleFound = $true
            break
        } elseif ($_.Name -like 'Az.*') {
            $intCountOfAzModules++
        }
    }

    if ($boolAzRootModuleFound -or $intCountOfAzModules -gt 10) {
        # If the root 'Az' module or more than 10 subcomponents are out of date, then it
        # makes sense to just update the whole thing at once:
        Write-Verbose ('Updating Module: Az (note: this may take a while!)')
        Install-Module -Name Az -Force
        # Now we can install other updates:
        $arrModuleUpdatesAvailable | ForEach-Object {
            $objThis = $_
            if ($objThis.Name -ne 'Az' -and $objThis.Name -notlike 'Az.*') {
                Write-Verbose ('Updating Module: ' + $objThis.Name)
                Install-Module -Name $objThis.Name -Force
            }
        }
    } else {
        $arrModuleUpdatesAvailable | ForEach-Object {
            Write-Verbose ('Updating Module: ' + $_.Name)
            Install-Module -Name $_.Name -Force
        }
    }
}
