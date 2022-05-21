function Get-OutOfDateModules {
    #region FunctionMetadata ################################################
    # This function compares the installed PowerShell modules to those of the
    # PowerShell session's configured PSRespositor(ies).My default, this function
    # compares installed PowerShell modules to those listed in the PSGallery. If the
    # version listed in the PSRepository is newer to that which is installed, then this
    # function returns the newer module information. All modules needing updates are
    # assembled and this function returns an array.
    #
    # This function takes no input.
    #
    # Example usage:
    # $arrModulesToInstall = Get-OutOfDateModules
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
    # Jeff Hicks posted similar code that inspired this function
    #endregion Acknowledgements ################################################

    #region DependsOn ################################################
    # None!
    #endregion DependsOn ################################################

    # TODO: validate that the current version of PowerShell is 5.0 or greater

    $hashtableNameSortOrder = @{}
    $hashtableNameSortOrder.Add('expression', 'Name')
    $hashtableNameSortOrder.Add('descending', $false)
    $hashtableVersionSortOrder = @{}
    $hashtableVersionSortOrder.Add('expression', 'Version')
    $hashtableVersionSortOrder.Add('descending', $true)
    $arrCustomSort = @($hashtableNameSortOrder, $hashtableVersionSortOrder)

    Write-Verbose 'Collecting information on the currently-installed PowerShell modules...'
    $arrInstalledModules = @(Get-Module -ListAvailable |
            Sort-Object -Property $arrCustomSort -Unique)

    $arrModulesToCheckForUpdate = @($arrInstalledModules | ForEach-Object { $_.Name } |
            Sort-Object -Unique)

    Write-Verbose 'Checking the PowerShell Module Gallery for available module versions...'
    $arrNewestModules = @($arrModulesToCheckForUpdate | ForEach-Object {
            Find-Module -Name $_ -ErrorAction SilentlyContinue
        })

    Write-Verbose 'Determining which of the installed modules need to be updated...'
    $result = @($arrNewestModules | ForEach-Object {
            $objNewestModuleFromGallery = $_
            $strModuleName = $objNewestModuleFromGallery.Name
            $versionNewestFromGallery = [version]($objNewestModuleFromGallery.Version)

            $arrInstalledVersions = @($arrInstalledModules |
                    Where-Object { $_.Name -eq $strModuleName } | 
                    ForEach-Object { $_.Version } |
                    Sort-Object -Descending
                )

                if ($arrInstalledVersions.Count -ge 1) {
                    # $arrInstalledVersions[0] is the newest-installed version
                    if ($arrInstalledVersions[0] -lt $versionNewestFromGallery) {
                        # The currently installed version is older than that of the gallery
                        return $objNewestModuleFromGallery
                    }
                }
            })

    # The following code forces the function to return an array, always, even when
    # there are zero or one elements in the array
    $intElementCount = 1
    if ($null -ne $result) {
        if ($result.GetType().FullName.Contains('[]')) {
            if (($result.Count -ge 2) -or ($result.Count -eq 0)) {
                $intElementCount = $result.Count
            }
        }
    }
    $strLowercaseFunctionName = $MyInvocation.InvocationName.ToLower()
    $boolArrayEncapsulation = $MyInvocation.Line.ToLower().Contains('@(' + $strLowercaseFunctionName + ')') -or $MyInvocation.Line.ToLower().Contains('@(' + $strLowercaseFunctionName + ' ')
    if ($boolArrayEncapsulation) {
        $result
    } elseif ($intElementCount -eq 0) {
        , @()
    } elseif ($intElementCount -eq 1) {
        , (, $result)
    } else {
        $result
    }
}
