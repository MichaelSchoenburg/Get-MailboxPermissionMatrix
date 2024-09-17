<#
.SYNOPSIS
    Gives an overview of all mailbox permission using a matrix table.

.DESCRIPTION
    Attention: You need to login to an Exchange Online Session first, using the PowerShell Exchange Online module.    
    Returns an array of pscustomobjects displaying a matrix of all mailbox permissions by default.
    Outputs the same as a CSV file if a path has been specified using the OutputFile parameter.

.PARAMETER OutputFile
    Specifies the path to the output file. This includes the name of the file, as well as the file extension CSV.

.EXAMPLE
    PS C:\> .\Get-MailboxPermissionMatrix.ps1
    Returns an array of pscustomobjects of all mailbox permissions.

.EXAMPLE
    PS C:\> .\Get-MailboxPermissionMatrix.ps1 | Format-Table
    Returns an array of pscustomobjects of all mailbox permissions and formats it in a matrix table.


.EXAMPLE
    PS C:\> .\Get-MailboxPermissionMatrix.ps1 -OutputFile "C:\output.csv"
    Outputs the CSV file to the specified location.

.INPUTS
    None. You cannot pipe objects to Get-MailboxPermissionMatrix.ps1.

.OUTPUTS
    Returns an array of pscustomobjects if no parameters have been used. 
    Outputs a CSV file if a path has been specified using the OutputFile parameter.

.LINK
    GitHub: https://github.com/MichaelSchoenburg/Get-MailboxPermissionMatrix

.NOTES
    Author: Michael Schönburg
    Version: v1.0
    Last Edit: 03.08.2022
    
    This projects code loosely follows the PowerShell Practice and Style guide, as well as Microsofts PowerShell scripting performance considerations.
    Style guide: https://poshcode.gitbook.io/powershell-practice-and-style/
    Performance Considerations: https://docs.microsoft.com/en-us/powershell/scripting/dev-cross-plat/performance/script-authoring-considerations?view=powershell-7.1
#>

#region INITIALIZATION
<# 
    Libraries, Modules, ...
#>

param (
    [Parameter(
        Mandatory = $false,
        Position = 0
    )]
    [System.IO.FileInfo]
    $OutputFile
)

#endregion INITIALIZATION
#region DECLARATIONS
<#
    Declare local variables and global variables
#>



#endregion DECLARATIONS
#region FUNCTIONS
<# 
    Declare Functions
#>

function Write-ConsoleLog {
    <#
    .SYNOPSIS
    Logs an event to the console.
    
    .DESCRIPTION
    Writes text to the console with the current date (US format) in front of it.
    
    .PARAMETER Text
    Event/text to be outputted to the console.
    
    .EXAMPLE
    Write-ConsoleLog -Text 'Subscript XYZ called.'
    
    Long form
    .EXAMPLE
    Log 'Subscript XYZ called.
    
    Short form
    #>
    [alias('Log')]
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true,
        Position = 0)]
        [string]
        $Text
    )

    # Save current VerbosePreference
    $VerbosePreferenceBefore = $VerbosePreference

    # Enable verbose output
    $VerbosePreference = 'Continue'

    # Write verbose output
    Write-Verbose "$( Get-Date -Format 'MM/dd/yyyy HH:mm:ss' ) - $( $Text )"

    # Restore current VerbosePreference
    $VerbosePreference = $VerbosePreferenceBefore
}

#endregion FUNCTIONS
#region EXECUTION
<# 
    Script entry point
#>

# Requirement: Is the module loaded?
if (-not ( Get-Module ExchangeOnlineManagement )) {
    
    # Requirement: Is the module even installed?
    if (-not ( Get-Module -ListAvailable ExchangeOnlineManagement )) {
        Write-Warning -Message "Exchange Online Management PowerShell Module not installed. Please install via 'Install-Module ExchangeOnlineManagement'"
    }
    Write-Warning -Message "Exchange Online Management PowerShell Module not imported. Please import via 'Import-Module ExchangeOnlineManagement'"
# Requirements satisfied
} else {
    try {
        Write-Verbose "Receiving and filtering list of mailboxes"
        $Mbxs = (Get-Mailbox).where({$_.Name -notlike 'DiscoverySearchMailbox*'})
    } catch [System.Management.Automation.CommandNotFoundException] {
        Write-Warning -Message "Exchange Online Management PowerShell session active. Please connect to Exchange Online via 'Connect-ExchangeOnline'"
    } catch {
        Write-Warning -Message "An unexpected error occurred while getting all mailboxes."
        return $_
    }

    $ProgressPreference = 'Continue'
    $PSStyle.Progress.View = 'Classic'

    try {
        $Columns = @()
        $Counter = 0
        foreach ($user in $Mbxs) {
            $Counter++
            $percent = ( ($Counter / $($Mbxs.Count) ) * 100 )
            Write-Progress -Activity "Processing mailbox ($Counter of $($Mbxs.Count))" -Status $User.DisplayName -PercentComplete $percent -Id 0
            Write-Verbose "Processing mailbox ($Counter of $($Mbxs.Count)) $($User.DisplayName)"
            $CalendarFolders = Get-MailboxFolderStatistics -Identity $user.SamAccountName -FolderScope Calendar
            $Counter2 = 0
            foreach ($folder in $CalendarFolders) {
                $Counter2++
                $percent2 = ( ($Counter2 / $($CalendarFolders.Count) ) * 100 )
                Write-Progress -Activity "Processing calendar ($Counter2 of $($CalendarFolders.Count))" -Status $folder.Name -PercentComplete $percent2 -Id 1 -ParentId 0
                Write-Verbose "∟ Processing calendar ($Counter2 of $($CalendarFolders.Count)) $($folder.Name)"
                $rows = [PSCustomObject]@{}
                $rows | Add-Member -MemberType NoteProperty -Name 'x' -Value "$($user.PrimarySmtpAddress):$($folder.Name)"
                Write-Verbose "  ∟ User.Identity = $($user.Identity)"
                Write-Verbose "  ∟ FolderID = $($folder.FolderId)"
                Write-Verbose "  ∟ FolderPath = $($folder.FolderPath)"
                Write-Verbose "  ∟ Folder.Identity = $($folder.Identity)"
                $CalPermissions = Get-MailboxFolderPermission -Identity "$($user.SamAccountName):$($folder.FolderId)"
                $MbxPermissions = Get-MailboxPermission -Identity $user.SamAccountName
                foreach ($trustee in $Mbxs) {
                    if ($CalPermissions.User.DisplayName -contains $trustee.DisplayName) {
                        $rows | Add-Member -MemberType NoteProperty -Name $trustee.PrimarySmtpAddress -Value 'x'
                    } elseif ($MbxPermissions.User -contains $trustee.PrimarySmtpAddress) {
                        $rows | Add-Member -MemberType NoteProperty -Name $trustee.PrimarySmtpAddress -Value 'o'
                    } else {
                        $rows | Add-Member -MemberType NoteProperty -Name $trustee.PrimarySmtpAddress -Value '-'
                    }
                }
                $Columns += $rows
            }
            Write-Progress -Id 1 -ParentId 0 -Completed
        }
        Write-Progress -Id 0 -Completed
    }
    catch {
        Write-Warning -Message "An error occurred during processing."
        return $_
    }

    try {
        if ($OutputFile) {
            $Columns | Export-Csv -NoTypeInformation -Path $OutputFile
            Write-ConsoleLog "CSV file saved at $( $OutputFile )"
        } else {
            return $Columns
        }
    }
    catch {
        Write-Warning -Message "An error occurred during final output."
        return $_
    }
}

#endregion EXECUTION
