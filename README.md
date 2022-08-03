# Get-MailboxPermissionMatrix

    NAME
        .\Get-MailboxPermissionMatrix.ps1

    SYNOPSIS
        Gives an overview of all mailbox permission using a matrix table.


    SYNTAX
        C:\Users\mschoenburg\Desktop\Get-MailboxPermissionMatrix.ps1 [[-OutputFile] <FileInfo>] [<CommonParameters>]


    DESCRIPTION
        Attention: You need to login to an Exchange Online Session first, using the PowerShell Exchange Online module.    
        Returns an array of pscustomobjects displaying a matrix of all mailbox permissions by default.
        Outputs the same as a CSV file if a path has been specified using the OutputFile parameter.


    PARAMETER
        -OutputFile <FileInfo>
            Specifies the path to the output file. This includes the name of the file, as well as the file extension CSV.

            Required?                   false
            Position?                       1
            Default value                 
            Accept pipeline input?      false
            Accept wildcard characters? false

        <CommonParameters>
            This cmdlet supports the common parameters: "Verbose", "Debug",
            "ErrorAction", "ErrorVariable", "WarningAction", "WarningVariable",
            "OutBuffer", "PipelineVariable" und "OutVariable". You can find more information under
            "about_CommonParameters" (http://go.microsoft.com/fwlink/?LinkID=113216). 

    INPUTS
        None. You cannot pipe objects to Get-MailboxPermissionMatrix.ps1.


    OUTPUTS
        Returns an array of pscustomobjects if no parameters have been used. 
        Outputs a CSV file if a path has been specified using the OutputFile parameter.


    NOTES
            Author: Michael Schönburg
            Version: v1.0
            Last Edit: 03.08.2022

            This projects code loosely follows the PowerShell Practice and Style guide, as well as Microsofts PowerShell scripting performance considerations.
            Style guide: https://poshcode.gitbook.io/powershell-practice-and-style/
            Performance Considerations: https://docs.microsoft.com/en-us/powershell/scripting/dev-cross-plat/performance/script-authoring-considerations?view=powershell-7.1

        -------------------------- EXAMPLE 1 --------------------------

        PS C:\>.\Get-MailboxPermissionMatrix.ps1

        Returns an array of pscustomobjects of all mailbox permissions.




        -------------------------- EXAMPLE 2 --------------------------

        PS C:\>.\Get-MailboxPermissionMatrix.ps1 | Format-Table

        Returns an array of pscustomobjects of all mailbox permissions and formats it in a matrix table.




        -------------------------- EXAMPLE 3 --------------------------

        PS C:\>.\Get-MailboxPermissionMatrix.ps1 -OutputFile "C:\output.csv"

        Outputs the CSV file to the specified location.





    RELATED LINKS
        GitHub: https://github.com/MichaelSchoenburg/Get-MailboxPermissionMatrix



