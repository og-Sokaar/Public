<#
+---------------------------------------------------------------------------------------------+ 
| PASSWORD CHECK AND SEND MAIL SCRIPT                                                         | 
+---------------------------------------------------------------------------------------------+ 
|                                                                                             |
|   AUTHOR      : og-sokaar                                                                   |
|   DESCRIPTION : PS Script to check the last time a password was changed and send email      |
|                 to user                                                                     |
+---------------------------------------------------------------------------------------------+ 
#>

# Variable declarations
$SMTPServer = "<>"
$SMTPFrom = "<>"
$SMTPSubject = "Reminder! Time to change your password"
$LogFile = "C:\PS-Logs\Passwordcheck.log"
$Userlist = $null

# Check to see if the log folder exist - if it doesn't then create it
if( -Not (Test-Path -Path c:\PS-Logs ) )
{
    New-Item -ItemType directory -Path c:\PS-Logs\ | Out-Null
}

# Create function to write to log file.
Function Write-Log {
    [CmdletBinding()]
    Param(
    [Parameter(Mandatory=$False)]
    [ValidateSet("INFO","WARN","ERROR","FATAL","DEBUG")]
    [String]
    $Level = "INFO",

    [Parameter(Mandatory=$True)]
    [string]
    $Message,

    [Parameter(Mandatory=$False)]
    [string]
    $logfile
    )

    $Stamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
    $Line = "$Stamp $Level $Message"
    If($logfile) {
        Add-Content $logfile -Value $Line
    }
    Else {
        Write-Output $Line
    }
}


# Connect to O365
Connect-MsolService

# Get all users and store in $userlist Variable
$Userlist = Get-MsolUser -all | Sort-Object displayname

# Check when o365 users password was last set
foreach ($User in $Userlist)
{
    if($User.lastpasswordchangetimestamp -le ((Get-Date).AddDays(-30)))
        # Write log file and Send email to any user that has not changed their password in 30 days
        {
        Write-Log -Message "$($User.DisplayName) has not changed their password since $($User.lastpasswordchangetimestamp). A Reminder email has been sent" -Level INFO -logfile $LogFile 
        Send-MailMessage -From "<$($SMTPFrom)>" -To "<$($user.UserPrincipalName)>" -Subject $SMTPSubject -Body "Hi $($user.DisplayName), our records show that your password has not been change in 30 days. Please log into Office365 and change it as soon as it is convenient. This email was sent from an unmonitored mailbox - Please do not reply to this email." -SmtpServer $SMTPServer
        }
}

cls
Write-Host "Operation completed. A log has been created: $($LogFile)."
Read-Host "Press Enter to continue"
