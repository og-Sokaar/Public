<#
+---------------------------------------------------------------------------------------------+ 
| USER CLEAN UP SCRIPT                                                                        | 
+---------------------------------------------------------------------------------------------+ 
|                                                                                             |
|   AUTHOR      : og-sokaar                                                                   |
|   DESCRIPTION : Script to check if a user has logged on in the last 60 days. If not         |
|                 move user to correct OU, delete profile and email HR                        |
+---------------------------------------------------------------------------------------------+ 
#>

# Declare Variables

$maxage = (Get-Date).AddDays(-90)
$disableddate = (Get-Date).AddDays(-90).ToString('dd/MM/yy')
$deletiondate = (Get-Date).AddDays(30).ToString('dd/MM/yy')
$creationdate = (Get-Date).AddMonths(-6)
$date = Get-Date -f d
$disabledou = "OU=Disabled Users,DC=contoso,DC=local"
$errorlog = "C:\temp\StaleADUsersErrors$((Get-Date).ToString('ddMMyy')).log"
$csv = "C:\temp\StaleADUsers$((Get-Date).ToString('ddMMyy')).csv"
$props = @()
$mailpass = ConvertTo-SecureString "whatever" -AsPlainText -Force
$mailcreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList "NT AUTHORITY\ANONYMOUS LOGON", $mailpass
$mailparams = @{
    from = "<insert sender address>"
    to = "<insert destination address"
    smtpserver = "<SMTP Server>"
    port = "25"
    subject = "Stale AD Users - $($date)"
    bodyashtml = $true
    Credential = $mailcreds
}

$css = 'h1 {
    margin-left: auto;
    margin-right: auto;
    text-transform: uppercase;
    text-align: center;
    font-size: 13pt;
    font-weight: bold;
    }
    
    h2 {
        margin-left: auto;
        margin-right: auto;
        text-transform: capitalize;
        text-align: center;
        font-family: "Segoe UI";
        font-size: 14pt;
        font-weight: bold;
    }


 


    body {
        margin-left: auto;
        margin-right: auto;
        text-align: center;
        font-family: "Segoe UI";
        font-weight: lighter;
        font-size: 9pt;
        color:#2f2f2f;
        background-color: white;
    }

    
    table {
        margin-left: auto;
        margin-right: auto;
        border-width: 1px;
        border-style: solid;
        border-color: #2f2f2f;
        border-collapse: collapse;
    }


    th {
        font-family: "Segoe UI";
        font-weight: lighter;
        color: white;
        text-transform: capitalize;
        margin-left: auto;
        margin-right: auto;
        border-width: 1px;
        border-style: solid;
        border-color: #2f2f2f;
        background-color: #d32f2f;
    }


    td {
        margin-left: auto;
        margin-right: auto;
        border-width: 1px;
        border-style: solid;
        border-color: #2f2f2f;
        background-color: white;
    }

'

$params = @{
    Server = "<insert target DC>"
}

$deleteduserserrorlog = "C:\temp\DeletedADUsersErrors$((Get-Date).ToString('ddMMyy')).log"
$deleteduserscsv = "C:\temp\DeletedADUsers$((Get-Date).ToString('ddMMyy')).csv"
$deletedusers = @()
$date = Get-Date -f d

# Get Active Directory users and properties - filter out anything not relevant

$users = Get-ADUser -Filter {mail -like "*"} -Properties title,department,enabled,mail,manager,lastlogontimestamp,pwdLastSet | ?{
    $_.department -notlike "service account" -and
    $_.title -ne $null -and
    $_.SamAccountName -notlike "*test*" -and
    $_.DistinguishedName -notlike "*room*" -and
    $_.DistinguishedName -notlike "*service account*" -and
    $_.DistinguishedName -notlike "*mailboxes*" -and
    $_.DistinguishedName -notlike "*secondary accounts*" -and
    $_.DistinguishedName -notlike "*external user access*" -and
    $_.DistinguishedName -notlike "*breakout*" -and
    $_.DistinguishedName -notlike "*central accounts*" -and
    $_.DistinguishedName -notlike "*CN=Users,DC=*" -and
    $_.DistinguishedName -notlike "*OU=Disabled Users*" -and
    $_.SamAccountName -notlike "*template" -and
    [datetime]::FromFileTime($_.lastlogontimestamp) -lt $maxage -and
    $_.pwdLastSet -ne 0
}


# ForEach loop - Check to find the users manager. Store manager in Variable "$manager"

foreach($user in $users){
        $manager = ""
        $manager = $user.Manager
        if($manager){
            try{
                $usermanager = $user.Manager
                $manager = (Get-ADUser $usermanager -ErrorAction Ignore @params).Name
            }catch{
                $manager = "Not found"
            }
        }else{
            $manager = "Not found"
        }
         # Set AD Users account to disabled, move to correct OU output to shell window the actions taken. If error store in error log
    try{
        Set-ADUser $user -Enabled:$false -Description "Disabled - $($date)" -ErrorAction Stop @params -WhatIf
        Move-ADObject $user -TargetPath $disabledou -Confirm:$False -ErrorAction Stop @params -WhatIf
        Write-Verbose "$($user.SamAccountName) has been disabled and move to $disabledou. This user will be deleted on $($deletiondate)" -Verbose
        $props += [pscustomobject]@{
            Name = $user.Name
            Username = $user.SamAccountName
            Title = $user.Title
            Department = $user.department
            Mail = $user.mail
            Manager = $manager
            LastLogon = [datetime]::FromFileTime($user.lastlogontimestamp)
            PasswordLastSet = [datetime]::FromFileTime($user.pwdlastset)
            FutureDeletionDate = $deletiondate
        }
    }catch{
        Write-Verbose "Error disabling/moving $($user.SamAccountName),$($_.exception.message)" -Verbose
        "Error disabling/moving $($user.SamAccountName)" | Out-File $errorlog -Append
    }

}

# Get Active Directory users and properties for none starter accounts - filter out anything not relevant

$usersnotstarted = Get-ADUser -Filter {mail -like "*"} -Properties * | ?{
    $_.department -notlike "service account" -and
    $_.title -ne $null -and
    $_.SamAccountName -notlike "*test*" -and
    $_.DistinguishedName -notlike "*room*" -and
    $_.DistinguishedName -notlike "*service account*" -and
    $_.DistinguishedName -notlike "*mailboxes*" -and
    $_.DistinguishedName -notlike "*secondary accounts*" -and
    $_.DistinguishedName -notlike "*external user access*" -and
    $_.DistinguishedName -notlike "*breakout*" -and
    $_.DistinguishedName -notlike "*central accounts*" -and
    $_.DistinguishedName -notlike "*CN=Users,DC=*" -and
    $_.DistinguishedName -notlike "*OU=Disabled Users*" -and
    $_.SamAccountName -notlike "*template" -and
    $_.pwdlastset -eq 0 -and
    $_.WhenCreated -lt $creationdate -and
    $_.LastLogonDate -eq $null
}



# ForEach loop - Check to find the users manager. Store manager in Variable "$manager"

foreach($user in $usersnotstarted){
        $manager = ""
        $manager = $user.Manager
        if($manager){
            try{
                $usermanager = $user.Manager
                $manager = (Get-ADUser -filter "DistinguishedName -eq '$($usermanager)'" -ErrorAction Ignore @params).Name
            }catch{
                $manager = "Not found"
            }
        }else{
            $manager = "Not found"
        }   
        # Set AD Users account to disabled change description, move to correct OU output to shell window the actions taken. If error store in error log     
    try{
        Set-ADUser $user -Enabled:$false -Description "Disabled - $($date)" -ErrorAction Stop @params -WhatIf
        Move-ADObject $user -TargetPath $disabledou -Confirm:$False -ErrorAction Stop @params -WhatIf
        Write-Verbose "$($user.SamAccountName) has been disabled and move to $disabledou. This user will be deleted on $($deletiondate)" -Verbose
        $props += [pscustomobject]@{
            Name = $user.Name
            Username = $user.SamAccountName
            Title = $user.Title
            Department = $user.department
            Mail = $user.mail
            Manager = $manager
            LastLogon = [datetime]::FromFileTime($user.lastlogontimestamp)
            PasswordLastSet = [datetime]::FromFileTime($user.pwdlastset)
            DeletionDate = $deletiondate
        }
    }catch{
        Write-Verbose "Error disabling/moving $($user.SamAccountName),$($_.exception.message)" -Verbose
        "Error disabling/moving $($user.SamAccountName)" | Out-File $errorlog -Append
    }
}

$props | Export-Csv $csv -NoTypeInformation -Append
$userstodelete = Get-ADUser -SearchBase $disabledou -Filter "Description -like 'Disabled - *$($disableddate)*'" -SearchScope 1
foreach($user in $userstodelete){
    try{
        Remove-ADUser $user -Confirm:$false -WhatIf
        Remove-Item -Path "\\lowell2\shares\profiles\$($user.SamAccountName)" -Force -WhatIf
        Remove-Item -path "\\lowell2.local\shares\Users_home\$($user.SamAccountName)$" -Force -WhatIf
        Write-Verbose "$($user.SamAccountName) has been deleted" -Verbose
        $deletedusers += @{
            Name = $user.Name
            Username = $user.SamAccountName
            Title = $user.Title
            Department = $user.department
            Mail = $user.mail
            Manager = $manager
            LastLogon = [datetime]::FromFileTime($user.lastlogon)
            Deleted = Yes
        }
    }catch{
        Write-Verbose "Error deleting $($user.SamAccountName)"
        "Error deleting $($user.SamAccountName), $($_.exception.message)" | Out-File $deleteduserserrorlog -Append
    }
}

# Export all users that have been deleted to CSV

$deletedusers | Export-Csv $deleteduserscsv -NoTypeInformation -Append

# Variable to store HTML Content

$html = $null

# If $error exists append HTML Variable with the number of errors that occured

if($error){
    $html += "<center><p>$($error.count) errors - Please see error files</p></center>"
    }

# Import CSV that contains all accounts that were disabled - add content to $HTML variable.

if(Import-Csv $csv){
    $html +=  [string](ConvertTo-Html -Head "<style>$css</style>" -Body "<h2>Accounts Disabled $(Get-Date -Format d)</h2>")
    $html +=  [string](Import-Csv $csv | select Name,UserName,Title,Department,Mail,Manager,LastLogon,PasswordLastSet,FutureDeletionDate | ConvertTo-Html)
}

# Import CSV that contains all accounts that were deleted - add content to $HTML variable.

if(import-csv $deleteduserscsv){
    $html +=  [string](ConvertTo-Html -Head "<style>$css</style>" -Body "<h2>Accounts deleted$(Get-Date -Format d)</h2>")
    $html +=  [string](Import-Csv $deleteduserscsv | select Name,UserName,Title,Department,Mail,Manager,LastLogon,PasswordLastSet,Deleted | ConvertTo-Html)
}

# Send SMTP message with the $html variable as the content.

if($html){
    Send-MailMessage @mailparams -Body $html
}