<#
+---------------------------------------------------------------------------------------------+ 
| SERVER INVENTORY SCRIPT                                                                     | 
+---------------------------------------------------------------------------------------------+ 
|                                                                                             |
|   AUTHOR      : og-sokaar                                                                   |
|   DESCRIPTION : Script to inventory servers gather information on server resources - CPUs   |
|                 Memory, HDD, Domain, roles and features and email data formated into HTML   |
|                 Table                                                                       |
+---------------------------------------------------------------------------------------------+ 
#>

$csv = "<Insert UNC path for output>"
$mailpass = ConvertTo-SecureString "whatever" -AsPlainText -Force
$mailcreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList "NT AUTHORITY\ANONYMOUS LOGON", $mailpass
$mailparams = @{
    from = "<insert email address>"
    to = "<insert email address>"
    smtpserver = "<insert smtp server>"
    port = "25"
    subject = "Server Inventory report - $($date)"
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

$creationdate = (Get-Date).AddDays(-6)

Import-Module ActiveDirectory

Function Convert-OutputForCSV {
    <#
        .SYNOPSIS
            Provides a way to expand collections in an object property prior
            to being sent to Export-Csv.

        .DESCRIPTION
            Provides a way to expand collections in an object property prior
            to being sent to Export-Csv. This helps to avoid the object type
            from being shown such as system.object[] in a spreadsheet.

        .PARAMETER InputObject
            The object that will be sent to Export-Csv

        .PARAMETER OutPropertyType
            This determines whether the property that has the collection will be
            shown in the CSV as a comma delimmited string or as a stacked string.

            Possible values:
            Stack
            Comma

            Default value is: Stack

        .NOTES
            Name: Convert-OutputForCSV
            Author: Boe Prox
            Created: 24 Jan 2014
            Version History:
                1.1 - 02 Feb 2014
                    -Removed OutputOrder parameter as it is no longer needed; inputobject order is now respected 
                    in the output object
                1.0 - 24 Jan 2014
                    -Initial Creation

        .EXAMPLE
            $Output = 'PSComputername','IPAddress','DNSServerSearchOrder'

            Get-WMIObject -Class Win32_NetworkAdapterConfiguration -Filter "IPEnabled='True'" |
            Select-Object $Output | Convert-OutputForCSV | 
            Export-Csv -NoTypeInformation -Path NIC.csv    
            
            Description
            -----------
            Using a predefined set of properties to display ($Output), data is collected from the 
            Win32_NetworkAdapterConfiguration class and then passed to the Convert-OutputForCSV
            funtion which expands any property with a collection so it can be read properly prior
            to being sent to Export-Csv. Properties that had a collection will be viewed as a stack
            in the spreadsheet.        
            
    #>
    #Requires -Version 3.0
    [cmdletbinding()]
    Param (
        [parameter(ValueFromPipeline)]
        [psobject]$InputObject,
        [parameter()]
        [ValidateSet('Stack','Comma')]
        [string]$OutputPropertyType = 'Stack'
    )
    Begin {
        $PSBoundParameters.GetEnumerator() | ForEach {
            Write-Verbose "$($_)"
        }
        $FirstRun = $True
    }
    Process {
        If ($FirstRun) {
            $OutputOrder = $InputObject.psobject.properties.name
            Write-Verbose "Output Order:`n $($OutputOrder -join ', ' )"
            $FirstRun = $False
            #Get properties to process
            $Properties = Get-Member -InputObject $InputObject -MemberType *Property
            #Get properties that hold a collection
            $Properties_Collection = @(($Properties | Where-Object {
                $_.Definition -match "Collection|\[\]"
            }).Name)
            #Get properties that do not hold a collection
            $Properties_NoCollection = @(($Properties | Where-Object {
                $_.Definition -notmatch "Collection|\[\]"
            }).Name)
            Write-Verbose "Properties Found that have collections:`n $(($Properties_Collection) -join ', ')"
            Write-Verbose "Properties Found that have no collections:`n $(($Properties_NoCollection) -join ', ')"
        }
 
        $InputObject | ForEach {
            $Line = $_
            $stringBuilder = New-Object Text.StringBuilder
            $Null = $stringBuilder.AppendLine("[pscustomobject] @{")

            $OutputOrder | ForEach {
                If ($OutputPropertyType -eq 'Stack') {
                    $Null = $stringBuilder.AppendLine("`"$($_)`" = `"$(($line.$($_) | Out-String).Trim())`"")
                } ElseIf ($OutputPropertyType -eq "Comma") {
                    $Null = $stringBuilder.AppendLine("`"$($_)`" = `"$($line.$($_) -join ', ')`"")                   
                }
            }
            $Null = $stringBuilder.AppendLine("}")
 
            Invoke-Expression $stringBuilder.ToString()
        }
    }
    End {}
}

$ServersToInventory = Get-ADComputer -Filter * -Properties * | where{$_.Created -ge $creationdate -and $_.DistinguishedName -like "<insert OU Name wildcard operators work>"}

$InventoryOut = @()



foreach($server in $ServersToInventory){
    $ProcessorInfo = Get-WmiObject win32_Processor -ComputerName $server.Name
    $CompInfo = Get-WmiObject win32_ComputerSystem -ComputerName $server.Name
    $OS = Get-WmiObject win32_OperatingSystem -ComputerName $server.Name
    $PhysicalMemory = Get-WmiObject CIM_PhysicalMemory -ComputerName $server.Name | Measure-Object -Property capacity -Sum | ForEach-Object {[math]::Round(($_.sum / 1GB), 2)}
    $ports = Get-NetTCPConnection -State Listen -CimSession $server.Name | Select-Object LocalPort, @{Name="Process";Expression={(Get-Process -Id $_.OwningProcess).ProcessName}} | sort localport
    $HDDInfo = Get-WmiObject -Class Win32_logicaldisk -ComputerName $server.name | Select-Object DeviceID, @{L="Capacity";E={"{0:N2}" -f ($_.Size/1GB)}}  
    foreach($CPU in $ProcessorInfo){
        $Inventory = New-Object PSObject
        Add-Member -InputObject $Inventory -MemberType NoteProperty -Name "ServerName" -Value $CPU.SystemName
        Add-Member -InputObject $Inventory -MemberType NoteProperty -Name "Processor" -Value $CPU.Name
        Add-Member -InputObject $Inventory -MemberType NoteProperty -Name "PhysicalCores" -Value $CPU.NumberOfCores
        Add-Member -InputObject $Inventory -MemberType NoteProperty -Name "LogicalCores" -Value $CPU.NumberOfLogicalProcessors
        Add-Member -InputObject $Inventory -MemberType NoteProperty -Name "TotalPhysicalMemory-GB" -Value $PhysicalMemory
        Add-Member -InputObject $Inventory -MemberType NoteProperty -Name "Domain" -Value $CompInfo.Domain
        Add-Member -InputObject $Inventory -MemberType NoteProperty -Name "IPv4Address" -Value $server.IPv4Address
        Add-Member -InputObject $Inventory -MemberType NoteProperty -Name "OrganisationalUnit" -Value $server.CanonicalName
        Add-Member -InputObject $Inventory -MemberType NoteProperty -Name "OpenPorts" -Value $ports
        Add-Member -InputObject $Inventory -MemberType NoteProperty -Name "HDDinfo" -Value $HDDInfo
        $Inventory
        $InventoryOut += $Inventory
    }
}

$InventoryOut | Convert-OutputForCSV | Export-Csv $csv -NoClobber -NoTypeInformation

$html = $null

if(Import-Csv $csv){
    $html +=  [string](ConvertTo-Html -Head "<style>$css</style>" -Body "<h2>Server Inventory $(Get-Date -Format d)</h2>")
    $html +=  [string](Import-Csv $csv | select * | ConvertTo-Html)
}


if($html){
    Send-MailMessage @mailparams -Body $html
}


