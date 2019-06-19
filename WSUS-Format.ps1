<#
WSUS-Format.ps1
Original Creator: Adam Sailer
Modified By: Daniel Yan
Last updated: 6/7/2019

Prints out an OGV as formatted to the WSUS spreadsheet. Sorts it by Last Status received.
Checks WSUS, Altiris database, GLPI database, DHCP leases, and Active Directory database.

Uses a 32 Byte AES Encrypted file for login credentials, current filenames that are assumed in the local directory are listed
#>

$invocation = (Get-Variable MyInvocation).Value
$directoryPath = Split-Path $invocation.MyCommand.Path
$userFilename = REDACTED
$passFilename = REDACTED

#Global vars
$format = 'yyyy.MM.dd'
$server = Get-WsusServer -Name REDACTED -PortNumber 8530
$dhcpServer = "REDACTED"
$dcServer = 'REDACTED'
$glpiServer = 'REDACTED'  
$selection

#Would like to make this use prompt into a function
$checkSelection = $true
$checkDate = $true
$current = Get-Date
while($checkSelection)
{
#Could be expanded out for additional selections (e.g. Model Type, OS type ...)

$selection = Read-Host -Prompt "Please select the range of computers you wish to view
[1] All Computers
[2] A date-range of Computers
Your Selection"
    
        if($selection.equals("1"))
        {
            $checkSelection = $false
        }
        elseif($selection.equals("2"))
        {
            $checkSelection = $false
        }
        else
        {
            Write-Host 'Please enter a valid selection'
            $checkSelection = $true        
        }
    }

    if($selection.equals("2"))
    {
        while($checkDate)
        {
            $startingDate = Read-Host -Prompt 'Input the starting date range you wish search by (MM/DD/YYYY)'
            $endingDate = Read-Host -Prompt 'Input the ending date range you wish to search by (MM/DD/YYYY)'
            if($startingDate.equals("now") -Or $startingDate.equals("today"))
            {
                $startingDate = $current
            }
            if($endingDate.equals("now"))
            {
                $endingDate = $current
            }
            elseif($endingDate.equals("today"))
            {
                $endingDate = [DateTime]::Today
            }
            if([DateTime] $endingDate -gt $current) 
            {
                Write-Host 'Cannot input an ending date after today'
                $checkDate = $true
            }
            elseif ([DateTime] $startingDate -gt [DateTime] $endingDate)
            {
                Write-Host 'Cannot input a starting date after end date'
                $checkDate = $true    
            }
            else 
            {
                $checkDate = $false
            }
        }
    }



import-module ActiveDirectory

# MS-SQL Querying for Altiris(Ghost)
$query_altiris = @"
     REDACTED QUERIES
"@

# MySQL Querying for GLPI
$query_glpi = @"
    REDACTED QUERIES
"@

#Query Altiris connection necessities
Function Query_Altiris
{
	[cmdletBinding()]
    param (
        [parameter(valueFromPipeline)]
        [validateNotNullOrEmpty()]
        [string]$query
        )

    process {

        write-host "`n`n@@ Running Query_Altiris" -fore magenta

        $server = 'REDACTED'  
        $database = 'REDACTED'
        $sqlQuery = $query
        $sqlConnectionString = "Server = $server; Database = $database; Integrated Security = True"
        $sqlConnection = new-object System.Data.SqlClient.SqlConnection($sqlConnectionString)
        $sqlCmd = new-object System.Data.SqlClient.SqlCommand($sqlQuery,$sqlConnection)
        $sqlAdapter = new-object System.Data.SqlClient.SqlDataAdapter($sqlCmd)
        $dataSet = new-object System.Data.DataSet
        $sqlAdapter.Fill($dataSet) | out-null <# needed! #>
        $sqlConnection.Close()

        $out = $dataSet.Tables[0];  ## write-host $out.Rows.Count
        
        return $out
    }
}

#Query GLPI connection necessities
Function Query_GLPI
{
	[cmdletBinding()]
    param (
        [parameter(valueFromPipeline)]
        [validateNotNullOrEmpty()]
        [string]$query
        )

    process {

        write-host "`n`n@@ Running Query_GLPI" -fore magenta

        [void][System.Reflection.Assembly]::LoadWithPartialName('mysql.data')
        
        $userkeyFile = $directoryPath + '\REDACTED.key'
        $passkeyFile = $directoryPath + '\REDACTED.key'

        $userKey = Get-Content $userKeyFile
        $passKey = Get-Content $passKeyFile

        $username = ConvertTo-SecureString -key $userKey -string (Get-Content "$($directoryPath)\$($userFilename)")  
        $password = ConvertTo-SecureString -key $passKey -string (Get-Content "$($directoryPath)\$($passFilename)")  
        
        $userBSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($username)
        $glpiUsername = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($userBSTR)
        $passBSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($password)
        $glpiPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($passBSTR)


        $database = 'REDACTED'
        $sqlQuery = $query
        $sqlConnectionString = "server=$glpiServer; uid=$glpiUsername; pwd=$glpiPassword; database=$database;"
        $sqlConnection = new-object MySql.Data.MySqlClient.MySqlConnection($sqlConnectionString)
        $sqlCmd = new-object MySql.Data.MySqlClient.MySqlCommand($sqlQuery, $sqlConnection)        
        $sqlAdapter = new-object MySql.Data.MySqlClient.MySqlDataAdapter($sqlCmd)
        $dataSet = new-object System.Data.DataSet
        $sqlAdapter.Fill($dataSet) | out-null
        $sqlConnection.Close()

        $out = $dataSet.Tables[0];
        
        return $out
    }
}

# Query Active directory connection necessities
Function Query_AD
{
    [cmdletBinding()]
    param (
        [parameter(valueFromPipeline)]
        [validateNotNullOrEmpty()]
        [string]$query
        )
    process {
    write-host "`n`n@@ Running Query_AD" -fore magenta
    
    #Older attributes left if needed
    # Ordered map for attributes
    $map = [ordered]@{
        CanonicalName = 1;
        Name = 1;
        OperatingSystem = 1;        
        SerialNumber = 1;
        }

    $properties = $map.Keys | ? { $map.$_ -gt 0 }


    get-adComputer -ldapFilter "(name=*)" -searchBase "dc=uwb,dc=edu"-properties $properties -server $dcServer | 
    ? { $_.CanonicalName -inotmatch 'server' } | select $properties | sort-object -property Name # |
    }
}


Function BinarySearch
{
    param($array, $inp, $property)

    [int]$min = 0
    [int]$max = $array.Count - 1
    
    $searched = $inp   

    while ($min -le $max)
    {        
        [int]$mid = $min + (($max - $min) / 2)        
        
        $inspected = $null
        $inspected = $array[$mid].$property
                
        try {
            if ($inspected -lt $searched) { $min = $mid + 1 }
            elseif ($inspected -gt $searched) { $max = $mid - 1 }
            else { return $($array[$mid]) }

        } catch { return }
    }
    return
}

# Main running process
Function Process
{
    write-host "`n`n@@ Running Process" -fore magenta
    
    # Cast to proper object type. All current UWB subnets
    # REDACTED INFO ON IP SUBNETS
        )

    # Empty array for all the leases
    $leases_all = @()

    $subnets | % { 
        Get-DhcpServerv4Lease -ComputerName REDACTED -ScopeId $_ | % { $leases_all += $_ }
    }

    $leases_recent = @()

    #Add all the most recent leases to $leases_recent
    $leases_all | Group-Object -property HostName | Sort-Object -Property Name | % {
        
        $leases_recent += $($_.Group | sort-object -Property LeaseExpiryTime | Select-Object -Last 1)
    }

    $targets = @($server.GetComputerTargets())
        
    if($selection.equals("1"))
    {
        #Sort currently by OptiPlex, LastReportedStatusTime, Role of workstation, and Datetime
        $item = ($targets | ? { $_.Model -imatch 'optiplex' } | ? { $_.ComputerRole -ieq 'workstation' } | 
        Sort-Object -Property LastReportedStatusTime)
    }
    
    elseif($selection.equals("2"))
    {
        #Sort currently by OptiPlex, LastReportedStatusTime, Role of workstation, and Datetime. Time is converted to GMT-7
        $item = ($targets | ? { $_.Model -imatch 'optiplex' } | ? { $_.ComputerRole -ieq 'workstation' } | 
        ? {(Get-Date $_.LastReportedStatusTime.AddHours(-7)) -ge [Datetime]$startingDate} | 
        ? {(Get-Date $_.LastReportedStatusTime.AddHours(-7)) -le [Datetime]$endingDate} |
        Sort-Object -Property LastReportedStatusTime) 
    }
    #This loop is left unclosed on purpose
    $item | % { $target = $_ 

        #Truncate uwb.edu
        $name = $null
        $name = $_.FullDomainName.ToUpper() -ireplace '.uwb.edu',''
        $serial = $null
        $serial = ($name | select-string -pattern '(\w|\d){4,}$').Matches.Value

        #Match based on name, serial, and serial number
        $match_ad = $null
        $match_ad = BinarySearch $array_ad $name 'name'
        if (!$match_ad) { $match_ad = BinarySearch $array_ad_serial $serial 'serialnumber' }
        if (!$match_ad -and $serial) { $match_ad = $array_ad_serial | ? { $_.SerialNumber -imatch $serial } }

        #Match based on name, serial, and serial number
        $match_altiris = $null
        $match_altiris = BinarySearch $array_altiris $name 'name'
        if (!$match_altiris) { $match_altiris = BinarySearch $array_altiris_serial $serial 'serial' }        
        if (!$match_altiris -and $match_ad) { $match_altiris = BinarySearch $array_altiris_serial $match_ad.SerialNumber 'serial' }
        if (!$match_altiris -and $serial) { $match_altiris = $array_altiris_serial | ? { $_.Serial -imatch $serial } }

        #Match based on name, serial, and serial number
        $match_glpi = $null
        $match_glpi = BinarySearch $array_glpi $name 'name'       
        if (!$match_glpi) { $match_glpi = BinarySearch $array_glpi_serial $serial 'serial' }        
        if (!$match_glpi -and $match_ad) { $match_glpi = BinarySearch $array_glpi_serial $match_ad.SerialNumber 'serial' }
        if (!$match_glpi -and $serial) { $match_glpi = $array_glpi_serial | ? { $_.Serial -imatch $serial } }

        #Match lease based on hostname
        $match_lease = $null
        $match_lease = BinarySearch $leases_recent $target.FullDomainName 'hostname'


            <# This section here is simplifying data to format to current WSUS spreadsheet#>
              #define active states for icons
              $activeArray = @(2,25,57,104)
    
              #Set the message for an active computer
              $status  = try { $match_altiris.Icon} catch { '' }; 
              if ($activeArray.Contains([int]$status)){ $activeStatus = "Active"}
              elseif ($status -eq $null){ $activeStatus = "Missing"} 
              else {$activeStatus =  "Inactive"}
              #All possible address states
              if($match_lease.AddressState -ceq 'Active'){$leaseStatus = "Active"}
              elseif($match_lease.AddressState -ceq 'ActiveReservation'){$leaseStatus = "Reservation-Active"}
              elseif($match_lease.AddressState -ceq 'InactiveReservation'){$leaseStatus = "Reservation-Inactive"}
              elseif($match_lease.AddressState -ceq 'Declined'){$leaseStatus = "Declined"}
              elseif($match_lease.AddressState -eq $null){$leaseStatus = "Missing"}
              else {$leaseStatus = "Inactive"}
              if($match_glpi.Status){$glpiStatus = $match_glpi.Status}
              else {$glpiStatus = "None"}
              if([System.String]$match_glpi.User -ne ""){$glpiUser = "GLPI User: " + $match_glpi.User}
              else {$glpiStatus = ""}
              #Old AD matching
              #$match_ad.OperatingSystem 
              #Match OS to Altiris
              if($match_altiris.OS) {$osStatus = $match_altiris.OS}
              else {$osStatus = "None"}
            <# This section here is simplifying data to format to current WSUS spreadsheet#>  
              
        $map = [ordered]@{

             'Name' = $name;
             #Convert time to GMT-7 as WSUS uses GMT or UTC 0:00
             'Last Status' = try{Get-Date $_.LastReportedStatusTime.AddHours(-7)} catch{};
             'Last Contact' = try {Get-Date $_.LastSyncTime.AddHours(-7)} catch {};
             'Contact IP' = try{$_.IPAddress} catch{};
             'Operating System' = $osStatus;
             'Computer Type' = try{$_.Model} catch{};
             'Notes' = try{$glpiUser} catch{};
             'Deploy Console' = $activeStatus;
             'Owner/Logged On' = try { $match_altiris.User } catch {};
             #Truncate uwb.edu/ from Canonical name
             'AD Account OU' = try { $match_ad.CanonicalName.substring(8) } catch {};
             'DHCP' = $leaseStatus;
             'GLPI' = try { $glpiStatus} catch {};
        }
        
        new-object psObject -property $map | write-output
    #Output the file to a gridview based on current date
    } | ogv -title "$($myInvocation.MyCommand.Name) : $(get-date)"
}
# Clear current terminal buffer
Clear
#Run queries
$array_ad = Query_AD
$array_ad_serial = $array_ad | sort-object SerialNumber

$array_altiris = $query_altiris | Query_Altiris
$array_glpi = $query_glpi | Query_GLPI

$array_altiris_serial = $array_altiris | sort-object Serial
$array_glpi_serial = $array_glpi | sort-object Serial

#Run actual function
Process
