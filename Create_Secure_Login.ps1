<#
 # Author: Daniel Yan
 # Creates a 32 byte AES encrypted user and password encrypted text files based on inputs
 # 
 # Date: 6/7/2019
 #>

#Function to create the necessary key files and user/pass txt files
Function Login
{
    	[cmdletBinding()]
    param (
        [parameter(valueFromPipeline)]
        [validateNotNullOrEmpty()]
        [string]$path
        )

    $username = Read-Host -Prompt "Please enter your username" -AsSecureString
    $password = Read-Host -Prompt "Please enter your password" -AsSecureString

    $userFilename = Read-Host -Prompt "Please enter your user filename"
    $passwordFilename = Read-Host -Prompt "Please enter your password filename"

    $passwordFile = $path + '\'+ $passwordFilename + '.txt'
    $userFile = $path + '\'+ $userFilename + '.txt'

    $userkeyFile = $path + '\userAES.key'
    $passkeyFile = $path + '\passAES.key'

    $userKey = New-Object Byte[] 32   # 16, 24, or 32 byte options for AES
    [Security.Cryptography.RNGCryptoServiceProvider]::Create().GetBytes($userKey)
    $userKey| out-file $userkeyFile

    $passKey = New-Object Byte[] 32   # 16, 24, or 32 byte options for AES
    [Security.Cryptography.RNGCryptoServiceProvider]::Create().GetBytes($passKey)
    $passKey | out-file $passkeyFile

    $userKey = Get-Content $userKeyFile
    $passKey = Get-Content $passKeyFile

    $username | ConvertFrom-SecureString -key $userKey | Out-File $userFile
    $password | ConvertFrom-SecureString -key $passKey | Out-File $passwordFile
}


$checkSelection = $true
$invocation = (Get-Variable MyInvocation).Value
$directoryPath = Split-Path $invocation.MyCommand.Path
$cmdlinePath = (Get-Item -Path ".\").FullName
            
while($checkSelection)
{
$selection = Read-Host -Prompt "Please select path you wish to store the file in
[1] Current path of the script
[2] Current path of the command line directory
Your Selection"
    
        if($selection.equals("1"))
        {
            $checkSelection = $false
            $directoryPath | Login

        }
        elseif($selection.equals("2"))
        {
            $checkSelection = $false
            $cmdlinePath | Login

        }
        else
        {
            Write-Host 'Please enter a valid selection'
            $checkSelection = $true        
        }
}
