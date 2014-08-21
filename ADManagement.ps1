# -------------------------------------------------
# MS AD Users Managment
# Copyright (c) 2014 Alexey Baikov <sysboss@mail.ru>
#
# Batch domain users managment
# -------------------------------------------------
# Version: 0.1

# Parameters
Param(
	[switch]$usage,
	[switch]$help,
	[string]$ou
)

# Errors Handling
# $ErrorActionPreference= 'silentlycontinue'

# Window style
$win = (Get-Host).UI.RawUI
$win.ForegroundColor = "white"
$win.WindowTitle     = "AD Users Managment"

if($help -or $usage -or $h){
    '
usage: '+$MyInvocation.InvocationName+' [options] FROM
  -help|h          Help (this info)
  -verbos          Verbose mode
  -ou              OU Distinguished Name
    '
    exit 1
}

if(!$ou){
	Write-Host "OU Distinguished Name is required"
	exit 2
}

Write-Host $ou

# import the AD module
if (-not (Get-Module ActiveDirectory)){
    Import-Module ActiveDirectory -ErrorAction Stop            
}

$ADobj = Get-ADUser -Filter * -SearchBase $ou

foreach($user in $ADobj){
    $username = $user.Name
	$password = $username.Substring(0,1).ToUpper() + "123456!"
	Write-Host "Resetting password for:" $username

	Try {
		$user | Set-ADAccountPassword -Reset -NewPassword (ConvertTo-SecureString -AsPlainText $password -Force)
	} Catch {
		Write-Host "Password reset failed"
	}

	Write-Host "  password: " $password

	Try {
		$user | Set-ADUser -ChangePasswordAtLogon $True
	} Catch {
		$ErrorMessage = $_.Exception.Message

        if ($ErrorMessage.Contains('PasswordNeverExpires')){
			Write-Host " - cannot require ChangePasswordAtLogon. PasswordNeverExpires is set to True."
		}else{
			Write-Host " - cannot require ChangePasswordAtLogon."
		}
	}

	Write-Host ""
}
