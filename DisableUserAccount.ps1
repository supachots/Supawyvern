#Disables specified users' AD and Office365 accounts.

add-PSSnapin  quest.activeroles.admanagement -ErrorAction SilentlyContinue -WarningAction SilentlyContinue

function Get-DNC
{
	Param (
   		$RDSE
   	)
   
   	$DomainDNC = $RDSE.defaultNamingContext
   	Return $DomainDNC
  
}
$NC = (Get-DNC([adsi]("LDAP://RootDSE")))

$DisabledAccountOU = "OU=Trash,"+$NC

function get-dn ($SAMName)
{
 	$root = [ADSI]''
 	$searcher = new-object System.DirectoryServices.DirectorySearcher($root)
	$searcher.filter = "(&(objectClass=user)(sAMAccountName= $SAMName))"
	$user = $searcher.findall()

	if ($user.count -gt 1)
      {     
            $count = 0
            foreach($i in $user)
            { 
			write-host $count ": " $i.path 
                  $count = $count + 1
            }

            $selection = Read-Host "Please select item: "
			return $user[$selection].path

      }
      else
      { 
	  	return $user[0].path
      }
}

function moveToDisabledOU($strDN)
{
	Get-ADUser -Identity $strDN | Move-ADObject -TargetPath $DisabledAccountOU
}

function unassignO365Licenses()
{
    Write-Host "******************************************"
	Write-Host "* Unassign Office 365 Licenses"
	Write-Host "* Removes all licenses from listed user"
	Write-Host "******************************************"
	Write-Host ""
    
    Write-host "Please enter your Office 365 account credentials"
    $Cred = Get-Credential
    Connect-MSOLService –Credential $Cred
    Get-MsolAccountSku
    $o365ID = Read-Host "Please Enter in the samAccountName of the user you wish to disable (eg supachots, yoshin, tomoyukis)"
    $o365ID = $o365ID -replace "\.","_" #Office 365 replaces dots with underscores
    $o365user = $o365ID+"@reallyenglish.com"
    $msolAcc = Get-MsolUser -UserPrincipalName $o365user -ErrorAction SilentlyContinue
    if(!$msolAcc -And (Get-MsolUser -UserPrincipalName ($o365ID+"@rejapan.onmicrosoft.com")))
    {
            Write-Host ("Please log into Office 365 and change " +$o365ID+"@rejapan.onmicrosoft.com"+" to " + $o365user) -ForegroundColor "Red"
            Write-Host ""
            $o365user = $o365ID+"@rejapan.onmicrosoft.com"
            $msolAcc = Get-MsolUser -UserPrincipalName $o365user
    }

    while(!$msolAcc)
    {
        $o365user = Read-Host "User not found. Please Enter in the samAccountName of the user you wish to unassign licenses (eg supachots, yoshin, tomoyukis)"
        $msolAcc = Get-MsolUser -UserPrincipalName $o365user
    }
    Get-MsolUser -UserPrincipalName $o365user | Set-MsolUserLicense -RemoveLicenses "rejapan:O365_BUSINESS" -ErrorAction SilentlyContinue
    Get-MsolUser -UserPrincipalName $o365user | Set-MsolUserLicense -RemoveLicenses "rejapan:STADARDPACK" -ErrorAction SilentlyContinue
    Get-MsolUser -UserPrincipalName $o365user | Set-MsolUserLicense -RemoveLicenses "rejapan:VISIOCLIENT" -ErrorAction SilentlyContinue
    Get-MsolUser -UserPrincipalName $o365user | Set-MsolUserLicense -RemoveLicenses "rejapan:OFFICESUBSCRIPTION" -ErrorAction SilentlyContinue

    Write-Host "The license has been removed from user." -ForegroundColor "Red"
    $Choice = Read-Host "Would you like to remove licenses from another account?"
	If ($Choice.ToLower() -eq "y"){	unassignO365Licenses }
    else{ exit }
}

function disableADAcc()
{

	CLS
	Write-Host "**********************************************************"
	Write-Host "* Disable Windows AD Accounts"
	Write-Host "* Disable an account and stamp the account in description"
	Write-Host "**********************************************************"
	Write-Host ""
	[console]::ForegroundColor = "yellow"
	[console]::BackgroundColor= "black"
	$Name = Read-Host "Please Enter in the samAccountName of the user you wish to disable (eg supachots, yoshin, tomoyukis)"
	[console]::ResetColor()
	$ADPath = Get-ADUser -Identity $Name
    while(!$ADPath)
    {
        $Name = Read-Host "User not found. Please Enter in the samAccountName of the user you wish to disable (eg supachots, yoshin, tomoyukis)"
        $ADPath = Get-ADUser -Identity $Name
    }
	$status = "disable"
	$path = get-dn $Name 
	"'" + $path + "'"  
	

	
	[console]::ForegroundColor = "cyan"
	[console]::BackgroundColor= "black"
	$Reason = Read-Host "Please enter an explanation/reason"
	[console]::ResetColor()
    $DateT = (get-date -f yyyy-MM-dd).tostring()+" "+(get-date -f HH:mm).tostring()+" UTC"+(get-date -UFormat %Z)
    
	if ($status -match "disable") 
		{
			# Disable the account
			$account=[ADSI]$path
			$account.psbase.invokeset("AccountDisabled", "True")
			$account.setinfo()
		}
    
    Set-ADUser -Identity $Name -Description "Account Disabled by $env:username on $DateT due to: $Reason"

    $trashit = Read-Host "Would you like to also move user object to trash OU?"
	If ($trashit.ToLower() -eq "y")
    {
        moveToDisabledOU $Name
        Write-Host ""
        Write-Host "The user has been disabled and moved." -ForegroundColor "Red"
        Write-Host ""
    }
    else
    {
        Write-Host ""
        Write-Host "The user has been disabled." -ForegroundColor "Red"
        Write-Host ""
    }
	
	$Choice = Read-Host "Would you like to disable another account?"
	If ($Choice.ToLower() -eq "y"){	disableADAcc }
    else{ exit }
}


Write-Host "**********************************************************"
Write-Host "* Disable Windows AD Accounts and"
Write-Host "* Unassign Office 365 Licenses"
Write-Host "**********************************************************"
Write-Host ""

$tasks = Read-Host "Enter '1' to disable AD Account, '2' to unassign Office 365 licenses, or '3' to perform both tasks"
If ($tasks.ToLower() -eq "1")
{ disableADAcc }
elseif ($tasks.ToLower() -eq "2")
{ unassignO365Licenses }
else
{
    disableADAcc
    unassignO365Licenses
}