#Disables specified users' AD and Office365 accounts

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

function assignO365Licenses()
{
    Write-Host "******************************************"
	Write-Host "* Assign Office 365 Licenses"
	Write-Host "* List and assign select licenses to user in Office 365"
	Write-Host "******************************************"
	Write-Host ""
    
    Write-host "Please enter your Office 365 admin credentials"
    $Cred = Get-Credential
    Connect-MSOLService –Credential $Cred
    Get-MsolAccountSku
    $o365ID = Read-Host "Please Enter in the samAccountName of the user you wish to assign licenses to (eg supachots, yoshin, tomoyukis)"
    $o365ID = $o365ID -replace "\.","_" #Office 365 replaces dots with underscore

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
        $o365ID = Read-Host "User not found. Please Enter in the samAccountName of the user you wish to assign licenses to (eg supachots, yoshin, tomoyukis)"
        $o365user = $o365ID+"@reallyenglish.com"
        $msolAcc = Get-MsolUser -UserPrincipalName $o365user
    }

    $accSku = Get-MsolAccountSku

    [int]$x = 0
    while($accSku.Length -gt $x)
    {
        Write-Host (-join (-join($x, ". "), $accSku.AccountSkuId[([convert]::ToInt32($x, 10))]))
        $x++
    }
    $licSelect = Read-Host "Pick which license you would like to assign to user by entering corresponding number"

    Get-MsolUser -UserPrincipalName $o365user | Set-MsolUserLicense -AddLicenses $accSku.AccountSkuId[([convert]::ToInt32($licSelect, 10))]

    Write-Host "The license has been added to user." -ForegroundColor "Red"
    $Choice = Read-Host "Would you like to add licenses to another account?"
	If ($Choice.ToLower() -eq "y"){	assignO365Licenses }
    else{ exit }
}
assignO365Licenses