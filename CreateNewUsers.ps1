###########################################################
#Variables created for file locations
$path = Split-Path -parent "D:\Scripts\NewUsers\*.*"
$csvfile = $path + "\UserList.csv"
$logfile = $path + "\logfile.txt"
$i        = 0
$date     = Get-Date

#Define AD Server
$ADServer = 'DC1.dcCMPY.COM'

#Prompt for Admin account credential
$GetAdminact = Get-Credential -Message "Enter in your Domain Admin account"

#Get Admin account credential for Exchange Online server
#$UserCredential = Get-Credential -Message "Enter in Admin Credentials for Azure AD Connect - Full email address required"
#$UserCredentialMigration = Get-Credential -Message "Enter SVC_Office365_Exchan account details for migration"

<# #The Azure AD Connect Server Name
$AADConnectServer = 'DC1H-AD02.RDFLX.COM' #>

# Import AD and Exchange modules
$ExchangeServer = "EXCH1.dcCMPY.COM"
$ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$ExchangeServer/PowerShell -Authentication Kerberos 
Import-PSSession $ExchangeSession -AllowClobber -DisableNameChecking | out-null

#Import Active Directory Module
Import-Module ActiveDirectory

#location variable for OU
$location = "OU=CMPYUsers,OU=UserAccounts,OU=CMPY,DC=COM"

#Function set to generate random password which meets complexity requirements
function Generate-RandomPassword {
  $Upper = [char[]]([int][char]'A'..[int][char]'Z') | Get-Random -Count 2
  $Lower = [char[]]([int][char]'a'..[int][char]'z') | Get-Random -Count 2
  $Special = "!@#$%^&*-./,<>" -split "" | Where-Object {$_} | Get-Random -Count 2
  $Number = 0..9 | Get-Random -Count 2
  ($Upper, $Lower, $Special, $Number | Get-Random -count 9999) -join ''
}

#Function to create ADUser 
Function Create-ADUsers {
"" | Out-File $logfile -append
"AD user creation logs for( " + $date + "): " | Out-File $logfile -append
"--------------------------------------------" | Out-File $logfile -append

#CSV is imported with relevant details
Import-Csv -Path $csvfile  | ForEach-Object { 
$GivenName = $_.'First Name'
$Surname = $_.'Last Name'
$Title = $_.Title
$Department = $_.Department
$Office = $_.Office

#Define domain
$Domain = "@CMPY.com"
#Define samAccountName
$sam = $GivenName.Substring(0, 1).ToLower() + $Surname.Replace(' ', '').Replace('-', '').ToLower()
#Define UPN
$UPN = $GivenName.ToLower() + "." + $Surname.Replace(' ', '').Replace('-', '').ToLower() + $Domain
#Define Display Name
$DisplayName = $GivenName + " " + $Surname
#Define email
$Mail = $GivenName + "." + $Surname.Replace(' ', '').Replace('-', '') + $Domain
#Define company
$Company = "CMPY Group"
#Define Description
$Description = $Title
#Sets random password, using the Generate-RandomPassword function created earlier
$Password = Generate-RandomPassword
#webpage
$webpage = "www.CMPY.com"

#Users are now created based off CSV file input. Try statement runs validation against sAMAccountName
Try   { $nameinAD = Get-ADUser -server $ADServer -Credential $GetAdminact -LDAPFilter "(sAMAccountName=$sam)" }
    Catch { }
    If(!$nameinAD)
    {
      $i++

#If "-enabled $TRUE" is not set, the account will be disabled by default
$setpassword = ConvertTo-SecureString -AsPlainText $password -force
      New-ADUser $sam -server $ADServer -Credential $GetAdminact `
      -GivenName $GivenName `
	  -ChangePasswordAtLogon $TRUE `
      -Surname $Surname `
	  -DisplayName $DisplayName `
	  -Office $Office `
      -Description $Description `
	  -EmailAddress $Mail `
	  -UserPrincipalName $UPN `
      -Company $Company `
	  -Department $Department `
	  -enabled $TRUE `
      -Title $Title `
	  -AccountPassword $setpassword `
      -HomePage $webpage `


#Define DN to use in the  Move-ADObject command below
 $dn = (Get-ADUser -server $ADServer -Credential $GetAdminact -Identity $sam).DistinguishedName
 
#The following moves the User using the $location variable set earlier from the default location in AD
 Move-ADObject -server $ADServer -Credential $GetAdminact -Identity $dn -TargetPath $location 
 
#The following will rename the using using the $DisplayName variable set earlierRename-ADObject
#Rename-ADObject only accepts DistinguishedNames
#Without the below, the users Display Name will be their sAMAccountName
 $newdn = (Get-ADUser -server $ADServer -Credential $GetAdminact -Identity $sam).DistinguishedName
 Rename-ADObject -server $ADServer -Credential $GetAdminact -Identity $newdn -NewName $DisplayName
 
 #Update log file with users created successfully
 $DisplayName + " Created successfully" | Out-File $logfile -append

 #User created successfully powershell output
 "$($User.'Display Name') Created successfully" | Out-File $logfile -append
 Write-Host "Creating AD account for $UPN" -ForegroundColor Yellow  
 Start-Sleep -s 5
 Write-Host "Success" -ForegroundColor Green  
 Start-Sleep -s 3
 Write-Host "Please wait a minute for the account to sync" -ForegroundColor Green
 Start-Sleep -s 15 
 
}
#Update log file with users not created powershell output  
Else
    { 
      $DisplayName + " Not Created - User Already Exists" | Out-File $logfile -append
	Write-Host "Creating AD User Object for $UPN | Failure, user already exists" -ForegroundColor Red  	  
    }
	
 	
#Create on premise Exchange mailbox
Try {
    Enable-Mailbox $UPN `
        -Database "MB1" `
        -ErrorAction Stop

    Write-Host "Creating Exchange Mailbox for $UPN" -ForegroundColor Yellow
	Start-Sleep -s 5
	Write-Host "Success" -ForegroundColor Green 	
	Start-Sleep -s 3
#Adding User to 'Successful Users' array, used to migrate to o365 later
    [Array]$SuccessfulUsers += $UPN	
}
catch {
    "Mailbox for " + $DisplayName + "Not Enabled" | Out-File $logfile -append
    Write-Host "Creating Exchange Mailbox for $UPN | Failed" -ForegroundColor Red
    Return
} 
  }
    }
# Run the create user function script 
Create-ADUsers




