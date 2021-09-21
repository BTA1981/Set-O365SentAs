<#
.SYNOPSIS
  Set SentAs rights from a CSV file.
.DESCRIPTION
  
.INPUTS
  CSV file with contents:
  DOMAIN\A.lastname;DOMAIN.local/DOMAIN/Accounts/User Accounts/OU/Supervisor
  DOMAIN\b.lastname2;DOMAIN.local/DOMAIN/Accounts/User Accounts/OU/Info
.OUTPUTS
  <Outputs if any, otherwise state None>
.NOTES
  Version:        1.0
  Author:         Bart Tacken
  Creation Date:  21-09-2021
  Purpose/Change: Initial script development
.PREREQUISITES
  Windows Management Framework 5.1
  Install-Module -Name MSOnline 
  Install-Module -Name ExchangeOnlineManagement (EXO V2 module)

.EXAMPLE
  <Example goes here. Repeat this attribute for more than one example>
  <Example explanation goes here>
#>

<#
param (
        [Parameter(Mandatory=$True)] # 
        [string]$Param1,

        [Parameter(Mandatory=$True)] # 
        [string]$Param2 # 

 ) # End Param
#>


#---------------------------------------------------------[Initialisations]--------------------------------------------------------
[string]$DateStr = (Get-Date).ToString("s").Replace(":","-") # +"_" # Easy sortable date string    
Start-Transcript ('c:\windows\temp\' + $DateStr  + '_Set-O365SentAs.log') -Force # Start logging

#Set Error Action to Silently Continue
$ErrorActionPreference = 'SilentlyContinue'
#If (!(Get-Module <module>) { Import-Module <module>}
#----------------------------------------------------------[Declarations]----------------------------------------------------------

#Any Global Declarations go here
$CredPath = "C:\Beheer\Script\key\O365cred.xml"
$KeyFilePath = "C:\Beheer\Script\key\O365key.key"
$CSVpath = 'C:\Beheer\Script\Set-O365SentAs\Users - SendAs.csv'
$CSVcontent = Import-Csv -Path $CSVpath -Delimiter ";" -Header "User", "Mailbox"
#-----------------------------------------------------------[Functions]------------------------------------------------------------
Function Connect-EXOnline {
    param($Credentials)
    $URL = "https://ps.outlook.com/powershell"     
    $EXOSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $URL -Credential $Credentials -Authentication Basic -AllowRedirection -Name "Exchange Online"
        Import-PSSession $EXOSession -AllowClobber
}
#-----------------------------------------------------------[Execution]------------------------------------------------------------
$Key = Get-Content $KeyFilePath
$credXML = Import-Clixml $CredPath #Import encrypted credential file into XML format
$secureStringPWD = ConvertTo-SecureString -String $credXML.Password -Key $key
$Credentials = New-Object System.Management.Automation.PsCredential($credXML.UserName, $secureStringPWD) # Create PScredential Object

Connect-ExchangeOnline -Credential $Credentials

ForEach ($Line in $CSVcontent) {

    $User = $Line.User
    $User = $User.split("\")[-1]
    $SharedMailbox = $Line.Mailbox
    $SharedMailbox = $SharedMailbox.split("/")[-1]

    Write-Host "Setting Sent As rights for user [$User] on MailBox [$SharedMailbox]"
    Add-RecipientPermission -Identity "$SharedMailbox" -Trustee "$User" -AccessRights SendAs -Confirm:$False #-whatif
    
    <#
    Try {
        #Add-RecipientPermission -Identity "facturen" -Trustee "$User" -AccessRights SendAs -whatif
    }
    Catch {
        Write-Host "Something went wrong with settings rights for user [$User] on MailBox [$SharedMailbox]
        Add-Content C:\temp\test.txt "`nThis is a new line""
    } 
    #>
}
Stop-Transcript
