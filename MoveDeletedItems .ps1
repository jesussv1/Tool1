<#
    DISCLAIMER: This application is a sample application. The sample is provided "as is" without 
    warranty of any kind. Microsoft further disclaims all implied warranties including without 
    limitation any implied warranties of merchantability or of fitness for a particular purpose. 
    The entire risk arising out of the use or performance of the samples remains with you. 
    In no event shall Microsoft or its suppliers be liable for any damages whatsoever 
    (including, without limitation, damages for loss of business profits, business interruption, 
    loss of business information, or other pecuniary loss arising out of the use of or inability 
    to use the samples, even if Microsoft has been advised of the possibility of such damages. 
    Because some states do not allow the exclusion or limitation of liability for consequential or 
    incidental damages, the above limitation may not apply to you.

    ************************************
    Created by: Jesus Salazar
    E-mail: jesussalazar@msn.com
    ************************************
    For a Mailbox in Office 365 Move all emails from Delete Folder to the archive Delete Folder

    ************************************
    Prerequisites
    ************************************
    1 - This script is only valid for cloud users is not working with synchronide users with Exchange on-Premise
    2.- We need a Global Administrator credentials
    3.- This solution only applied for a user with license of Exchange Online Plan 1 and Plan 2 it is not valid for Kiosco License

    ************************************

    Use of the script:

    This Script Will be enable the Archive for the user
    If the user has license E3 or E5 (Exchange Plan 2) Also enable the unlimited Archive
    
    This Script will be Move all emails from Delete folder more older of 1 day on the Primary Mailbox to to archive folder in the primary mailbox

    PS>.\MoveDeleteItems.ps1 -MailboxName test@o365genius.com 

    Move all emails from inbox/main mailbox to Inbox/archive folder in the archive online. 
    it will ask for username and password for Global Admin.
AGREGAR ESTAS LINEAS
$Politica=Get-RetentionPolicy "archivar 360 dias" -ErrorAction ignore
if ($politica) {Write-host "ya existe" } else {write-host "continuar con el proceso"  }
ya existe
$Politica=Get-RetentionPolicy "archivar 460 dias" -ErrorAction ignore
if ($politica) {Write-host "ya existe" } else {write-host "continuar con el proceso"  }
continuar con el proceso

#>

param (
  [Parameter(Position=0,Mandatory=$True,HelpMessage='Specifies the mailbox to be accessed')]
  [ValidateNotNullOrEmpty()]
  [string]$MailboxName
);

[string]$warning = 'Yellow'                      # Color for warning messages
[string]$myerror = 'Red'                           # Color for error messages
[string]$LogFile = '.\Log.txt'             # Path of the Log File


$icount = 0;
$i=0;

#if username or password are empty, ask for both

if ([string]::IsNullOrEmpty($username) -or [string]::IsNullOrEmpty($password)) {
  $UserCredential = Get-Credential -Message 'Admin Username and password'
  if ($UserCredential) {
    $username = $UserCredential.UserName
    $BSTR = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($UserCredential.Password)
    $Password = [Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
  }
  else { 
    Write-Error -Message 'Admin credential needed'
    return $False
  }
}


# Make sure the Import-Module command matches the Microsoft.Exchange.WebServices.dll location of EWS Managed API, chosen during the installation

Start-Transcript

#CONNECT TO EXCHANGE ONLINE

$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
Import-PSSession $Session -AllowClobber

#ENABLE ARCHIVE FOR THIS USER (	CLOUD USER )

Enable-Mailbox $MailboxName -Archive 

#ENABLE AUTOEXPANDARCHIVE FOR THIS USER ONLY APPLY FOR USER WITHN LICENSE E3, E5 O ARCHIVE EXCHANGE ONLINE ADD ON.

Enable-Mailbox $MailboxName -AutoExpandingArchive 

#CREATE A RETENTION TAG:

New-RetentionPolicyTag -Name "MoveDeleteditems" -Type RecoverableItems -AgeLimitForRetention 1 -RetentionAction MoveToArchive 

#CREATE THE RETENTION POLICY:

New-RetentionPolicy “Borrar Deleted Items” -RetentionPolicyTagLinks “MoveDeletedItems“

#APPLY THIS RETENTION POLICY FOR THIS USER

Set-Mailbox -identity $MailboxName -RetentionPolicy "Borrar Deleted Items"

#FORCE TO START THE ARCHIVING PROCESS

Start-ManagedFolderAssistant -Identity $MailboxName

Get-MailboxStatistics -Identity $MailboxName |FL DisplayName,TotalDeletedItemSize, TotalItemSize

Stop-Transcript