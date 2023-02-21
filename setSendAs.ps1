<# 
  Connect to the specified Exchange Online environment and set the SendAs permission for the specified user or Office 365 group. Will install the EXO management module if it does not exist
#>

Function Connect-EXO {
    <#
      .SYNOPSIS
          Connects to EXO when no connection exists. Checks for EXO v3 module
    #>
    param(
      [Parameter(
        Mandatory = $true
      )]
      [string]$adminUPN
    )
    
    process {
      # Check if EXO is installed and connect if no connection exists
      if ($null -eq (Get-Module -ListAvailable -Name ExchangeOnlineManagement))
      {
        Write-Host "Exchange Online PowerShell v3 module is requied, do you want to install it?" -ForegroundColor Yellow
        
        $install = Read-Host Do you want to install module? [Y] Yes [N] No 
        if($install -match "[yY]") { 
          Write-Host "Installing Exchange Online PowerShell v3 module" -ForegroundColor Cyan
          Install-Module ExchangeOnlineManagement -Repository PSGallery -AllowClobber -Force
        }else{
          Write-Error "Please install EXO v3 module."
        }
      }
      $ModuleInstalled = Get-Module -ListAvailable -Name ExchangeOnlineManagement
      if ($null -ne $ModuleInstalled) {
        # Check which version of Exchange Online is installed
        if ($ModuleInstalled.version -like "3.*" ) {
          # Check if there is a active EXO sessions
          if ((Get-ConnectionInformation).tokenStatus -ne 'Active') {
            Write-Host 'Connecting to Exchange Online' -ForegroundColor Cyan
            Connect-ExchangeOnline -UserPrincipalName $adminUPN -ShowBanner:$false
          }
        }else{
          # Check if there is a active EXO sessions
          $psSessions = Get-PSSession | Select-Object -Property State, Name
          If (((@($psSessions) -like '@{State=Opened; Name=ExchangeOnlineInternalSession*').Count -gt 0) -ne $true) {
            Write-Host 'Connecting to Exchange Online' -ForegroundColor Cyan
            Connect-ExchangeOnline -UserPrincipalName $adminUPN
          }
        }
      }else{
        Write-Host "Please install EXO v3 module." -ForegroundColor Red
        exit 1
      }
    }
  }

$AuthUserName = Read-Host "Admin email to sign in as"
$SendAsEmail = Read-Host "User email to be given SendAs Trustee permission"
$UserOrGroup = Read-Host "Set SendAs for Office 365 group or single user? [G] Group [U] User"
if ($UserOrGroup -match "[gG]") {
  $Group = Read-Host "Office 365 group to assign SendAs to"
} elseif ($UserOrGroup -match "[uU]") {
  $UserEmail = Read-Host "Email of user to add SendAs?"
} else {
  Write-Host "Please specify either G or U to assign SendAs permissions." -ForegroundColor Red
  exit 1
}
Connect-EXO $AuthUserName
$SendAsUser = Get-Recipient $SendAsEmail
$CurrentSendAsUsers = Get-RecipientPermission -Trustee $SendAsUser.name
if ($UserOrGroup -match "[gG]") {
  try {
  $Users = Get-UnifiedGroupLinks -id $Group -LinkType members -ErrorAction  stop
  } catch {
    Write-Host "Unable to process group $Group. Please verify the group exists" -ForegroundColor Red
    Write-Host 'Disconnecting from Exchange Online' -ForegroundColor Cyan
    Disconnect-ExchangeOnline -Confirm:$false
    exit 1
  }
  ForEach($User in $Users) {
    # Check to see if current user is the trustee
    if ($User.name -ne $SendAsUser.name) {
      # Check to see if current user already is assigned to the trustee
      if ($null -eq (Compare-Object $CurrentSendAsUsers.identity $User.name -IncludeEqual -ExcludeDifferent).inputObject) {
        Write-Host "Adding permission for $SendAsUser to SendAs $User" -ForegroundColor Cyan
        Add-RecipientPermission $User.name -AccessRights SendAs -Trustee $SendAsUser -Confirm:$false 
      }
    } 
  }
} elseif ($UserOrGroup -match "[uU]") {
  try {
  $User = Get-Recipient $UserEmail -ErrorAction stop
  } catch {
    Write-Host "Unable to process user $User. Please verify the user exists" -ForegroundColor Red
    Write-Host 'Disconnecting from Exchange Online' -ForegroundColor Cyan
    Disconnect-ExchangeOnline -Confirm:$false
    exit 1
  }
  if ($User.name -eq $SendAsUser.name) {
    Write-Host "Trustee user and target user $User are the same" -ForegroundColor Yellow
    Write-Host 'Disconnecting from Exchange Online' -ForegroundColor Cyan
    Disconnect-ExchangeOnline -Confirm:$false
    exit
  }
  # Check to see if current user already is assigned to the trustee
  if ($null -eq (Compare-Object $CurrentSendAsUsers.identity $User.name -IncludeEqual -ExcludeDifferent).inputObject) {
    Add-RecipientPermission $User.name -AccessRights SendAs -Trustee $SendAsUser -Confirm:$false
  } else {
    Write-Host "$SendAsUser already has SendAs permissions for $User" -ForegroundColor Cyan
  }
}
Write-Host 'Disconnecting from Exchange Online' -ForegroundColor Cyan
Disconnect-ExchangeOnline -Confirm:$false