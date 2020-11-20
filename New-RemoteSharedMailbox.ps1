param ($Alias,$DisplayName)

#	Show countdown timer

Function Start-Countdown 
{   
    Param(
        [Int32]$Seconds = 600,
        [string]$Message = "Waiting for 10 minutes"
    )
    ForEach ($Count in (1..$Seconds))
    {   Write-Progress -Id 1 -Activity $Message -Status "Waiting for $Seconds seconds, $($Seconds - $Count) left" -PercentComplete (($Count / $Seconds) * 100)
        Start-Sleep -Seconds 1
    }
    Write-Progress -Id 1 -Activity $Message -Status "Completed" -PercentComplete 100 -Completed
}



#	Load Exchange Powershell module 
#add-pssnapin Microsoft.Exchange.Management.PowerShell.E2010
add-pssnapin Microsoft.Exchange.Management.PowerShell.SnapIn
#	Load Active Directory Powershell module 
import-module activedirectory

# Setup variables
$DomainController="servername.domain.lan"
$OU='domain.lan/Groups/Mail/Shared Mailboxes'
$OU2='domain.lan/Groups/Mail/Shared Mailbox Groups'
$UPNdomain = "@domain.nl"
$UPNRemoteDomain = "@domain.onmicrosoft.com"



if ($Alias)
{
    if ($Alias.Contains('@')) { $Alias = $Alias.Substring(0,$Alias.IndexOf('@')) }
    $AliasMailbox = Get-Mailbox $Alias -ErrorAction SilentlyContinue
    $AliasMailUser = Get-MailUser $Alias -ErrorAction SilentlyContinue
    if ($AliasMailbox -or $AliasMailUser)
    {
        Write-Host "The Alias specified already exists" -ForegroundColor red
        $Alias = $null
    }
}
while (!$Alias)
{
    $Alias = Read-Host -Prompt "Alias (max 20 caracters)"
    if ($Alias)
    {
        if ($Alias.Contains('@')) { $Alias = $Alias.Substring(0,$Alias.IndexOf('@')) }
        $AliasMailbox = Get-Mailbox $Alias -ErrorAction SilentlyContinue
        $AliasMailUser = Get-MailUser $Alias -ErrorAction SilentlyContinue
        if ($AliasMailbox -or $AliasMailUser)
        {
            Write-Host "The Alias specified already exists" -ForegroundColor red
            $Alias = $null
        }
    }
}

if ($DisplayName)
{
    $DisplayNameMailbox = Get-Mailbox $DisplayName -ErrorAction SilentlyContinue
    $DisplayNameMailUser = Get-MailUser $DisplayName -ErrorAction SilentlyContinue
    if ($DisplayNameMailbox -or $DisplayNameMailUser)
    {
        Write-Host "The Display Name specified already exists" -ForegroundColor red
        $DisplayName = $null
    }
}
while (!$DisplayName)
{
    $DisplayName = Read-Host -Prompt "Display Name (As many caracters as you like)"
    if ($DisplayName)
    {
        $DisplayNameMailbox = Get-Mailbox $DisplayName -ErrorAction SilentlyContinue
        $DisplayNameMailUser = Get-MailUser $DisplayName -ErrorAction SilentlyContinue
        if ($DisplayNameMailbox -or $DisplayNameMailUser)
        {
            Write-Host "The Display Name specified already exists" -ForegroundColor red
            $DisplayName = $null
        }
    }
}

# Setup more variables
$Alias=$Alias.ToLower()
$UPN=$Alias + $UPNDomain
$UPNRemoteRoutingAddress = $Alias + $UPNRemoteDomain

Sleep 10
# Create the SharedMailbox
Write-Host "Creating Shared Mailbox" -ForegroundColor green
New-RemoteMailbox -RemoteRoutingAddress "$UPNRemoteRoutingAddress" -Shared -UserPrincipalName "$UPN" -OnPremisesOrganizationalUnit $OU -Alias $alias -Name $alias -DisplayName $displayname -PrimarySmtpAddress $UPN -SamAccountName $alias -DomainController $domaincontroller
Write-Host "Created Shared Mailbox" -ForegroundColor green

Sleep 10
# Set the description for the SharedMailbox
Write-Host "Set Description" -ForegroundColor green
Set-ADUser $Alias -Description "Shared Mailbox t.b.v. $Displayname"
Write-Host "Description set" -ForegroundColor green

Sleep 10
# Create the distributiongroup for security use
Write-Host "Creating Office365 Distributiongroup" -ForegroundColor green
New-DistributionGroup  -DisplayName "SM.$alias" -Type Security -Alias "SM.$alias" -Name "SM.$alias" -Organizationalunit $OU2
Write-Host "Office365 Distributiongroup created" -ForegroundColor green

Sleep 10
# Hide distributiongroup
Write-Host "Set Office365 distributiongroup hidden" -ForegroundColor green
Set-DistributionGroup -Identity "SM.$alias" -HiddenFromAddressListsEnabled:$true
Write-Host "Office365 distributiongroup set to hidden" -ForegroundColor green

Sleep 30
# Sync AADConnect and wait for the account to show up online
Write-Host "Starting Adsynccycle now" -ForegroundColor red
Invoke-Command -ComputerName servername.domain.lan -Port 5986 -UseSSL -ScriptBlock { Start-ADSyncSyncCycle -PolicyType Delta }
Write-Host "Adsynccycle has run" -ForegroundColor green

Write-Host "Waiting for AzureAD sync" -ForegroundColor green
#Start-Countdown -Seconds 600 -Message "Waiting for 10 minutes"

$Time = 600
$i = 0
Do {
    $i++
    Write-Progress -Activity 'Waiting for 10 minutes' -Status 'Status' -PercentComplete (($i/$Time)*100) -SecondsRemaining ($Time-$i)
    Start-Sleep 1
} Until ($i -eq $Time)

# Set the PowerShell session to use the proxy
netsh winhttp set proxy proxy.domain.lan:8080
Write-Host "Proxy Set" -ForegroundColor green

# Connect to ExchangeOnline PowerShell
Connect-ExchangeOnline -ShowProgress $true
Write-Host "Connected to ExchangeOnline" -ForegroundColor green

# Disable Mailbox features
Write-Host "Disabeling OWA, POP, IMAP, ActiveSync" -ForegroundColor green
Set-CASMailbox -Identity $Alias -imapenabled $false -owaenabled $false -OWAforDevicesEnabled $false -popEnabled $false -ActiveSyncEnabled $false -PopUseProtocolDefaults $false -ImapUseProtocolDefaults $false
Write-Host "OWA, POP, IMAP, ActiveSync disabled" -ForegroundColor green

Sleep 10
# Add the distributiongroup to the sharedmailbox with Full Access
Write-Host "Setting Mailbox Full Access Permissions" -ForegroundColor green
Add-MailboxPermission –Identity: $Alias –AccessRights:FullAccess –user:"SM.$Alias"
Write-Host "Full Access Permissions set" -ForegroundColor green

Sleep 10

# Add the distributiongroup to the sharedmailbox with Send-as
Write-Host "Setting Mailbox Send-as Permissions" -ForegroundColor green
Add-ADPermission -Identity "$Alias" -user "SM.$Alias" -ExtendedRights 'Send-as' -DomainController $DomainController
Write-Host "Send-as Permissions set" -ForegroundColor green

Sleep 10

# Reset proxy to direct access
netsh winhttp reset proxy
Write-Host "Proxy Set to default" -ForegroundColor green
Write-Host "Script Finished" -ForegroundColor green
Write-Host "Close this window " -ForegroundColor Red
