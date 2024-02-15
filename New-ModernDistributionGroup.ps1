<#
.SYNOPSIS
    Creates Microsoft 365 group that is only enabled for Outlook.
.DESCRIPTION
    Create new modern distribution group. This will create a Microsoft 365 group with only Outlook attributes. 
    SharePoint, Teams, etc. will not be created for these groups. 
    The advantage of these over distribution groups in Exchange is that they can be dynamic groups. 
    Also, static groups can have access reviews, if needed. 
    
    The created group will: 
        - have the welcome message disabled
        - be hidden from the Groups list in Outlook
        - shown in the Global address list
.EXAMPLE
    PS > New-ModernDistributionGroup -Name "New Group"

    Create a new group 
.EXAMPLE     
    PS > New-ModernDistributionGroup -Name "New Group" -Description "This is a new group" -PrimarySMTPAddress "NewGroup@contoso.com"
 
    Create a new group, specifying description and primary SMTP address. 
.EXAMPLE     
    PS > New-ModernDistributionGroup -Name "New Group" -ProxyAddress "new@contoso.com" -ReceiveFromExternal
 
    Create a new group, specifying proxy address and receiveFromExternal.
.EXAMPLE 
    PS > New-ModernDistributionGroup -Name "New Group" -MembershipRule '(user.department -eq "Information Technology Team")'
 
    Create a new group with dynamic membership 
.EXAMPLE
    $Splat = @{
            Name                = 'New Group'
            MembershipRule      = '(user.mailNickname -eq "username")'
            Description         = 'This is a new group'
            PrimarySMTPAddress  = 'New_Group@contoso.com'
            ProxyAddresses      = 'nGroup@contoso.com', 'ThisIsANewGroup@contoso.com'
            ReceiveFromExternal = $true
        }
    PS > New-ModernDistributionGroup @Splat

    Create a new group with a splat.
.NOTES
    1) By creating the group using the Exchange cmdlets (New- and Set-UnifiedGroup) we can suppress the welcome message.
        - Both suppressing the welcome message and not showing in Teams are important
    2) After creating the group using EXO, it can then be converted to a dynamic group using Update-MgGroup.

    Written by Peter Kirschman
#>

function New-ModernDistributionGroup {
    [CmdletBinding(HelpUri = 'https://mountain.atlassian.net/wiki/spaces/ST1/pages/2741043230/New-ModernDistributionGroup')]
    [Alias()]
    [OutputType([String])]
    Param (
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true)]
        [string]
        $Name,
        [Parameter(
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true)]
        [string]
        $MembershipRule,
        [Parameter(
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true)]
        [string]
        $Description,
        [Parameter(
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true)]
        [string]
        $PrimarySMTPAddress,
        [Parameter(
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true)]
        [string]
        $ProxyAddress
    )
    begin {
        $BeginError = @()
        Write-Verbose "Checking prerequisites are loaded"
        if (-not (Get-Command New-UnifiedGroup -ErrorAction SilentlyContinue)) {
            $BeginError += "Not connected to ExchangeOnline"
        }

        try { 
            Get-MGContext -ErrorAction Stop | Out-Null
        } 
        catch [Microsoft.Open.Azure.AD.CommonLibrary.AadNeedAuthenticationException] { 
            $BeginError += "Not connected to MGGraph"
        }

        if ($BeginError) {
            Write-Error $($BeginError -join " - ") -ErrorAction Stop
        }

        function buildMailAlias([string]$Alias) {

            $IllegalCharacters = 0..34 + 39..41 + 44, 47 + 58..60 + 62 + 64 + 91..93 + 127..160
            $IllegalCharacters += [int][char]'.'
        
            #pull IllegalCharacters
            foreach ($c in $IllegalCharacters) {
                $escaped = [regex]::Escape([char]$c)
        
                if ($Alias -match $escaped) {
                    Write-Warning "Illegal character code '$c' detected in User Alias: [$Alias], removing"
                    $Alias = $Alias -replace $escaped
                }
            }
            return $Alias
        }
    }
    
    process {
        $MGGroup = $null
        Write-Verbose "Creating new group [$Name]"
        Write-Verbose "Building mail alias"
        $Alias = buildMailAlias -Alias $Name
        Write-Verbose "Mail alias is [$Alias]"

        if (-not $PrimarySMTPAddress) {
            $PrimarySmtpAddress = $Alias + '@contoso.com'
        }

        Write-Verbose "Creating new DL with primary SMTP address [$PrimarySMTPAddress]"

        $NewGroupSplat = @{
            DisplayName             = $Name
            Name                    = $Alias
            PrimarySmtpAddress      = $PrimarySmtpAddress
            AutoSubscribeNewMembers = $true
            SubscriptionEnabled     = $true
            AccessType              = 'Private'
        }
        Write-Verbose "Running New-UnifiedGroup"
        
        $NewGroup = New-UnifiedGroup @NewGroupSplat

        Write-Verbose "Hide Welcome Message and hide from Outlook"
        Set-UnifiedGroup -Identity $NewGroup.Id -UnifiedGroupWelcomeMessageEnabled:$false -HiddenFromExchangeClientsEnabled
        Write-Verbose "Show in the GAL."
        #This is necessary because the -HiddenFromExchangeClientsEnabled switch flips HiddenFromAddressListsEnabled to false.
        Set-UnifiedGroup -Identity $NewGroup.Id -HiddenFromAddressListsEnabled:$false

        if ($ProxyAddresses) {
            foreach ($ProxyAddress in $ProxyAddresses) {
                Set-UnifiedGroup -Identity $NewGroup.Name -EmailAddresses  @{Add = $ProxyAddress }
            }
        }

        if ($ReceiveFromExternal) {
            Set-UnifiedGroup -Identity $NewGroup.Name -RequireSenderAuthenticationEnabled $false
        }

        $Loop = 0
        while (-not $MGGroup -or ($Loop -gt 12)) {
            try { $MGGroup = Get-MGgroup -GroupId $NewGroup.ExternalDirectoryObjectId -ErrorAction Stop }
            catch { <# Empty catch #> }
            if (-not $MGGroup) {
                $Loop++
                Write-Verbose "MG group not found. Sleeping for 10 seconds and trying again."
                Start-Sleep -Seconds 10
            }
            else { Write-Verbose "MG group [$($MGGroup.DisplayName)] found " }
        }
        if (-not $MGGroup) {
            Write-Error "MG group not found" -ErrorAction Stop
        }

        if ($MembershipRule) {
            Write-Verbose "Convert to dynamic and add membership rule"
            #https://docs.microsoft.com/en-us/azure/active-directory/enterprise-users/groups-change-type
            $dynamicGroupTypeString = "DynamicMembership"
            [System.Collections.ArrayList]$groupTypes = $MGGroup.GroupTypes
            $groupTypes.Add($dynamicGroupTypeString) | Out-Null
            Update-MgGroup -GroupId $MGGroup.Id -MembershipRule $MembershipRule -MembershipRuleProcessingState "On" -GroupTypes $groupTypes.ToArray() -ErrorAction Continue | Out-Null

        }
        if ($Description) {
            Write-Verbose "Getting MG group"
            Write-Verbose "Set description [$Description]"
            Update-MgGroup -GroupId $MGGroup.Id -Description $Description -ErrorAction Continue | Out-Null
        }
        Get-MGgroup -GroupId $NewGroup.ExternalDirectoryObjectId | Select-Object DisplayName, Mail, Description, MembershipRule
    }
}

