<#
.SYNOPSIS
    Employee Identity Lifecycle Automation
.DESCRIPTION
    Reconciles employee identity and access from Workday CSV data across:
      - Active Directory
      - Microsoft Entra ID
      - Intune-targeting Entra groups
      - Azure Virtual Desktop application-group access

    Supported lifecycle events:
      - Hire
      - Role change
      - Department transfer
      - Termination

    Design principles:
      - Workday is the authoritative HR source.
      - Role/department mappings define desired access.
      - The script compares desired vs. actual access and reconciles drift.
      - Dry-run mode is enabled with -WhatIf.
      - Every material action is written to an audit log.
      - Re-running the same input is intended to be idempotent.

.NOTES
    Portfolio/lab reference implementation.
    Test in a non-production tenant/domain before use.
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory)]
    [ValidateScript({ Test-Path $_ })]
    [string]$WorkdayCsv,

    [Parameter()]
    [ValidateScript({ Test-Path $_ })]
    [string]$ConfigPath = "$PSScriptRoot\AccessMappings.json",

    [Parameter()]
    [string]$AuditLogPath = "$PSScriptRoot\Audit\EmployeeIdentityAudit.jsonl",

    [Parameter()]
    [switch]$AuditOnly,

    [Parameter()]
    [switch]$SkipCloud,

    [Parameter()]
    [switch]$SkipActiveDirectory
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# Utility / logging
# ---------------------------------------------------------------------------

function Write-AuditEvent {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$EmployeeId,
        [Parameter(Mandatory)][string]$UserPrincipalName,
        [Parameter(Mandatory)][string]$Event,
        [Parameter(Mandatory)][string]$System,
        [Parameter(Mandatory)][string]$Action,
        [Parameter(Mandatory)][ValidateSet('Planned','Success','Skipped','Warning','Failed')][string]$Result,
        [string]$Details = ''
    )

    $directory = Split-Path -Parent $AuditLogPath
    if ($directory -and -not (Test-Path $directory)) {
        New-Item -ItemType Directory -Path $directory -Force | Out-Null
    }

    $record = [ordered]@{
        Timestamp         = (Get-Date).ToUniversalTime().ToString('o')
        EmployeeId        = $EmployeeId
        UserPrincipalName = $UserPrincipalName
        LifecycleEvent    = $Event
        System            = $System
        Action            = $Action
        Result            = $Result
        Details           = $Details
        RunBy             = [Environment]::UserName
        Computer          = $env:COMPUTERNAME
    }

    ($record | ConvertTo-Json -Compress) | Add-Content -Path $AuditLogPath -Encoding UTF8
}

function Invoke-ManagedChange {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][pscustomobject]$Employee,
        [Parameter(Mandatory)][string]$System,
        [Parameter(Mandatory)][string]$Action,
        [Parameter(Mandatory)][scriptblock]$ScriptBlock,
        [string]$Target = $Employee.UserPrincipalName
    )

    if ($AuditOnly) {
        Write-Host "[AUDIT ONLY] $System :: $Action :: $Target"
        Write-AuditEvent -EmployeeId $Employee.EmployeeId `
                         -UserPrincipalName $Employee.UserPrincipalName `
                         -Event $Employee.LifecycleEvent `
                         -System $System -Action $Action -Result Planned `
                         -Details "AuditOnly: no change made."
        return
    }

    if ($PSCmdlet.ShouldProcess($Target, "$System - $Action")) {
        try {
            & $ScriptBlock
            Write-Host "[SUCCESS] $System :: $Action :: $Target"
            Write-AuditEvent -EmployeeId $Employee.EmployeeId `
                             -UserPrincipalName $Employee.UserPrincipalName `
                             -Event $Employee.LifecycleEvent `
                             -System $System -Action $Action -Result Success
        }
        catch {
            Write-AuditEvent -EmployeeId $Employee.EmployeeId `
                             -UserPrincipalName $Employee.UserPrincipalName `
                             -Event $Employee.LifecycleEvent `
                             -System $System -Action $Action -Result Failed `
                             -Details $_.Exception.Message
            throw
        }
    }
    else {
        Write-AuditEvent -EmployeeId $Employee.EmployeeId `
                         -UserPrincipalName $Employee.UserPrincipalName `
                         -Event $Employee.LifecycleEvent `
                         -System $System -Action $Action -Result Skipped `
                         -Details "ShouldProcess declined or -WhatIf used."
    }
}

function ConvertTo-StringArray {
    param($Value)

    if ($null -eq $Value) { return @() }
    if ($Value -is [string]) {
        if ([string]::IsNullOrWhiteSpace($Value)) { return @() }
        return @($Value)
    }
    return @($Value | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
}

function Compare-AccessSet {
    [CmdletBinding()]
    param(
        [string[]]$Current,
        [string[]]$Desired
    )

    $currentNormalized = @($Current | Sort-Object -Unique)
    $desiredNormalized = @($Desired | Sort-Object -Unique)

    [pscustomobject]@{
        Add    = @($desiredNormalized | Where-Object { $_ -notin $currentNormalized })
        Remove = @($currentNormalized | Where-Object { $_ -notin $desiredNormalized })
        Keep   = @($desiredNormalized | Where-Object { $_ -in $currentNormalized })
    }
}

# ---------------------------------------------------------------------------
# Configuration / Workday import
# ---------------------------------------------------------------------------

function Import-AccessConfiguration {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)

    $config = Get-Content -Path $Path -Raw | ConvertFrom-Json

    foreach ($required in 'Organization','Departments','Roles') {
        if (-not $config.PSObject.Properties.Name.Contains($required)) {
            throw "Configuration missing required section '$required'."
        }
    }

    return $config
}

function Import-WorkdayEmployees {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)]$Config
    )

    $rows = Import-Csv -Path $Path
    if (-not $rows) {
        throw "Workday CSV contains no employee records."
    }

    $requiredColumns = @(
        'EmployeeId','FirstName','LastName','UserPrincipalName',
        'Department','JobTitle','ManagerUpn','Status','LifecycleEvent'
    )

    foreach ($column in $requiredColumns) {
        if ($column -notin $rows[0].PSObject.Properties.Name) {
            throw "Workday CSV is missing required column '$column'."
        }
    }

    foreach ($row in $rows) {
        $event = $row.LifecycleEvent.Trim()
        if ($event -notin @('Hire','RoleChange','DepartmentTransfer','Termination','Reconcile')) {
            throw "Unsupported LifecycleEvent '$event' for EmployeeId '$($row.EmployeeId)'."
        }

        [pscustomobject]@{
            EmployeeId        = $row.EmployeeId.Trim()
            FirstName         = $row.FirstName.Trim()
            LastName          = $row.LastName.Trim()
            DisplayName       = "$($row.FirstName.Trim()) $($row.LastName.Trim())"
            UserPrincipalName = $row.UserPrincipalName.Trim().ToLowerInvariant()
            SamAccountName    = if ($row.SamAccountName) { $row.SamAccountName.Trim() } else { ($row.UserPrincipalName -split '@')[0] }
            Department        = $row.Department.Trim()
            JobTitle          = $row.JobTitle.Trim()
            ManagerUpn        = $row.ManagerUpn.Trim().ToLowerInvariant()
            Status            = $row.Status.Trim()
            LifecycleEvent    = $event
            Location          = if ($row.Location) { $row.Location.Trim() } else { '' }
            EmploymentType    = if ($row.EmploymentType) { $row.EmploymentType.Trim() } else { '' }
        }
    }
}

# ---------------------------------------------------------------------------
# Desired-state calculation
# ---------------------------------------------------------------------------

function Get-DesiredAccess {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][pscustomobject]$Employee,
        [Parameter(Mandatory)]$Config
    )

    $department = $Config.Departments.PSObject.Properties |
        Where-Object Name -eq $Employee.Department |
        Select-Object -ExpandProperty Value -First 1

    if (-not $department -and $Employee.LifecycleEvent -ne 'Termination') {
        throw "No department mapping exists for '$($Employee.Department)'."
    }

    $role = $Config.Roles.PSObject.Properties |
        Where-Object Name -eq $Employee.JobTitle |
        Select-Object -ExpandProperty Value -First 1

    if (-not $role -and $Employee.LifecycleEvent -ne 'Termination') {
        throw "No role mapping exists for '$($Employee.JobTitle)'."
    }

    if ($Employee.LifecycleEvent -eq 'Termination' -or $Employee.Status -match 'Terminated|Inactive') {
        return [pscustomobject]@{
            OuPath       = $Config.Organization.DisabledUsersOu
            ADGroups     = @()
            EntraGroups  = @()
            IntuneGroups = @()
            AVD          = @()
        }
    }

    $adGroups = @(
        ConvertTo-StringArray $department.ADGroups
        ConvertTo-StringArray $role.ADGroups
    ) | Sort-Object -Unique

    $entraGroups = @(
        ConvertTo-StringArray $department.EntraGroups
        ConvertTo-StringArray $role.EntraGroups
    ) | Sort-Object -Unique

    # Intune assignment is intentionally group-based. Existing Intune policies
    # should target these Entra security groups.
    $intuneGroups = @(
        ConvertTo-StringArray $department.IntuneGroups
        ConvertTo-StringArray $role.IntuneGroups
    ) | Sort-Object -Unique

    $avdAssignments = @(
        ConvertTo-StringArray $department.AVD
        ConvertTo-StringArray $role.AVD
    ) | Sort-Object -Unique

    [pscustomobject]@{
        OuPath       = [string]$department.OuPath
        ADGroups     = $adGroups
        EntraGroups  = $entraGroups
        IntuneGroups = $intuneGroups
        AVD          = $avdAssignments
    }
}

# ---------------------------------------------------------------------------
# Active Directory
# ---------------------------------------------------------------------------

function Ensure-ADModule {
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        throw "ActiveDirectory PowerShell module is not installed."
    }
    Import-Module ActiveDirectory -ErrorAction Stop
}

function Get-ADEmployee {
    param([pscustomobject]$Employee)

    Get-ADUser -Filter "employeeID -eq '$($Employee.EmployeeId)'" -Properties * -ErrorAction SilentlyContinue |
        Select-Object -First 1
}

function New-ADEmployee {
    [CmdletBinding()]
    param(
        [pscustomobject]$Employee,
        [pscustomobject]$Desired,
        $Config
    )

    $existing = Get-ADEmployee -Employee $Employee
    if ($existing) {
        Write-AuditEvent -EmployeeId $Employee.EmployeeId `
            -UserPrincipalName $Employee.UserPrincipalName `
            -Event $Employee.LifecycleEvent -System 'ActiveDirectory' `
            -Action 'Create user' -Result Skipped -Details 'Account already exists.'
        return $existing
    }

    if (-not $Desired.OuPath) {
        throw "No OU path calculated for $($Employee.UserPrincipalName)."
    }

    $password = ConvertTo-SecureString $Config.Organization.InitialPassword -AsPlainText -Force

    Invoke-ManagedChange -Employee $Employee -System ActiveDirectory -Action 'Create user' -ScriptBlock {
        New-ADUser `
            -Name $Employee.DisplayName `
            -GivenName $Employee.FirstName `
            -Surname $Employee.LastName `
            -DisplayName $Employee.DisplayName `
            -SamAccountName $Employee.SamAccountName `
            -UserPrincipalName $Employee.UserPrincipalName `
            -EmployeeID $Employee.EmployeeId `
            -Department $Employee.Department `
            -Title $Employee.JobTitle `
            -Path $Desired.OuPath `
            -AccountPassword $password `
            -Enabled $true `
            -ChangePasswordAtLogon $true
    }

    Get-ADEmployee -Employee $Employee
}

function Set-ADEmployeeAttributes {
    [CmdletBinding()]
    param(
        [pscustomobject]$Employee,
        $AdUser
    )

    if (-not $AdUser) { return }

    $replace = @{
        department = $Employee.Department
        title      = $Employee.JobTitle
    }

    Invoke-ManagedChange -Employee $Employee -System ActiveDirectory -Action 'Update employee attributes' -ScriptBlock {
        Set-ADUser -Identity $AdUser.DistinguishedName -Replace $replace
    }

    if ($Employee.ManagerUpn) {
        $manager = Get-ADUser -Filter "UserPrincipalName -eq '$($Employee.ManagerUpn)'" -ErrorAction SilentlyContinue
        if ($manager -and $AdUser.Manager -ne $manager.DistinguishedName) {
            Invoke-ManagedChange -Employee $Employee -System ActiveDirectory -Action 'Update manager' -ScriptBlock {
                Set-ADUser -Identity $AdUser.DistinguishedName -Manager $manager.DistinguishedName
            }
        }
    }
}

function Set-ADEmployeeOu {
    [CmdletBinding()]
    param(
        [pscustomobject]$Employee,
        $AdUser,
        [string]$DesiredOu
    )

    if (-not $AdUser -or -not $DesiredOu) { return }

    $currentParent = ($AdUser.DistinguishedName -split ',', 2)[1]
    if ($currentParent -ne $DesiredOu) {
        Invoke-ManagedChange -Employee $Employee -System ActiveDirectory -Action "Move to OU '$DesiredOu'" -ScriptBlock {
            Move-ADObject -Identity $AdUser.DistinguishedName -TargetPath $DesiredOu
        }
    }
}

function Sync-ADGroups {
    [CmdletBinding()]
    param(
        [pscustomobject]$Employee,
        $AdUser,
        [string[]]$DesiredGroups,
        $Config
    )

    if (-not $AdUser) { return }

    $managedPrefix = [string]$Config.Organization.ManagedADGroupPrefix
    $currentGroups = @(
        Get-ADPrincipalGroupMembership -Identity $AdUser |
        Where-Object { $_.Name -like "$managedPrefix*" } |
        Select-Object -ExpandProperty Name
    )

    $diff = Compare-AccessSet -Current $currentGroups -Desired $DesiredGroups

    foreach ($groupName in $diff.Add) {
        $group = Get-ADGroup -Identity $groupName -ErrorAction Stop
        Invoke-ManagedChange -Employee $Employee -System ActiveDirectory -Action "Add group '$groupName'" -ScriptBlock {
            Add-ADGroupMember -Identity $group -Members $AdUser
        }
    }

    foreach ($groupName in $diff.Remove) {
        $group = Get-ADGroup -Identity $groupName -ErrorAction Stop
        Invoke-ManagedChange -Employee $Employee -System ActiveDirectory -Action "Remove group '$groupName'" -ScriptBlock {
            Remove-ADGroupMember -Identity $group -Members $AdUser -Confirm:$false
        }
    }
}

function Disable-ADEmployee {
    [CmdletBinding()]
    param(
        [pscustomobject]$Employee,
        $AdUser,
        $Config
    )

    if (-not $AdUser) { return }

    if ($AdUser.Enabled) {
        Invoke-ManagedChange -Employee $Employee -System ActiveDirectory -Action 'Disable user' -ScriptBlock {
            Disable-ADAccount -Identity $AdUser.DistinguishedName
        }
    }

    # Remove only groups managed by this automation.
    $managedPrefix = [string]$Config.Organization.ManagedADGroupPrefix
    $managedMemberships = @(
        Get-ADPrincipalGroupMembership -Identity $AdUser |
        Where-Object { $_.Name -like "$managedPrefix*" }
    )

    foreach ($group in $managedMemberships) {
        Invoke-ManagedChange -Employee $Employee -System ActiveDirectory -Action "Remove group '$($group.Name)'" -ScriptBlock {
            Remove-ADGroupMember -Identity $group -Members $AdUser -Confirm:$false
        }
    }

    Set-ADEmployeeOu -Employee $Employee -AdUser $AdUser -DesiredOu $Config.Organization.DisabledUsersOu
}

# ---------------------------------------------------------------------------
# Microsoft Graph / Entra ID / Intune targeting groups
# ---------------------------------------------------------------------------

function Ensure-GraphConnection {
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
        throw "Microsoft Graph PowerShell SDK is not installed."
    }

    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    Import-Module Microsoft.Graph.Users -ErrorAction Stop
    Import-Module Microsoft.Graph.Groups -ErrorAction Stop

    $context = Get-MgContext
    if (-not $context) {
        Connect-MgGraph -Scopes @(
            'User.ReadWrite.All',
            'Group.ReadWrite.All',
            'Directory.ReadWrite.All'
        ) -NoWelcome
    }
}

function Get-EntraEmployee {
    param([pscustomobject]$Employee)

    Get-MgUser -Filter "userPrincipalName eq '$($Employee.UserPrincipalName)'" `
        -Property Id,DisplayName,UserPrincipalName,AccountEnabled,Department,JobTitle `
        -ConsistencyLevel eventual -ErrorAction SilentlyContinue |
        Select-Object -First 1
}

function Get-EntraGroupByName {
    param([Parameter(Mandatory)][string]$DisplayName)

    $escaped = $DisplayName.Replace("'", "''")
    Get-MgGroup -Filter "displayName eq '$escaped'" -Property Id,DisplayName |
        Select-Object -First 1
}

function Test-EntraGroupMembership {
    param(
        [Parameter(Mandatory)][string]$UserId,
        [Parameter(Mandatory)][string]$GroupId
    )

    $memberships = Get-MgUserMemberOf -UserId $UserId -All
    return $GroupId -in @($memberships.Id)
}

function Sync-EntraGroups {
    [CmdletBinding()]
    param(
        [pscustomobject]$Employee,
        $EntraUser,
        [string[]]$DesiredGroups,
        [string]$SystemLabel,
        $Config
    )

    if (-not $EntraUser) {
        Write-AuditEvent -EmployeeId $Employee.EmployeeId `
            -UserPrincipalName $Employee.UserPrincipalName `
            -Event $Employee.LifecycleEvent -System $SystemLabel `
            -Action 'Reconcile groups' -Result Warning `
            -Details 'Entra user not found. Hybrid synchronization may not have completed yet.'
        return
    }

    $managedPrefix = [string]$Config.Organization.ManagedEntraGroupPrefix
    $currentManaged = @()

    $memberships = Get-MgUserMemberOf -UserId $EntraUser.Id -All
    foreach ($membership in $memberships) {
        if ($membership.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.group') {
            $group = Get-MgGroup -GroupId $membership.Id -Property Id,DisplayName -ErrorAction SilentlyContinue
            if ($group -and $group.DisplayName -like "$managedPrefix*") {
                $currentManaged += $group.DisplayName
            }
        }
    }

    $diff = Compare-AccessSet -Current $currentManaged -Desired $DesiredGroups

    foreach ($groupName in $diff.Add) {
        $group = Get-EntraGroupByName -DisplayName $groupName
        if (-not $group) {
            throw "Entra group '$groupName' was not found."
        }

        Invoke-ManagedChange -Employee $Employee -System $SystemLabel -Action "Add group '$groupName'" -ScriptBlock {
            New-MgGroupMemberByRef -GroupId $group.Id -BodyParameter @{
                '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$($EntraUser.Id)"
            }
        }
    }

    foreach ($groupName in $diff.Remove) {
        $group = Get-EntraGroupByName -DisplayName $groupName
        if ($group) {
            Invoke-ManagedChange -Employee $Employee -System $SystemLabel -Action "Remove group '$groupName'" -ScriptBlock {
                Remove-MgGroupMemberByRef -GroupId $group.Id -DirectoryObjectId $EntraUser.Id
            }
        }
    }
}

function Disable-EntraEmployee {
    [CmdletBinding()]
    param(
        [pscustomobject]$Employee,
        $EntraUser,
        $Config
    )

    if (-not $EntraUser) { return }

    if ($EntraUser.AccountEnabled) {
        Invoke-ManagedChange -Employee $Employee -System EntraID -Action 'Block sign-in' -ScriptBlock {
            Update-MgUser -UserId $EntraUser.Id -AccountEnabled:$false
        }
    }

    Invoke-ManagedChange -Employee $Employee -System EntraID -Action 'Revoke active sessions' -ScriptBlock {
        Revoke-MgUserSignInSession -UserId $EntraUser.Id | Out-Null
    }

    # Remove only groups controlled by this automation.
    $managedPrefix = [string]$Config.Organization.ManagedEntraGroupPrefix
    $memberships = Get-MgUserMemberOf -UserId $EntraUser.Id -All

    foreach ($membership in $memberships) {
        if ($membership.AdditionalProperties.'@odata.type' -ne '#microsoft.graph.group') { continue }

        $group = Get-MgGroup -GroupId $membership.Id -Property Id,DisplayName -ErrorAction SilentlyContinue
        if ($group -and $group.DisplayName -like "$managedPrefix*") {
            Invoke-ManagedChange -Employee $Employee -System EntraID -Action "Remove group '$($group.DisplayName)'" -ScriptBlock {
                Remove-MgGroupMemberByRef -GroupId $group.Id -DirectoryObjectId $EntraUser.Id
            }
        }
    }
}

# ---------------------------------------------------------------------------
# Azure Virtual Desktop
# Configuration format:
# "AVD": [
#   "ResourceGroupName|ApplicationGroupName"
# ]
# ---------------------------------------------------------------------------

function Ensure-AzureConnection {
    if (-not (Get-Module -ListAvailable -Name Az.Accounts)) {
        throw "Az PowerShell modules are not installed."
    }

    Import-Module Az.Accounts -ErrorAction Stop
    Import-Module Az.Resources -ErrorAction Stop
    Import-Module Az.DesktopVirtualization -ErrorAction Stop

    if (-not (Get-AzContext -ErrorAction SilentlyContinue)) {
        Connect-AzAccount | Out-Null
    }
}

function ConvertFrom-AVDMapping {
    param([Parameter(Mandatory)][string]$Mapping)

    $parts = $Mapping -split '\|', 2
    if ($parts.Count -ne 2) {
        throw "Invalid AVD mapping '$Mapping'. Expected 'ResourceGroupName|ApplicationGroupName'."
    }

    [pscustomobject]@{
        ResourceGroupName   = $parts[0]
        ApplicationGroupName = $parts[1]
    }
}

function Get-CurrentAVDAssignments {
    [CmdletBinding()]
    param(
        [pscustomobject]$Employee,
        $Config
    )

    $results = @()
    foreach ($mapping in (ConvertTo-StringArray $Config.Organization.ManagedAVDApplicationGroups)) {
        $avd = ConvertFrom-AVDMapping $mapping
        $appGroup = Get-AzWvdApplicationGroup `
            -ResourceGroupName $avd.ResourceGroupName `
            -Name $avd.ApplicationGroupName `
            -ErrorAction SilentlyContinue

        if (-not $appGroup) { continue }

        $assignment = Get-AzRoleAssignment `
            -SignInName $Employee.UserPrincipalName `
            -Scope $appGroup.Id `
            -RoleDefinitionName 'Desktop Virtualization User' `
            -ErrorAction SilentlyContinue

        if ($assignment) {
            $results += $mapping
        }
    }
    return $results
}

function Sync-AVDAccess {
    [CmdletBinding()]
    param(
        [pscustomobject]$Employee,
        [string[]]$DesiredAVD,
        $Config
    )

    $current = Get-CurrentAVDAssignments -Employee $Employee -Config $Config
    $diff = Compare-AccessSet -Current $current -Desired $DesiredAVD

    foreach ($mapping in $diff.Add) {
        $avd = ConvertFrom-AVDMapping $mapping
        Invoke-ManagedChange -Employee $Employee -System AzureVirtualDesktop -Action "Grant '$($avd.ApplicationGroupName)'" -ScriptBlock {
            New-AzRoleAssignment `
                -SignInName $Employee.UserPrincipalName `
                -ResourceName $avd.ApplicationGroupName `
                -ResourceGroupName $avd.ResourceGroupName `
                -RoleDefinitionName 'Desktop Virtualization User' `
                -ResourceType 'Microsoft.DesktopVirtualization/applicationGroups' | Out-Null
        }
    }

    foreach ($mapping in $diff.Remove) {
        $avd = ConvertFrom-AVDMapping $mapping
        $appGroup = Get-AzWvdApplicationGroup `
            -ResourceGroupName $avd.ResourceGroupName `
            -Name $avd.ApplicationGroupName

        $assignments = Get-AzRoleAssignment `
            -SignInName $Employee.UserPrincipalName `
            -Scope $appGroup.Id `
            -RoleDefinitionName 'Desktop Virtualization User' `
            -ErrorAction SilentlyContinue

        foreach ($assignment in $assignments) {
            Invoke-ManagedChange -Employee $Employee -System AzureVirtualDesktop -Action "Revoke '$($avd.ApplicationGroupName)'" -ScriptBlock {
                Remove-AzRoleAssignment -InputObject $assignment
            }
        }
    }
}

# ---------------------------------------------------------------------------
# Access reporting / reconciliation
# ---------------------------------------------------------------------------

function Get-EmployeeAccessSnapshot {
    [CmdletBinding()]
    param(
        [pscustomobject]$Employee,
        [pscustomobject]$Desired,
        $Config,
        $AdUser,
        $EntraUser
    )

    $adCurrent = @()
    if ($AdUser) {
        $prefix = [string]$Config.Organization.ManagedADGroupPrefix
        $adCurrent = @(
            Get-ADPrincipalGroupMembership -Identity $AdUser |
            Where-Object { $_.Name -like "$prefix*" } |
            Select-Object -ExpandProperty Name
        )
    }

    $entraCurrent = @()
    if ($EntraUser) {
        $prefix = [string]$Config.Organization.ManagedEntraGroupPrefix
        $memberships = Get-MgUserMemberOf -UserId $EntraUser.Id -All
        foreach ($membership in $memberships) {
            if ($membership.AdditionalProperties.'@odata.type' -ne '#microsoft.graph.group') { continue }
            $group = Get-MgGroup -GroupId $membership.Id -Property Id,DisplayName -ErrorAction SilentlyContinue
            if ($group -and $group.DisplayName -like "$prefix*") {
                $entraCurrent += $group.DisplayName
            }
        }
    }

    $avdCurrent = @()
    if (-not $SkipCloud) {
        try { $avdCurrent = Get-CurrentAVDAssignments -Employee $Employee -Config $Config }
        catch { $avdCurrent = @() }
    }

    [pscustomobject]@{
        EmployeeId        = $Employee.EmployeeId
        UserPrincipalName = $Employee.UserPrincipalName
        Department        = $Employee.Department
        JobTitle          = $Employee.JobTitle
        LifecycleEvent    = $Employee.LifecycleEvent
        AD                = Compare-AccessSet -Current $adCurrent -Desired $Desired.ADGroups
        Entra             = Compare-AccessSet -Current $entraCurrent -Desired (@($Desired.EntraGroups) + @($Desired.IntuneGroups) | Sort-Object -Unique)
        AVD               = Compare-AccessSet -Current $avdCurrent -Desired $Desired.AVD
    }
}

function Show-AccessSnapshot {
    param([pscustomobject]$Snapshot)

    Write-Host ""
    Write-Host "============================================================"
    Write-Host "ACCESS RECONCILIATION: $($Snapshot.UserPrincipalName)"
    Write-Host "Event: $($Snapshot.LifecycleEvent)"
    Write-Host "Department: $($Snapshot.Department)"
    Write-Host "Role: $($Snapshot.JobTitle)"
    Write-Host "------------------------------------------------------------"

    foreach ($sectionName in 'AD','Entra','AVD') {
        $section = $Snapshot.$sectionName
        Write-Host "$sectionName"
        Write-Host "  Add:    $((@($section.Add) -join ', '))"
        Write-Host "  Remove: $((@($section.Remove) -join ', '))"
        Write-Host "  Keep:   $((@($section.Keep) -join ', '))"
    }
}

# ---------------------------------------------------------------------------
# Lifecycle orchestration
# ---------------------------------------------------------------------------

function Invoke-Onboarding {
    [CmdletBinding()]
    param(
        [pscustomobject]$Employee,
        [pscustomobject]$Desired,
        $Config
    )

    $adUser = $null
    if (-not $SkipActiveDirectory) {
        $adUser = New-ADEmployee -Employee $Employee -Desired $Desired -Config $Config
        if ($adUser) {
            Set-ADEmployeeAttributes -Employee $Employee -AdUser $adUser
            $adUser = Get-ADEmployee -Employee $Employee
            Set-ADEmployeeOu -Employee $Employee -AdUser $adUser -DesiredOu $Desired.OuPath
            Sync-ADGroups -Employee $Employee -AdUser $adUser -DesiredGroups $Desired.ADGroups -Config $Config
        }
    }

    $entraUser = $null
    if (-not $SkipCloud) {
        $entraUser = Get-EntraEmployee -Employee $Employee

        if ($entraUser) {
            Sync-EntraGroups -Employee $Employee -EntraUser $entraUser `
                -DesiredGroups $Desired.EntraGroups -SystemLabel EntraID -Config $Config

            Sync-EntraGroups -Employee $Employee -EntraUser $entraUser `
                -DesiredGroups $Desired.IntuneGroups -SystemLabel IntuneAccess -Config $Config

            Sync-AVDAccess -Employee $Employee -DesiredAVD $Desired.AVD -Config $Config
        }
        else {
            Write-Warning "Entra user not yet visible for $($Employee.UserPrincipalName). Cloud reconciliation can be re-run after sync."
        }
    }
}

function Invoke-EmployeeChange {
    [CmdletBinding()]
    param(
        [pscustomobject]$Employee,
        [pscustomobject]$Desired,
        $Config
    )

    $adUser = $null
    if (-not $SkipActiveDirectory) {
        $adUser = Get-ADEmployee -Employee $Employee
        if (-not $adUser) {
            throw "AD account not found for EmployeeId '$($Employee.EmployeeId)'."
        }

        Set-ADEmployeeAttributes -Employee $Employee -AdUser $adUser
        $adUser = Get-ADEmployee -Employee $Employee
        Set-ADEmployeeOu -Employee $Employee -AdUser $adUser -DesiredOu $Desired.OuPath
        $adUser = Get-ADEmployee -Employee $Employee
        Sync-ADGroups -Employee $Employee -AdUser $adUser -DesiredGroups $Desired.ADGroups -Config $Config
    }

    if (-not $SkipCloud) {
        $entraUser = Get-EntraEmployee -Employee $Employee
        if ($entraUser) {
            Sync-EntraGroups -Employee $Employee -EntraUser $entraUser `
                -DesiredGroups $Desired.EntraGroups -SystemLabel EntraID -Config $Config

            Sync-EntraGroups -Employee $Employee -EntraUser $entraUser `
                -DesiredGroups $Desired.IntuneGroups -SystemLabel IntuneAccess -Config $Config

            Sync-AVDAccess -Employee $Employee -DesiredAVD $Desired.AVD -Config $Config
        }
    }
}

function Invoke-Offboarding {
    [CmdletBinding()]
    param(
        [pscustomobject]$Employee,
        [pscustomobject]$Desired,
        $Config
    )

    if (-not $SkipActiveDirectory) {
        $adUser = Get-ADEmployee -Employee $Employee
        if ($adUser) {
            Disable-ADEmployee -Employee $Employee -AdUser $adUser -Config $Config
        }
    }

    if (-not $SkipCloud) {
        $entraUser = Get-EntraEmployee -Employee $Employee
        if ($entraUser) {
            Disable-EntraEmployee -Employee $Employee -EntraUser $entraUser -Config $Config
            Sync-AVDAccess -Employee $Employee -DesiredAVD @() -Config $Config
        }
    }
}

function Invoke-Reconciliation {
    [CmdletBinding()]
    param(
        [pscustomobject]$Employee,
        [pscustomobject]$Desired,
        $Config
    )

    Invoke-EmployeeChange -Employee $Employee -Desired $Desired -Config $Config
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

try {
    $config = Import-AccessConfiguration -Path $ConfigPath

    if (-not $SkipActiveDirectory) {
        Ensure-ADModule
    }

    if (-not $SkipCloud) {
        Ensure-GraphConnection
        Ensure-AzureConnection
    }

    $employees = @(Import-WorkdayEmployees -Path $WorkdayCsv -Config $config)

    foreach ($employee in $employees) {
        try {
            Write-Host ""
            Write-Host "Processing $($employee.EmployeeId) :: $($employee.UserPrincipalName) :: $($employee.LifecycleEvent)"

            $desired = Get-DesiredAccess -Employee $employee -Config $config

            $adUser = $null
            $entraUser = $null
            if (-not $SkipActiveDirectory) { $adUser = Get-ADEmployee -Employee $employee }
            if (-not $SkipCloud) { $entraUser = Get-EntraEmployee -Employee $employee }

            $snapshot = Get-EmployeeAccessSnapshot `
                -Employee $employee -Desired $desired -Config $config `
                -AdUser $adUser -EntraUser $entraUser

            Show-AccessSnapshot -Snapshot $snapshot

            switch ($employee.LifecycleEvent) {
                'Hire' {
                    Invoke-Onboarding -Employee $employee -Desired $desired -Config $config
                }
                'RoleChange' {
                    Invoke-EmployeeChange -Employee $employee -Desired $desired -Config $config
                }
                'DepartmentTransfer' {
                    Invoke-EmployeeChange -Employee $employee -Desired $desired -Config $config
                }
                'Termination' {
                    Invoke-Offboarding -Employee $employee -Desired $desired -Config $config
                }
                'Reconcile' {
                    Invoke-Reconciliation -Employee $employee -Desired $desired -Config $config
                }
            }
        }
        catch {
            Write-Error "Employee '$($employee.EmployeeId)' failed: $($_.Exception.Message)"
            Write-AuditEvent -EmployeeId $employee.EmployeeId `
                -UserPrincipalName $employee.UserPrincipalName `
                -Event $employee.LifecycleEvent `
                -System Orchestrator -Action 'Process employee' -Result Failed `
                -Details $_.Exception.Message
        }
    }
}
catch {
    Write-Error $_
    exit 1
}
