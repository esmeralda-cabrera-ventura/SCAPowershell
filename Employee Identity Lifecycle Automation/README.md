# Employee Identity Lifecycle Automation

PowerShell Joiner-Mover-Leaver (JML) automation lab for keeping employee Active Directory access aligned with HR/Workday lifecycle changes.

## Current lab

- Domain controller: `DC01`
- Active Directory domain: `corp.ventura.local`
- NetBIOS domain: `VENTURA`
- Windows Server 2025 Standard Evaluation
- Active Directory + DNS
- PowerShell 5.1 compatible
- Local AD testing uses `-SkipCloud`

## OU structure

```text
corp.ventura.local
├── Employees
│   ├── Engineering
│   ├── Finance
│   └── Human Resources
├── Disabled Users
└── Groups
```

## AD access model

Department baseline groups:

```text
Finance
├── IAM-Finance-Users
└── IAM-Finance-SharedDrive

Engineering
├── IAM-Engineering-Users
└── IAM-Engineering-SharedDrive

Human Resources
├── IAM-HR-Users
└── IAM-HR-SharedDrive
```

Role-specific AD groups: none.

Every AD group is a department baseline group in the `Groups` OU. Job roles (Financial Analyst, Cloud Engineer, HR Specialist) have no AD groups, so they only drive the `Title` attribute in AD. Role access is mapped for the cloud targets only (Entra, Intune, AVD).

## Test accounts

Managers:

- Jennifer Dovorany — `jennifer.dovorany@corp.ventura.local`
- Robert Smith — `robert.smith@corp.ventura.local`

Test employee:

- Sarah Johnson
- Employee ID `10001`
- SamAccountName `sjohnson`
- UPN `sarah.johnson@corp.ventura.local`

## Project files

```text
EmployeeIdentityAutomation/
├── Invoke-EmployeeIdentityLifecycle.ps1
├── AccessMappings.json
├── Reset-TestEmployee.ps1
├── README.md
└── EmployeeLifeCycleTestCases/
    ├── Workday-Hire.csv
    ├── Workday-RoleChange.csv
    ├── Workday-DepartmentTransfer.csv
    └── Workday-Termination.csv
```

The `Audit` folder is created automatically at runtime.

## Test order

1. `Workday-Hire.csv`
2. `Workday-RoleChange.csv`
3. `Workday-DepartmentTransfer.csv`
4. `Workday-Termination.csv`

### Test 1 — Hire

Expected:

- Sarah is created
- Finance OU
- Financial Analyst
- Jennifer Dovorany as manager
- `IAM-Finance-Users`
- `IAM-Finance-SharedDrive`

### Test 2 — RoleChange

Expected:

- Sarah stays in Finance OU
- title changes to Cloud Engineer
- Jennifer remains manager
- AD groups unchanged:
  - `IAM-Finance-Users`
  - `IAM-Finance-SharedDrive`

### Test 3 — DepartmentTransfer

Expected:

- Sarah moves to Engineering OU
- department changes to Engineering
- title remains Cloud Engineer
- manager changes to Robert Smith
- Finance groups removed
- Engineering groups added:
  - `IAM-Engineering-Users`
  - `IAM-Engineering-SharedDrive`

### Test 4 — Termination

Expected:

- Sarah account disabled
- managed `IAM-*` groups removed
- Sarah moved to `OU=Disabled Users,DC=corp,DC=ventura,DC=local`

## Run a dry test

Example:

```powershell
.\Invoke-EmployeeIdentityLifecycle.ps1 `
  -WorkdayCsv .\EmployeeLifeCycleTestCases\Workday-Hire.csv `
  -ConfigPath .\AccessMappings.json `
  -SkipCloud `
  -WhatIf
```

Remove `-WhatIf` for the real execution.

## Verify Sarah

```powershell
Get-ADUser sjohnson `
  -Properties EmployeeID,Department,Title,Manager,UserPrincipalName,DistinguishedName,Enabled |
  Select-Object Name,SamAccountName,UserPrincipalName,EmployeeID,Department,Title,Manager,DistinguishedName,Enabled
```

```powershell
Get-ADPrincipalGroupMembership sjohnson |
  Select-Object Name |
  Sort-Object Name
```

## Reset before rerunning all four tests

Run:

```powershell
.\Reset-TestEmployee.ps1
```

This deletes only Sarah Johnson (`sjohnson`). It does not delete the managers, groups, or OUs.

## Audit logging

Runtime audit records are written to:

```text
C:\Lab\Audit\EmployeeIdentityAudit.jsonl
```

## Lab password

The lab-only initial password in `AccessMappings.json` is:

```text
TEMPORARY PASSWORD GOES HERE
```

Do not use this plaintext-password pattern in production. Use a secure secret store or generate temporary credentials at runtime.

## Cloud targets

The configuration also contains downstream mappings for:

- Microsoft Entra ID
- Intune-targeted Entra groups
- Azure Virtual Desktop

These remain part of the desired access model, but `-SkipCloud` prevents cloud execution in the current Hyper-V lab.
