Import-Module ActiveDirectory

$SamAccountName = "sjohnson"
$user = Get-ADUser -Filter "SamAccountName -eq '$SamAccountName'" -ErrorAction SilentlyContinue

if ($null -eq $user) {
    Write-Host "Sarah Johnson does not exist. The lab is already ready for Test 1."
    return
}

Write-Host "Deleting test account: $($user.DistinguishedName)"
Remove-ADUser -Identity $user.DistinguishedName -Confirm:$false

$verify = Get-ADUser -Filter "SamAccountName -eq '$SamAccountName'" -ErrorAction SilentlyContinue
if ($null -eq $verify) {
    Write-Host "[SUCCESS] Sarah Johnson (sjohnson) was deleted. Ready for Workday-Hire.csv."
} else {
    Write-Error "Sarah Johnson still exists. Reset failed."
}
