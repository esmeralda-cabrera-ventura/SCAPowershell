Get-ADComputer -Filter * -Properties OperatingSystem, IPv4Address, LastLogonDate, Enabled |
Select-Object Name, Enabled, OperatingSystem, IPav4Address, LastLogonDate |
Sort-Object LastLogonDate -Descending |
Export-CSv ".\AD-Computer-Inventory.csv" -NoTypeInformation
