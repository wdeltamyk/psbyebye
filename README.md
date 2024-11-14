# psbyebye
PowerShell offboarding Script

Removes users from O365 groups based on the prefix of of "xEM -"
Including the " -" ensures employee's who just so happent to have "xem" as letters in their names are not offboarded by mistake.

It creates a log so you can see who was affected and what groups they were removed from.

Requires AzureAD and ExchangeOnlineManagement modules.
Install via:
Install-Module AzureAd | Install-Module ExchangeOnlineManagement
