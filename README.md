# DeployPrinters
Rewrite of @ozthe2's DeployPrinters powershell script (https://github.com/ozthe2/Powershell/blob/master/Active-Directory/DeployPrinters).

https://fearthepanda.com/powershell/2014/11/18/PowerShell-How-to-speed-up-login-times-using-group-policy-preferences/ stated:

'I will be actively working on this script so that it also works with users in specific OU’s as well as Users being members of security groups. Alternatively, go ahead and make any mods yourself if you cannot be bothered to wait for me.'

So I did. ;)

Additional Features:

-Removes all shared printers before running through the XML (can be turned off by switching RASP to $false).

-Can check to see if a root member of an OU or down a path and act accordingly.

-Checked Nested AD Groups for User and Computer entries.
