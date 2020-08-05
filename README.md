# DeployPrinters
Rewrite of @ozthe2's DeployPrinters powershell script (https://github.com/ozthe2/Powershell/blob/master/Active-Directory/DeployPrinters).

https://fearthepanda.com/powershell/2014/11/18/PowerShell-How-to-speed-up-login-times-using-group-policy-preferences/ stated:

'I will be actively working on this script so that it also works with users in specific OU’s as well as Users being members of security groups. Alternatively, go ahead and make any mods yourself if you cannot be bothered to wait for me.'

So I did. ;)

Additional Features:

-Removes all shared printers before running through the XML (can be turned on by switching RASP to $true and off by switching it to $false).<br>
-Can check to see if a root member of an OU or down a path and act accordingly.<br>
-Checks AD Groups (including Nested) of the User and Computer to see if there is a Group match.<br>
-Checks for Computer Name and User Name matches.

## Update:<br>
My good friend and 'Coding Goddess' Katy has also made an excellent script for processing printers in Group Policy, you can check it out here. https://katynicholson.uk/2020/08/powershell-printer-script/
