# Unique ID of GPO with Printers listed.
$GPOUID = "{01234567-89AB-CDEF-0123-456789ABCDEF}"
# Remove All Shared Printers before reading GPO.
$RASP = $false

$DomainName = (gwmi win32_computersystem).Domain
$ShortDomainName = $DomainName.Split(".")[0]
$ComputerDN = ([ADSISEARCHER]"(&(objectCategory=computer)(objectClass=computer)(cn=$env:COMPUTERNAME))").FindOne().Properties.distinguishedname[0]
$ComputerGroups = ([ADSISEARCHER]"(member:1.2.840.113556.1.4.1941:=$ComputerDN)").FindAll().GetEnumerator() | ForEach-Object {$_.Properties}
$ComputerOU = $ComputerDN.substring(([string]$ComputerDN).IndexOf(",")+1)
$UserDN = ([ADSISEARCHER]"samaccountname=$($env:Username)").FindOne().Properties.distinguishedname[0]
$UserGroups = ([ADSISEARCHER]"(member:1.2.840.113556.1.4.1941:=$UserDN)").FindAll().GetEnumerator() | ForEach-Object {$_.Properties}
$UserOU = $ComputerDN.substring(([string]$ComputerDN).IndexOf(",")+1)
$UserSID = (New-Object System.Security.Principal.SecurityIdentifier(([ADSISearcher]"(&(objectClass=User)(cn=$env:Username))").FindOne().GetDirectoryEntry().ObjectSID.Value,0)).Value
$NetObject = New-Object -ComObject WScript.Network

[xml]$ListOfPrinters = Get-Content "\\$DomainName\SYSVOL\$DomainName\Policies\$GPOUID\User\Preferences\Printers\Printers.xml"

function funcAddPrinter
{
    $NetObject.AddWindowsPrinterConnection($PrinterRecord.Properties.path)
    If ($PrinterRecord.Properties.default)
    {
        $NetObject.SetDefaultPrinter($PrinterRecord.Properties.path)
    }
    #Write-Host "Added" $PrinterRecord.Properties.path $CachedP
}

function funcDeletePrinter
{
    $NetObject.RemovePrinterConnection($PrinterRecord.Properties.path)
    #Write-Host "Deleted" $PrinterRecord.Properties.path $CachedP
}

# Remove all shared printers.
If ($RASP -eq $true)
{
    ForEach ($DetectedPrinter in $NetObject.EnumPrinterConnections()){If ($DetectedPrinter.StartsWith("\\")) {$NetObject.RemovePrinterConnection($DetectedPrinter)}}
}

# Check for Printers in XML file.
ForEach ($PrinterObject in $ListOfPrinters.Printers.SharedPrinter)
{
    # Check each instance a printer is meantioned.
    ForEach ($PrinterRecord in $PrinterObject)
    {
        # Check through OU matches.
        If ($PrinterRecord.Filters.FilterOrgUnit -ne $NULL)
        {
            # Check if there is only one OU to check, or multiple.
            If ($PrinterRecord.Filters.FilterOrgUnit.Count -eq $NULL)
            {
                $CachedP = $PrinterRecord.Filters.FilterOrgUnit.name
                # Check if we are looking for a device in any OU under this branch, or if only looking for machines imediately in this OU (Direct Member).
                If (($PrinterRecord.Filters.FilterOrgUnit.directMember -ne $NULL) -and ($PrinterRecord.Filters.FilterOrgUnit.directMember -ne 0))
                {
                    # Check if we are comparing the User account OU, or the Computer account OU.
                    If ($PrinterRecord.Filters.FilterOrgUnit.userContext -eq 1)
                    {
                        # Do the following if only one of these conditions is true, 1) The User is in the OU, 2) A Not operator is used.
                        If (($UserOU.Contains($PrinterRecord.Filters.FilterOrgUnit.name)) -xor ($PrinterRecord.Filters.FilterOrgUnit.not -eq 1))
                        {If ($PrinterRecord.Properties.action -eq "D") {. funcDeletePrinter} Else {. funcAddPrinter}}
                    }
                    Else
                    {
                        # Do the following if only one of these conditions is true, 1) The Computer is in the OU, 2) A Not operator is used.
                        If (($ComputerOU.Contains($PrinterRecord.Filters.FilterOrgUnit.name)) -xor ($PrinterRecord.Filters.FilterOrgUnit.not -eq 1))
                        {If ($PrinterRecord.Properties.action -eq "D") {. funcDeletePrinter} Else {. funcAddPrinter}}
                    }
                }
                Else
                {
                    # Check if we are comparing the User account OU, or the Computer account OU.
                    If ($PrinterRecord.Filters.FilterOrgUnit.userContext -eq 1)
                    {
                        # Do the following if only one of these conditions is true, 1) The User is in the imediate OU, 2) A Not operator is used.
                        If (($UserOU -eq $PrinterRecord.Filters.FilterOrgUnit.name) -xor ($PrinterRecord.Filters.FilterOrgUnit.not -eq 1))
                        {If ($PrinterRecord.Properties.action -eq "D") {. funcDeletePrinter} Else {. funcAddPrinter}}
                    }
                    Else
                    {
                        # Do the following if only one of these conditions is true, 1) The Computer is in the imediate OU, 2) A Not operator is used.
                        If (($ComputerOU -eq $PrinterRecord.Filters.FilterOrgUnit.name) -xor ($PrinterRecord.Filters.FilterOrgUnit.not -eq 1))
                        {If ($PrinterRecord.Properties.action -eq "D") {. funcDeletePrinter} Else {. funcAddPrinter}}
                    }
                }
            }
            Else
            {
                $Count = $PrinterRecord.Filters.FilterOrgUnit.Count
                $CompareWithNextRecord = 0
                $LastLoopCorrect = 1
                $SeriesCorrect = 0
                For ($i=0; $i -lt $Count; $i++)
                {
                    $LoopCorrect = 0
                    If ($PrinterRecord.Filters.FilterOrgUnit[$i].bool -eq "OR") {$CompareWithNextRecord = 0} Else {$CompareWithNextRecord = 1}
                    If ($i -eq 0) {$CompareWithNextRecord = 1}
                    If (($SeriesCorrect -eq 1) -and ($CompareWithNextRecord -eq 0))
                    {
                        If ($PrinterRecord.Properties.action -eq "D") {. funcDeletePrinter} Else {. funcAddPrinter}
                        $SeriesCorrect = 0
                        $LastLoopCorrect = 1
                    }
                    $CachedP = $PrinterRecord.Filters.FilterOrgUnit[$i].name
                    If (($PrinterRecord.Filters.FilterOrgUnit[$i].directMember -ne $NULL) -and ($PrinterRecord.Filters.FilterOrgUnit[$i].directMember -ne 0))
                    {
                        If ($PrinterRecord.Filters.FilterOrgUnit[$i].userContext -eq 1)
                        {
                            If (($UserOU.Contains($PrinterRecord.Filters.FilterOrgUnit[$i].name)) -xor ($PrinterRecord.Filters.FilterOrgUnit[$i].not -eq 1))
                            {$LoopCorrect = 1} Else {$LoopCorrect = 0}
                        }
                        Else
                        {
                            If (($ComputerOU.Contains($PrinterRecord.Filters.FilterOrgUnit[$i].name)) -xor ($PrinterRecord.Filters.FilterOrgUnit[$i].not -eq 1))
                            {$LoopCorrect = 1} Else {$LoopCorrect = 0}
                        }
                    }
                    Else
                    {
                        If ($PrinterRecord.Filters.FilterOrgUnit[$i].userContext -eq 1)
                        {
                            If (($UserOU -eq $PrinterRecord.Filters.FilterOrgUnit[$i].name) -xor ($PrinterRecord.Filters.FilterOrgUnit[$i].not -eq 1))
                            {$LoopCorrect = 1} Else {$LoopCorrect = 0}
                        }
                        Else
                        {
                            If (([string]$ComputerOU -eq [string]$PrinterRecord.Filters.FilterOrgUnit[$i].name) -xor ($PrinterRecord.Filters.FilterOrgUnit[$i].not -eq 1))
                            {$LoopCorrect = 1} Else {$LoopCorrect = 0}
                        }
                    }
                    If ($CompareWithNextRecord -eq 0) {$LastLoopCorrect = $LoopCorrect}
                    If (($LastLoopCorrect -eq 1) -and ($LoopCorrect -eq 1)) {$SeriesCorrect = 1}
                    If ($LoopCorrect -eq 0) {$LastLoopCorrect = 0}
                }
                If ($SeriesCorrect -eq 1)
                {
                    If ($PrinterRecord.Properties.action -eq "D") {. funcDeletePrinter} Else {. funcAddPrinter}
                }
                $CompareWithNextRecord = 0
                $LastLoopCorrect = 1
                $SeriesCorrect = 0
                $LoopCorrect = 0
            }
        }

        # Check through Group matches.
        If ($PrinterRecord.Filters.FilterGroup -ne $NULL)
        {
            # Check if there is only one Group to check, or multiple.
            If ($PrinterRecord.Filters.FilterGroup.Count -eq $NULL)
            {
                $CachedP = $PrinterRecord.Filters.FilterGroup.name
                $Pass = 0
                $Fail = 0
                # Check if we are comparing the User to a Group, or the Computer to a Group.
                If ($PrinterRecord.Filters.FilterGroup.userContext -eq 1) {$ListOfGroups = $UserGroups} Else {$ListOfGroups = $ComputerGroups}
                ForEach ($GroupRecord in $ListOfGroups)
                {
                    $GroupName = $GroupRecord.name
                    $GroupOldName = $GroupRecord.samaccountname
                    If ($PrinterRecord.Filters.FilterGroup.sid -ne $null)
                    {
                        $SID=(New-Object System.Security.Principal.SecurityIdentifier(([ADSISearcher]"(&(objectClass=Group)(cn=$GroupName))").FindOne().GetDirectoryEntry().ObjectSID.Value,0)).Value
                        # Do the following if only one of these conditions is true, 1) The object is in the Group, 2) A Not operator is used.
                        If (($SID -eq $PrinterRecord.Filters.FilterGroup.sid) -and ($PrinterRecord.Filters.FilterGroup.not -eq 1))
                        {
                            $Fail = 1
                        }
                        ElseIf ($SID -eq $PrinterRecord.Filters.FilterGroup.sid)
                        {
                            $Pass = 1
                        }
                    }
                    Else
                    {
                        # Do the following if only one of these conditions is true, 1) The object is in the Group, 2) A Not operator is used.
                        If (((("$ShortDomainName\$GroupName") -like $PrinterRecord.Filters.FilterGroup.samaccountname) -or (("$ShortDomainName\$GroupOldName") -like $PrinterRecord.Filters.FilterGroup.samaccountname)) -and ($PrinterRecord.Filters.FilterGroup.not -eq 1))
                        {
                            $Fail = 1
                        }
                        ElseIf ((("$ShortDomainName\$GroupName") -like $PrinterRecord.Filters.FilterGroup.samaccountname) -or (("$ShortDomainName\$GroupOldName") -like $PrinterRecord.Filters.FilterGroup.samaccountname))
                        {
                            $Pass = 1
                        }
                    }
                }
                If (($Pass -eq 1) -and ($Fail -eq 0)) {If ($PrinterRecord.Properties.action -eq "D") {. funcDeletePrinter} Else {. funcAddPrinter}}
            }
            Else
            {
                $Count = $PrinterRecord.Filters.FilterGroup.Count
                $CompareWithNextRecord = 0
                $LastLoopCorrect = 1
                $SeriesCorrect = 0
                For ($i=0; $i -lt $Count; $i++)
                {
                    $LoopCorrect = 0
                    If ($PrinterRecord.Filters.FilterGroup[$i].bool -eq "OR") {$CompareWithNextRecord = 0} Else {$CompareWithNextRecord = 1}
                    If ($i -eq 0) {$CompareWithNextRecord = 1}
                    If (($SeriesCorrect -eq 1) -and ($CompareWithNextRecord -eq 0))
                    {
                        If ($PrinterRecord.Properties.action -eq "D") {. funcDeletePrinter} Else {. funcAddPrinter}
                        $SeriesCorrect = 0
                        $LastLoopCorrect = 1
                    }
                    $CachedP = $PrinterRecord.Filters.FilterGroup[$i].name
                    $Pass = 0
                    $Fail = 0
                    # Check if we are comparing the User to a Group, or the Computer to a Group.
                    If ($PrinterRecord.Filters.FilterGroup[$i].userContext -eq 1) {$ListOfGroups = $UserGroups} Else {$ListOfGroups = $ComputerGroups}
                    ForEach ($GroupRecord in $ListOfGroups)
                    {
                        $GroupName = $GroupRecord.name
                        $GroupOldName = $GroupRecord.samaccountname
                        If ($PrinterRecord.Filters.FilterGroup[$i].sid -ne $null)
                        {
                            $SID=(New-Object System.Security.Principal.SecurityIdentifier(([ADSISearcher]"(&(objectClass=Group)(cn=$GroupName))").FindOne().GetDirectoryEntry().ObjectSID.Value,0)).Value
                            # Do the following if only one of these conditions is true, 1) The object is in the Group, 2) A Not operator is used.
                            If (($SID -eq $PrinterRecord.Filters.FilterGroup[$i].sid) -and ($PrinterRecord.Filters.FilterGroup[$i].not -eq 1))
                            {
                                $Fail = 1
                            }
                            ElseIf ($SID -eq $PrinterRecord.Filters.FilterGroup[$i].sid)
                            {
                                $Pass = 1
                            }
                        }
                        Else
                        {
                            # Do the following if only one of these conditions is true, 1) The object is in the Group, 2) A Not operator is used.
                            If (((("$ShortDomainName\$GroupName") -like $PrinterRecord.Filters.FilterGroup[$i].samaccountname) -or (("$ShortDomainName\$GroupOldName") -like $PrinterRecord.Filters.FilterGroup[$i].samaccountname)) -and ($PrinterRecord.Filters.FilterGroup[$i].not -eq 1))
                            {
                                $Fail = 1
                            }
                            ElseIf ((("$ShortDomainName\$GroupName") -like $PrinterRecord.Filters.FilterGroup[$i].samaccountname) -or (("$ShortDomainName\$GroupOldName") -like $PrinterRecord.Filters.FilterGroup[$i].samaccountname))
                            {
                                $Pass = 1
                            }
                        }
                    }
                    If (($Pass -eq 1) -and ($Fail -eq 0)) {$LoopCorrect = 1} Else {$LoopCorrect = 0}
                    If ($CompareWithNextRecord -eq 0) {$LastLoopCorrect = $LoopCorrect}
                    If (($LastLoopCorrect -eq 1) -and ($LoopCorrect -eq 1)) {$SeriesCorrect = 1}
                    If ($LoopCorrect -eq 0) {$LastLoopCorrect = 0}
                }
                If ($SeriesCorrect -eq 1)
                {
                    If ($PrinterRecord.Properties.action -eq "D") {. funcDeletePrinter} Else {. funcAddPrinter}
                }
                $CompareWithNextRecord = 0
                $LastLoopCorrect = 1
                $SeriesCorrect = 0
                $LoopCorrect = 0
            }
        }

        # Check through Computer Name matches.
        If ($PrinterRecord.Filters.FilterComputer -ne $NULL)
        {
            # Check if there is only one Computer Name to check, or multiple.
            If ($PrinterRecord.Filters.FilterComputer.Count -eq $NULL)
            {
                $CachedP = $PrinterRecord.Filters.FilterComputer.name
                # Check if name matches the computer name or fqdn.
                # Do the following if only one of these conditions is true, 1) The name matches the computer name or fqdn 2) A Not operator is used.
                If ((($env:COMPUTERNAME -like $PrinterRecord.Filters.FilterComputer.name) -or (($env:COMPUTERNAME+"."+$DomainName) -like $PrinterRecord.Filters.FilterComputer.name)) -xor ($PrinterRecord.Filters.FilterComputer.not -eq 1))
                {If ($PrinterRecord.Properties.action -eq "D") {. funcDeletePrinter} Else {. funcAddPrinter}}
            }
            Else
            {
                $Count = $PrinterRecord.Filters.FilterComputer.Count
                $CompareWithNextRecord = 0
                $LastLoopCorrect = 1
                $SeriesCorrect = 0
                For ($i=0; $i -lt $Count; $i++)
                {
                    $LoopCorrect = 0
                    If ($PrinterRecord.Filters.FilterComputer[$i].bool -eq "OR") {$CompareWithNextRecord = 0} Else {$CompareWithNextRecord = 1}
                    If ($i -eq 0) {$CompareWithNextRecord = 1}
                    If (($SeriesCorrect -eq 1) -and ($CompareWithNextRecord -eq 0))
                    {
                        If ($PrinterRecord.Properties.action -eq "D") {. funcDeletePrinter} Else {. funcAddPrinter}
                        $SeriesCorrect = 0
                        $LastLoopCorrect = 1
                    }
                    $CachedP = $PrinterRecord.Filters.FilterComputer[$i].name
                    If ((($env:COMPUTERNAME -like $PrinterRecord.Filters.FilterComputer[$i].name) -or (($env:COMPUTERNAME+"."+$DomainName) -like $PrinterRecord.Filters.FilterComputer[$i].name)) -xor ($PrinterRecord.Filters.FilterComputer[$i].not -eq 1))
                    {$LoopCorrect = 1} Else {$LoopCorrect = 0}
                    If ($CompareWithNextRecord -eq 0) {$LastLoopCorrect = $LoopCorrect}
                    If (($LastLoopCorrect -eq 1) -and ($LoopCorrect -eq 1)) {$SeriesCorrect = 1}
                    If ($LoopCorrect -eq 0) {$LastLoopCorrect = 0}
                }
                If ($SeriesCorrect -eq 1)
                {
                    If ($PrinterRecord.Properties.action -eq "D") {. funcDeletePrinter} Else {. funcAddPrinter}
                }
                $CompareWithNextRecord = 0
                $LastLoopCorrect = 1
                $SeriesCorrect = 0
                $LoopCorrect = 0
            }
        }

        # Check through User Name matches.
        If ($PrinterRecord.Filters.FilterUser -ne $NULL)
        {
            # Check if there is only one User to check, or multiple.
            If ($PrinterRecord.Filters.FilterUser.Count -eq $NULL)
            {
                $CachedP = $PrinterRecord.Filters.FilterUser.name
                # Check if a SID is meantioned, if so, match on SID, if not, match on name.
                If ($PrinterRecord.Filters.FilterUser.sid -ne $null)
                {
                    # Do the following if only one of these conditions is true, 1) The SID matches the currnet user, 2) A Not operator is used.
                    If (($PrinterRecord.Filters.FilterUser.sid -eq $UserSID) -xor ($PrinterRecord.Filters.FilterUser.not -eq 1))
                    {If ($PrinterRecord.Properties.action -eq "D") {. funcDeletePrinter} Else {. funcAddPrinter}}
                }
                Else
                {
                    # Do the following if only one of these conditions is true, 1) The username matches the currnet user, 2) A Not operator is used.
                    If ((("$ShortDomainName\$env:Username") -like $PrinterRecord.Filters.FilterUser.name) -xor ($PrinterRecord.Filters.FilterUser.not -eq 1))
                    {If ($PrinterRecord.Properties.action -eq "D") {. funcDeletePrinter} Else {. funcAddPrinter}}
                }
            }
            Else
            {
                $Count = $PrinterRecord.Filters.FilterUser.Count
                $CompareWithNextRecord = 0
                $LastLoopCorrect = 1
                $SeriesCorrect = 0
                For ($i=0; $i -lt $Count; $i++)
                {
                    $LoopCorrect = 0
                    If ($PrinterRecord.Filters.FilterUser[$i].bool -eq "OR") {$CompareWithNextRecord = 0} Else {$CompareWithNextRecord = 1}
                    If ($i -eq 0) {$CompareWithNextRecord = 1}
                    If (($SeriesCorrect -eq 1) -and ($CompareWithNextRecord -eq 0))
                    {
                        If ($PrinterRecord.Properties.action -eq "D") {. funcDeletePrinter} Else {. funcAddPrinter}
                        $SeriesCorrect = 0
                        $LastLoopCorrect = 1
                    }
                    $CachedP = $PrinterRecord.Filters.FilterUser[$i].name
                    If ($PrinterRecord.Filters.FilterUser[$i].sid -ne $null)
                    {
                        # Do the following if only one of these conditions is true, 1) The SID matches the currnet user, 2) A Not operator is used.
                        If (($PrinterRecord.Filters.FilterUser[$i].sid -eq $UserSID) -xor ($PrinterRecord.Filters.FilterUser[$i].not -eq 1))
                        {$LoopCorrect = 1} Else {$LoopCorrect = 0}
                    }
                    Else
                    {
                        # Do the following if only one of these conditions is true, 1) The username matches the currnet user, 2) A Not operator is used.
                        If ((("$ShortDomainName\$env:Username") -like $PrinterRecord.Filters.FilterUser[$i].name) -xor ($PrinterRecord.Filters.FilterUser[$i].not -eq 1))
                        {$LoopCorrect = 1} Else {$LoopCorrect = 0}
                    }
                    If ($CompareWithNextRecord -eq 0) {$LastLoopCorrect = $LoopCorrect}
                    If (($LastLoopCorrect -eq 1) -and ($LoopCorrect -eq 1)) {$SeriesCorrect = 1}
                    If ($LoopCorrect -eq 0) {$LastLoopCorrect = 0}
                }
                If ($SeriesCorrect -eq 1)
                {
                    If ($PrinterRecord.Properties.action -eq "D") {. funcDeletePrinter} Else {. funcAddPrinter}
                }
                $CompareWithNextRecord = 0
                $LastLoopCorrect = 1
                $SeriesCorrect = 0
                $LoopCorrect = 0
            }
        }
    }
}
