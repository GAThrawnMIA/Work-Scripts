Function Get-SCCMObjectLocation {
<#
.Synopsis
   Searches for an SCCM object by ID and displays its path
.DESCRIPTION
   Finds an SCCM object by its SCCM object ID (Collection ID, Package ID, etc), and then displays its position in the console folder hierarchy
.PARAMETER SMSId
    An SCCM ID in the standard [3 letter site code][5 hex digits] format

    eg SMS00001
.EXAMPLE
   Get-SCCMObjectLocation -SMSId "LOL001BB"
root\Live Legacy\Lync 2013 Fixes\[Package] Repair Send To Outlook post O2k13 MUI install
.EXAMPLE
    Get-SCCMObjectLocation -SMSId "LOL00450"
root\Test\Test - James\[Device Collection] Adobe LiveCycle Servers
.NOTES
   Needs SCCM Console installed on the machine
   James B
   Basic idea (and pointer to SMS_ObjectContainerNode and SMS_ObjectContainerItem) from Peter van der Woude's blog entry:
   https://www.petervanderwoude.nl/post/get-the-folder-location-of-an-object-in-configmgr-2012-via-powershell/
.LINK
    My link
#>
    param(
        [Parameter(Mandatory=$true,
                ValueFromPipeline=$true,
                ValueFromPipelineByPropertyName=$true, 
                ValueFromRemainingArguments=$false, 
                Position=0)][string]$SMSId,
        [string]$SiteCode = (Get-CMSite).SiteCode,
        [string]$SiteServerName = (Get-CMSite).ServerName)

    #Find the container directly containing the item
    $ContainerItem = Get-WmiObject -Namespace root/SMS/site_$($SiteCode) -ComputerName $SiteServerName -Query "select * from SMS_ObjectContainerItem where InstanceKey = '$($SMSId)'"
    If ([string]::IsNullOrEmpty($ContainerItem)) {
        "no object containers found for $SMSId"
        break;
    }
    #($ContainerItem).ObjectType
    #($ContainerItem).ObjectTypeName
    #($ContainerItem).ContainerNodeID
    $ContainerNodeId = ($ContainerItem).ContainerNodeID

    <# Object Types
    2 = SMS_Package
    5000 = SMS_Collection_Device
    #>
    Switch (($ContainerItem).ObjectType) {
        2 {$ObjectType = "Package"; $ObjectName = (Get-CMPackage -ID $SMSId).Name}
        5000 {$ObjectType = "Device Collection"; $ObjectName = (Get-CMDeviceCollection -Id $SMSId).Name}
        default {$ObjectType = "unknown object type: $(($ContainerItem).ObjectType)"; $ObjectName = "unknown object name ($SMSId)"}
    }

    $OutputString = "[$ObjectType] $ObjectName"
    $ContainerNodeId
    #If ($ContainerNodeId -eq "") {break;}
    #break;
    #ContainerNodeID of 0 is the root
    While ($ContainerNodeId -ne 0) {
        #Find details of that container
        $ContainerNode = Get-WmiObject -Namespace root/SMS/site_$($SiteCode) -ComputerName $SiteServerName -Query "select * from SMS_ObjectContainerNode where ContainerNodeID = '$($ContainerNodeId)'"
        #($ContainerNode).Name
        $ContainerName = ($ContainerNode).Name
        #($ContainerNode).ParentContainerNodeID
        $ContainerNodeId = ($ContainerNode).ParentContainerNodeID
        $OutputString = "$ContainerName\$OutputString"
    }
    "root\$OutputString"
}