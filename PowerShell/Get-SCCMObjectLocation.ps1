Function Get-SCCMObjectLocation {
<#
.Synopsis
   Searches for an SCCM object by ID and displays its path
.DESCRIPTION
   Finds an SCCM object by its SCCM object ID (Collection ID, Package ID, etc), and then displays its position in the console
   folder hierarchy.
.PARAMETER SMSId
    An SCCM ID in the standard [3 letter site code][5 hex digits] format, eg SMS00001
.PARAMETER SiteCode
    A three letter SCCM site code, eg "ABC"   (Optional)
.PARAMETER SiteServerName
    A resolvable name for an SCCM site server, eg "sccm01.company.com"   (Optional)
.EXAMPLE
    Get-SCCMObjectLocation -SMSId "ABC00166"
root\Application Deployment\MS Access App-V 	[SMS_Collection_User]
.EXAMPLE
   Get-SCCMObjectLocation -SMSId "ABC001BB"
root\Dell PowerEdge Drivers OM7.3.0\PE1950-Microsoft Windows 2008 R2 SP1-OM7.3 	[SMS_DriverPackage]
.EXAMPLE
    Get-SCCMObjectLocation -SMSId ABC000CC

    WARNING: Multiple objects with ID: LOL000CC
    root\x64\Win 7 Ent x64 with Office 2010 	[SMS_ImagePackage]
    root\Server roles\All Domain Controllers 	[SMS_Collection_Device]
.NOTES
   Needs the SCCM Console (version 2012 SP1 or later) installed on the machine.
   Needs to be run from the CMSite PSDrive (which should be automatically created if you run this from the console's "Connect via 
   Windows PowerShell" menu option, or if you load the ConfigMgr PowerShell module, via:
   Import-Module (Join-Path -Path (Split-Path -Parent $ENV:SMS_ADMIN_UI_PATH)\ConfigurationManager.psd1)
   If you don't know it, you can change to the PSDrive using:
   Set-Location -LiteralPath "$((Get-PSDrive -PSProvider CMSite).Name):"
   
   Written by James Blatchford, July 2016
   https://github.com/GAThrawnMIA/Work-Scripts/blob/master/PowerShell/Get-SCCMObjectLocation.ps1
   
   Basic idea (and pointer to SMS_ObjectContainerNode and SMS_ObjectContainerItem) from Peter van der Woude's blog entry:
   https://www.petervanderwoude.nl/post/get-the-folder-location-of-an-object-in-configmgr-2012-via-powershell/
.LINK
    ConfigurationManager
.LINK
    about_ConfigurationManager_Cmdlets
.LINK
    https://github.com/GAThrawnMIA/Work-Scripts/blob/master/PowerShell/Get-SCCMObjectLocation.ps1
.LINK
    http://gathrawn.jard.co.uk/2016/07/find-objects-within-folders-in-sccm.html
#>
    param(
        [Parameter(Mandatory=$true,
                ValueFromPipeline=$true,
                ValueFromPipelineByPropertyName=$true, 
                Position=0)][string]$SMSId,
        [string]$SiteCode = (Get-CMSite).SiteCode,
        [string]$SiteServerName = (Get-CMSite).ServerName)

    #Find the container directly containing the item
    $ContainerItem = Get-WmiObject -Namespace root/SMS/site_$($SiteCode) -ComputerName $SiteServerName -Query "select * from SMS_ObjectContainerItem where InstanceKey = '$($SMSId)'"
    If (!$ContainerItem) {
        $ObjectName = Get-WmiObject -Namespace root/SMS/site_$($SiteCode) -ComputerName $SiteServerName -Query "select * from SMS_ObjectName where ObjectKey = '$($SMSId)'"
        If (!$ObjectName) {
            Write-Warning "No object or containers found for $SMSId"
            break;
        }
        Else
        {
            If ($ObjectName -is [array]) {
                Write-Warning "Multiple objects with ID: $SMSId"
                Foreach ($Object in $ObjectName) {
                    $ObjectOutputString = "$ObjectOutputString`nroot\$(($Object).Name)"
                    $Object
                }
                Return $ObjectOutputString
            }
            Else {
                Return "root\$(($ObjectName).Name)"
            }
            break;
        }
    }

    If ($ContainerItem -is [array]) {
        Write-Warning "Multiple objects with ID: $SMSId"
        Foreach ($Item In $ContainerItem) {
            $tempOutputString = Get-SCCMContainerHierarchy -ContainerNodeId $Item.ContainerNodeID -ObjectType $Item.ObjectType -ObjectTypename $Item.ObjectTypeName -SiteCode $SiteCode -SiteServerName $SiteServerName
            $OutputString = "$OutputString`nroot\$tempOutputString"
        }
        Return "$OutputString"
    }
    Else {
        #One object found
        $OutputString = Get-SCCMContainerHierarchy -ContainerNodeId ($ContainerItem).ContainerNodeID -SiteCode $SiteCode -SiteServerName $SiteServerName
        Return "root\$OutputString"
    }
    
    
}

Function Get-SCCMContainerHierarchy {
    param(
        [Parameter(Mandatory=$true,
                Position=0)][string]$ContainerNodeId,
        [Parameter(Mandatory=$true,
                Position=1)][string]$SiteCode = (Get-CMSite).SiteCode,
        [Parameter(Mandatory=$true,
                Position=2)][string]$SiteServerName = (Get-CMSite).ServerName,
        [Parameter(Mandatory=$false)]$ObjectType = ($ContainerItem).ObjectType,
        [Parameter(Mandatory=$false)]$ObjectTypeName = ($ContainerItem).ObjectTypeName)
    
    Switch ($ObjectType) {
        2       {$ObjectTypeText = $ObjectTypeName; $ObjectName = (Get-CMPackage -ID $SMSId).Name} # Package
        7       {$ObjectTypeText = $ObjectTypeName; $ObjectName = (Get-CMQuery -ID $SMSId).Name} # Query
        14      {$ObjectTypeText = $ObjectTypeName; $ObjectName = (Get-CMOperatingSystemInstaller -ID $SMSId).Name} # OS Install Package
        18      {$ObjectTypeText = $ObjectTypeName; $ObjectName = (Get-CMOperatingSystemImage -ID $SMSId).Name} # OS Image
        20      {$ObjectTypeText = $ObjectTypeName; $ObjectName = (Get-CMTaskSequence -ID $SMSId).Name} # Task Sequence
        23      {$ObjectTypeText = $ObjectTypeName; $ObjectName = (Get-CMDriverPackage -ID $SMSId).Name} # Driver Package
        19      {$ObjectTypeText = $ObjectTypeName; $ObjectName = (Get-CMBootImage -ID $SMSId).Name} # Boot Image
        5000    {$ObjectTypeText = $ObjectTypeName; $ObjectName = (Get-CMDeviceCollection -Id $SMSId).Name} # Device Collection
        5001    {$ObjectTypeText = $ObjectTypeName; $ObjectName = (Get-CMUserCollection -Id $SMSId).Name} # User Collection
        default {$ObjectTypeText = "unknown object type: '$($ObjectTypeName)' = $($ObjectType)"; $ObjectName = "unknown object name ($SMSId)"}
    }

    $OutputString = "$ObjectName `t[$ObjectTypeText]"
    #ContainerNodeID of 0 is the root
    While ($ContainerNodeId -ne 0) {
        #Find details of that container
        $ContainerNode = Get-WmiObject -Namespace root/SMS/site_$($SiteCode) -ComputerName $SiteServerName -Query "select * from SMS_ObjectContainerNode where ContainerNodeID = '$($ContainerNodeId)'"
        $ContainerName = ($ContainerNode).Name
        $ContainerNodeId = ($ContainerNode).ParentContainerNodeID
        $OutputString = "$ContainerName\$OutputString"
    }
    Return $OutputString
}
