
Function Get-SCCMCollections {
<#
.Synopsis
   Show all direct and indirect collection memberships for an object
.DESCRIPTION
   Find all direct and indirect collection memberships for a user or computer object in SCCM
.PARAMETER noListMultiple
    When searching for an object by its SMS ResourceName, by default this cmdlet will output a warning if it detects multiple objects
    with the same name and list them all by name and ID. This will cause problems if this cmdlet is not the last cmdlet in a
    pipeline. set -noListMultiple to suppress separately listing the details of duplicates.
.EXAMPLE
   Get-SCCMCollections -type Computer -resourceName server01
.EXAMPLE
   Get-SCCMCollections -type User -resourceName bloggsf -noListMultiple | Out-GridView
.NOTES
   Needs the SCCM Console (version 2012 SP1 or later) installed on the machine.
   Needs to be run from the CMSite PSDrive (which should be automatically created if you run this from the console's "Connect via 
   Windows PowerShell" menu option, or if you load the ConfigMgr PowerShell module, via:
   Import-Module (Join-Path -Path (Split-Path -Parent $ENV:SMS_ADMIN_UI_PATH)\ConfigurationManager.psd1)
   If you don't know it, you can change to the PSDrive using:
   Set-Location -LiteralPath "$((Get-PSDrive -PSProvider CMSite).Name):"

   Written by James Blatchford, February 2017
.LINK
   https://github.com/GAThrawnMIA/Work-Scripts/blob/master/PowerShell/Get-SCCMCollections.ps1
#>
    param(
    [Parameter(Mandatory=$true,Position=1)]
    [ValidateSet("User", "Computer")]
    [string]$type,

    [Parameter(Mandatory=$true,Position=2,ParameterSetName='Name')]
    [string]$resourceName,

    [Parameter(Mandatory=$true,Position=3,ParameterSetName='Id')]
    [string]$resourceId,

    [Parameter(Mandatory=$false,ParameterSetName='Name')]
    [switch]$noListMultiple,

    [Parameter()]
    [string]$sccmServerName = "sccm",

    [Parameter()]
    [string]$sccmSiteCode = (Get-PSDrive -PSProvider CMSite).Name
    )

    If ($type -eq "User") {
        If ($PSCmdlet.ParameterSetName -eq "Id") {
            $Users = Get-CMResource -ResourceId $resourceId -ResourceType User -Fast | Select-Object -Property ResourceId,FullUserName,Name
        }
        ElseIf ($PSCmdlet.ParameterSetName -eq "Name") {
            $Users = Get-CMUser -Name "*$resourceName" | Select-Object -Property ResourceID,Name
            #Allow for multiple users matching the same wildcarded user ID (eg different domains, prefixed letters)
            If ( ($Users | Measure-Object).Count -gt 1) {
                Write-Warning "Multiple users found with name: $resourceName. Check ResourceIDs in output to differentiate them."
                If ($noListMultiple -eq $false) {
                    $Users | ft -AutoSize
                }
            }
            
        }
        $Users | Foreach {
            $ResourceID = $_.ResourceID
            $Collections = Get-WmiObject -ComputerName $sccmServerName -Namespace "root\sms\site_$sccmSiteCode" -Class "SMS_FullCollectionMembership" -Filter "ResourceId = '$($_.ResourceId)'" | Select-Object -Property CollectionID,IsDirect,Name
            $Collections | Foreach {
                $CollectionName = Get-CMUserCollection -Id $_.CollectionID | Select-Object -Property CollectionID,Name
                New-Object -TypeName psobject -Property (@{"ResourceName" = $_.Name; "ResourceId" = $ResourceID; "CollectionName" = $CollectionName.Name; "CollectionID" = $_.CollectionID; "IsDirect" = $_.IsDirect})
            }
        }
    }
    ElseIf ($type -eq "Computer") {
        If ($PSCmdlet.ParameterSetName -eq "Id") {
            $Devices = Get-CMResource -ResourceId $ResourceId -ResourceType System -Fast
        }
        ElseIf ($PSCmdlet.ParameterSetName -eq "Name") {
            $Devices = Get-CMDevice -Name $resourceName | Select-Object -Property ResourceId,Name,LastActiveTime
            #Allow for Duplicate device names with different IDs (rebuilds, etc)
            If ( ($Devices | Measure-Object).Count -gt 1) {
                Write-Warning "Multiple devices found with name like: $resourceName. Check ResourceIDs in output to differentiate them."
                If ($noListMultiple -eq $false) {
                    $Devices | ft -AutoSize
                }
            }
        }
        $Devices | Foreach {
            $ResourceID = $_.ResourceID
            $Collections = Get-WmiObject -ComputerName $sccmServerName -Namespace "root\sms\site_$sccmSiteCode" -Class "SMS_FullCollectionMembership" -Filter "ResourceId = '$($_.ResourceId)'" | Select-Object -Property CollectionID,IsDirect,Name
            $Collections | Foreach {
                $CollectionName = Get-CMDeviceCollection -Id $_.CollectionID | Select-Object -Property CollectionID,Name
                New-Object -TypeName psobject -Property (@{"ResourceName" = $_.Name; "ResourceId" = $ResourceID; "CollectionName" = $CollectionName.Name; "CollectionID" = $_.CollectionID; "IsDirect" = $_.IsDirect})
            }
        }
    }
    Else {
        Write-Error "Unknown object type"
    }
}
