#More info https://gathrawn.jard.co.uk/2019/04/retrieving-package-source.html

$CSVOutPath = 'c:\temp\AppSrc.csv'

#Change to your site's ConfigMgr PSDrive before running this script

#Get each Application, then get the source path from each Applications's Deployment Types
Write-Host "Fetching application details - this takes a few minutes, please wait...`n"

$Applications = Get-CMApplication
$AppCount = $Applications.Count
Write-Host "$AppCount applications found.`n"

#Iterate over the apps list pulling out the details for each app. Takes a couple of minutes.
$i = 1
$AppSourceList = ForEach ($App in $Applications)
{
    Write-Progress -Activity 'Checking apps and deployment types' -Id 1 -PercentComplete $(($i / $AppCount) * 100) -CurrentOperation "app $i / $AppCount"
    $PackageXml = [xml]$App.SDMPackageXML
    #An app can have multiple Deployment Types, each with their own source location. DT details are stored in the XML properties
    ForEach($DT in $PackageXml.AppMgmtDigest.DeploymentType) {
        $DtTitle = $DT.Title.'#text'    #need to quote property names with hashes in them, normal backtick escaping doesn't work
        $DtTech = $DT.Technology
        $DtLocation = $DT.Installer.Contents.Content.Location
        New-Object -TypeName psobject -Property (@{AppDisplayName = $App.LocalizedDisplayName; PackageID = $App.PackageID;
        CiId = $app.CI_ID; Enabled = $app.IsEnabled; Superseded = $app.IsSuperseded; HasContent = $app.HasContent;
        DepTitle = $DtTitle; DepTypeTech = $DtTech; DepTypeSrcLocation = $DtLocation})
    }
    $i++
}

#$AppSourceList | Out-GridView
$AppSourceList | Export-Csv -Path $CSVOutPath -NoTypeInformation
