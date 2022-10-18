# Parameter for adding LPW to tenant
Param(
[Parameter()][string]$TenantName,
[Parameter()][string]$AdminEmail
)

#connection for m365lp configuration/site creation and cdnconfig and admin page creation
$clSite = "https://$TenantName.sharepoint.com/sites/M365LP"
$clSiteconnection = Connect-PnPOnline -Url $clSite -Interactive -ReturnConnection 

#connection for m365lp configuration/site creation and cdnconfig and admin page creation
$siteCreationUrl = "https://$TenantName.sharepoint.com/"
$siteCreationConnection = Connect-PnPOnline -Url $siteCreationUrl -Interactive -ReturnConnection -TenantAdminUrl $AppCatalogURL

# App catalog URL
$AppCatalogURL = "https://$TenantName.sharepoint.com/sites/appcatalog"
$AppCatalogURLConnection = Connect-PnPOnline -Url $AppCatalogURL -Interactive -ReturnConnection -TenantAdminUrl $AppCatalogURL

#  CREATE APP CATALOG
Register-PnPAppCatalogSite -Url $AppCatalogURL -Owner $AdminEmail -TimeZoneId 4 -Connection $AppCatalogURLConnection

# Create m365LP site
New-PnPSite -Type CommunicationSite -Title M365LearningPathways -Url https://$TenantName.sharepoint.com/sites/M365LP -Lcid 1053 -Connection $siteCreationConnection

# #UPLOAD SPFx APP TO SHAREPOINT APP CATALOG
$AppFilePath = "C:\Repos\custom-learning-office-365\src\customlearning.sppkg"
#Connect-PnPOnline -Url $AppCatalogURL -Interactive 
#Add App to App catalog

#waiting for app catalog to be created
#do {  
  #$appcatalogPresent = Get-PnPTenantAppCatalogUrl -Connection $addAppcatalogConnection
  Start-Sleep -Seconds 120
  $isNotProvisioned = Get-PnPTenantAppCatalogUrl -Connection $siteCreationConnection
  if ($null -eq $isNotProvisioned) {
    Write-Host "Waiting 120 secs for creation of app catalog"  
    Start-Sleep -Seconds 120
  }
 
#} while ($null -eq $appcatalogPresent)

#Deploy App to the Tenant
Add-PnPApp -Path $AppFilePath -Scope Tenant -Publish -Connection $AppCatalogURLConnection

#Publish-PnPApp -Identity $App.Id

#UPDATE LEARNING PATWAYS WITH CDN AND ADMIN PAGES
try {
  Set-PnPStorageEntity -Key MicrosoftCustomLearningCdn -Value "https://yonas101.github.io/custom-learning-office-365/learningpathways/" -Description "CDN source for Microsoft 365 learning pathways Content" -Connection $clSiteconnection
  Get-PnPStorageEntity -Key MicrosoftCustomLearningCdn -Connection $clSiteconnection
  Set-PnPStorageEntity -Key MicrosoftCustomLearningSite -Value $clSite -Description "M365 learning pathways Site Collection" -Connection $clSiteconnection
  Get-PnPStorageEntity -Key MicrosoftCustomLearningSite -Connection $clSiteconnection
  Set-PnPStorageEntity -Key MicrosoftCustomLearningTelemetryOn -Value $true -Description "M365 learning pathways Telemetry Collection" -Connection $clSiteconnection
  Get-PnPStorageEntity -Key MicrosoftCustomLearningTelemetryOn -Connection $clSiteconnection

  $clv = Get-PnPListItem -List "SitePages" -Query "<View><Query><Where><Eq><FieldRef Name='FileLeafRef'/><Value Type='Text'>CustomLearningViewer.aspx</Value></Eq></Where></Query></View>" -Connection $clSiteconnection
  if ($null -eq $clv) {
    $clvPage = Add-PnPPage -Name "CustomLearningViewer" -Connection $clSiteconnection
    Add-PnPPageSection -Page $clvPage -SectionTemplate OneColumn -Order 1 -Connection $clSiteconnection
    Add-PnPPageWebPart -Page $clvPage -Component "Microsoft 365 learning pathways" -Connection $clSiteconnection
    Set-PnPPage -Identity $clvPage -Publish -Connection $clSiteconnection
    $clv = Get-PnPListItem -List "SitePages" -Query "<View><Query><Where><Eq><FieldRef Name='FileLeafRef'/><Value Type='Text'>CustomLearningViewer.aspx</Value></Eq></Where></Query></View>" -Connection $clSiteconnection
  }
  $clv["PageLayoutType"] = "SingleWebPartAppPage"
  $clv.Update()
  Invoke-PnPQuery -Connection $clSiteconnection
    
  $cla = Get-PnPListItem -List "SitePages" -Query "<View><Query><Where><Eq><FieldRef Name='FileLeafRef'/><Value Type='Text'>CustomLearningAdmin.aspx</Value></Eq></Where></Query></View>"  -Connection $clSiteconnection
  if ($null -eq $cla) {
    $claPage = Add-PnPPage "CustomLearningAdmin" -Publish -Connection $clSiteconnection
    Add-PnPPageSection -Page $claPage -SectionTemplate OneColumn -Order 1 -Connection $clSiteconnection
    Add-PnPPageWebPart -Page $claPage -Component "Microsoft 365 learning pathways administration" -Connection $clSiteconnection
    Set-PnPPage -Identity $claPage -Publish -Connection $clSiteconnection
    $cla = Get-PnPListItem -List "SitePages" -Query "<View><Query><Where><Eq><FieldRef Name='FileLeafRef'/><Value Type='Text'>CustomLearningAdmin.aspx</Value></Eq></Where></Query></View>" -Connection $clSiteconnection
  }
  $cla["PageLayoutType"] = "SingleWebPartAppPage"
  $cla.Update()
  Invoke-PnPQuery -Connection $clSiteconnection
    
}
catch {
  Write-Error "Failed to authenticate to $siteUrl"
  Write-Error $_.Exception
}

Disconnect-PnPOnline

