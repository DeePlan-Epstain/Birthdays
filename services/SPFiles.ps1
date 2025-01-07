#Config Variables
$SiteURL = "https://projects1.sharepoint.com"
$FolderSiteRelativeUrl = "/sites/portal/training/DevelopmentTeamDocs" #Folder's Site Relative Path
$UserName = "roey-deeplan"
$Password = "ruthvani123!"
$SecurePassword = ConvertTo-SecureString -String $Password -AsPlainText -Force
$Cred = New-Object -TypeName System.Management.Automation.PSCredential -argumentlist $UserName, $SecurePassword
  
Try {
  #Connect to PnP Online
  Connect-PnPOnline -Url $SiteURL -WarningAction Ignore -Credentials $Cred
      
  #Get All Items from Folder in Batch
  $ListItems = Get-PnPListItem -List $ListName -PageSize 2000 | Sort-Object ID -Descending
     
  #Get List Items from the folder
  $ItemsFromFolder = $ListItems | Where {$_.FieldValues["FileDirRef"] -match $FolderSiteRelativeUrl }
   
  Write-host "Total Number of Items in the Library:"$ListItems.count
  Write-host "Total Number of Items in the Folder:"$ItemsFromFolder.count
   
  #Powershell to delete all files from a folder
  ForEach ($Item in $ItemsFromFolder)
  {
      #Remove-PnPListItem -List $ListName -Identity $Item.Id -Recycle -Force | Out-Null
      Write-host -f DarkYellow "Removed File at"$Item.FieldValues["FileRef"]
  }
}
Catch {
  write-host "Error: $($_.Exception.Message)" -foregroundcolor Red
}