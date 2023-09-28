#Load SharePoint CSOM Assemblies
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"

#Config Parameters
$SiteURL = "https://sharepoint-site-url.com"
$CSVPath = ".\export.csv"
    
#Function to get all files of a folder
Function Get-FilesFromFolder([Microsoft.SharePoint.Client.Folder]$Folder) {

  $ListItemCollection = New-Object System.Collections.ArrayList
 
  #Get All Files of the Folder
  $Ctx.load($Folder.files)
  $Ctx.ExecuteQuery()

  #list all files in Folder
  ForEach ($File in $Folder.files) {
    #Get the File Name or do something
    [void]$ListItemCollection.Add([PSCustomObject]@{
        DocumentTitle   = $File.Name
        LastModified = $List.LastItemModifiedDate
        Path = $Folder.ServerRelativeUrl
      })
  }

  $ListItemCollection | Export-Csv $CSVPath -NoTypeInformation -Append
   
  #Recursively Call the function to get files of all folders
  $Ctx.load($Folder.Folders)
  $Ctx.ExecuteQuery()
 
  #Exclude "Forms" system folder and iterate through each folder
  ForEach ($SubFolder in $Folder.Folders | Where-Object { $_.Name -ne "Forms" }) {
    Get-FilesFromFolder -Folder $SubFolder
  }
}
 
#powershell list files in sharepoint online library
Function Get-SPODocLibraryFiles() {
  param
  (
    [Parameter(Mandatory = $true)] [string] $SiteURL,
    [Parameter(Mandatory = $true)] [string] $LibraryName
  )

  Try {
    #Get the Library and Its Root Folder
    $Library = $Ctx.web.Lists.GetByTitle($LibraryName)
    $Ctx.Load($Library)
    $Ctx.Load($Library.RootFolder)
    $Ctx.ExecuteQuery()
 
    #Call the function to get Files of the Root Folder
    Get-FilesFromFolder -Folder $Library.RootFolder
 
  }
  Catch {
    Write-Host -f Red "Error Getting Files from Library!" $_.Exception.Message
  }
}


Try {

  # Get Credentials
  $cred = Get-Credential

  #Setup the context
  $Ctx = New-Object Microsoft.SharePoint.Client.ClientContext($SiteURL)
  $Ctx.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Cred.Username, $Cred.Password)
          
  #Get all lists from the Web
  $Lists = $Ctx.Web.Lists
  $Ctx.Load($Lists)
  $Ctx.ExecuteQuery()
  $ListItemCollection = New-Object System.Collections.ArrayList
   
 
  #Iterate through Lists
  ForEach ($List in $Lists | Where-Object { $_.hidden -eq $false }) {

    Get-SPODocLibraryFiles -SiteURL $SiteURL -LibraryName $List.Title

  }
}
Catch {
  Write-Host -f Red "Error:" $_.Exception.Message
}
 
