#Load SharePoint CSOM Assemblies
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"

#Config Parameters
$AdminSite = "https://admin-site-url.com"
$CSVPath = ".\export.csv"

    
#Function to get all files of a folder
Function Get-FilesFromFolder([Microsoft.SharePoint.Client.Folder]$Folder) {

  $ListItemCollection = New-Object System.Collections.ArrayList
 
  #Get All Files of the Folder
  $Ctx.load($Folder.files)
  $Ctx.ExecuteQuery()

  #list all files in Folder
  ForEach ($File in $Folder.files) {

    $extension = [System.IO.Path]::GetExtension($File.Name)

    [void]$ListItemCollection.Add([PSCustomObject]@{
        DocumentTitle = $File.Name
        LastModified  = $File.TimeLastModified
        Path          = $Folder.ServerRelativeUrl
        FileSizeInMb  = [Math]::Round($File.Length / 1048576, 2)
        FileExtension = $extension
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

function Export-AllSPOSites {
  param (
    [Parameter(Mandatory = $true)]  [string] $AdminSite
  )
  
  try {
    Connect-SPOService -Url $AdminSite -Credential $cred  
    $sites = Get-SPOSite
  }
  catch {
    Write-Error "Error connecting to Admin Site!" $_.Exception.Message
  }

  $sites.Url
}


Try {

  # Get Credentials
  $cred = Get-Credential
  $sites = Export-AllSPOSites -AdminSite $AdminSite

  foreach ($site in $sites) {


    #Setup the context
    $Ctx = New-Object Microsoft.SharePoint.Client.ClientContext($site)
    $Ctx.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Cred.Username, $Cred.Password)
          
    #Get all lists from the Web
    $Lists = $Ctx.Web.Lists
    $Ctx.Load($Lists)
    $Ctx.ExecuteQuery()
    $ListItemCollection = New-Object System.Collections.ArrayList
   
 
    #Iterate through Lists
    ForEach ($List in $Lists | Where-Object { $_.hidden -eq $false }) {

      Get-SPODocLibraryFiles -SiteURL $site -LibraryName $List.Title

    }
  }
}
Catch {
  Write-Host -f Red "Error:" $_.Exception.Message
}
 
