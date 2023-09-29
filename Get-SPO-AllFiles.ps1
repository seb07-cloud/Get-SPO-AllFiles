param (
  [Parameter(Mandatory = $false)]
  [switch]$csvExport,

  [Parameter(Mandatory = $false)]
  [switch]$csvImport,

  [Parameter(Mandatory = $false,
    HelpMessage = "Full Path to the CSV, including the Name of the File")]
  [ValidateNotNullOrEmpty()]
  [string]$csvName
)

#Load SharePoint CSOM Assemblies
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"

#Config Parameters
$AdminSite = "https://xxxxxxxxxxx-admin.sharepoint.com"

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
  
  $ListItemCollection | Export-Csv $newFullCsvPath -NoTypeInformation -Append
   
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
    (Get-SPOSite).Url | Select-Object @{Name = 'Url'; Expression = { $_ } } | Export-Csv $csvName -NoTypeInformation
  }
  catch {
    Write-Error "Error connecting to Admin Site!" $_.Exception.Message
  }
}


Try {

  # Get Credentials
  $cred = Get-Credential

  if ($csvExport) {

    if (Test-Path -Path $csvName) {
      Remove-Item -Path $csvName
    }

    Export-AllSPOSites -AdminSite $AdminSite
    break
  }
  else {
    try {
      $sites = Import-Csv -Path $csvName -Delimiter ";"

      $fullCsvPath = Resolve-Path $csvName
      $directory = [System.IO.Path]::GetDirectoryName($fullCsvPath)
      $directory = Split-Path -Parent $fullCsvPath
      $newCsvName = 'SpFileFolderExport.csv'
      $newFullCsvPath = Join-Path -Path $directory -ChildPath $newCsvName 

      if (Test-Path -Path $newFullCsvPath) {
        Remove-Item -Path $newFullCsvPath
      }
    }
    catch {
      Write-Error "Couldn't import CSV from $($csvName) - File not found!"
      break
    }

    foreach ($site in $sites) {

      Write-Host $site.Url

      #Setup the context
      $Ctx = New-Object Microsoft.SharePoint.Client.ClientContext($site.Url)
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
}
Catch {
  Write-Host -f Red "Error:" $_.Exception.Message
}
 
