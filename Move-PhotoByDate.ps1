#requires -version 4
<#
.SYNOPSIS
  Script to organize and move photos into directories by date.
.DESCRIPTION
  This PowerShell script organizes photos by their date of capture. It moves each photo into a folder structure based on the year and month of the photo's creation date. If the capture date is unavailable, the script uses the file's last write time.
.PARAMETER PicturePath
  The path to the folder containing the photos to be organized.
.PARAMETER TargetPath
  The path to the folder where the photos should be moved. The script will create subfolders for each year and month.
.INPUTS
  String paths for PicturePath and TargetPath.
.NOTES
  Version:        1.0
  Author:         MrManjah
  Creation Date:  14/10/2024
  Purpose/Change: Initial script development
.EXAMPLE
  Move-Pictures -PicturePath "\\DiskStation\Photo\Amélie a trier" -TargetPath "\\DiskStation\Photo\"
  
  This example moves photos from the source folder to the target folder, organizing them by year and month.
#>

#---------------------------------------------------------[Script Parameters]------------------------------------------------------

Param (
  [String]
  # Path to the folder containing photos to be organized
  $PicturePath = "",
  
  [String]
  # Target folder where photos should be moved
  $TargetPath = ""
)

#---------------------------------------------------------[Initialisations]--------------------------------------------------------

# Set Error Action to Silently Continue
$ErrorActionPreference = 'SilentlyContinue'

# Import Modules & Snap-ins
Import-Module PSLogging

#----------------------------------------------------------[Declarations]----------------------------------------------------------

# Script Version
$ScriptVersion = '1.0'

#---------------------------------------------------------[Functions and Logic]----------------------------------------------------

# Function to get the date a photo was taken
function Get-DateTaken {
  param (
    [Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
    [Alias('FullName')]
    [String]
    $Path
  )
    
  begin {
    $shell = New-Object -COMObject Shell.Application
  }
    
  process {
    $returnvalue = [PSCustomObject]@{
      Name          = Split-Path $Path -Leaf
      Folder        = Split-Path $Path
      DateTaken     = $null
      LastWriteTime = (Get-Item $Path).LastWriteTime
    }
        
    $shellfolder = $shell.Namespace($returnvalue.Folder)
    if ($shellfolder) {
      $shellfile = $shellfolder.ParseName($returnvalue.Name)
      if ($shellfile) {
        $returnvalue.DateTaken = $shellfolder.GetDetailsOf($shellfile, 12)
      }
    }

    $returnvalue
  }
}

# Function to move pictures based on their date
function Move-Pictures {
  param (
    [String]
    $PicturePath,
    [String]
    $TargetPath
  )

  $files = Get-ChildItem -Path $PicturePath -Recurse -File -ErrorAction SilentlyContinue
  $i = 0

  foreach ($file in $files) {
    $DateTaken = $file | Get-DateTaken

    if ($DateTaken.DateTaken) {
      $date = Get-Date -Date ($DateTaken.DateTaken -replace "[^0-9/\:\s]")
    }
    else {
      $date = $file.LastWriteTime
    }

    $year = $date.ToString("yyyy")
    $month = $date.ToString("MMMM", (Get-Culture).DateTimeFormat)
    $day = $date.ToString("dd")

    Write-Host "Processing $($file.Name) - Date: $($day) $($month) $($year)" -ForegroundColor Yellow

    $Directory = Join-Path -Path $TargetPath -ChildPath "$year\$month"

    if (!(Test-Path $Directory)) {
      New-Item -Path $Directory -ItemType Directory | Out-Null
      Write-Host "Creating folder $($Directory)" -ForegroundColor Green
    }

    if ($file.DirectoryName -ne $Directory) {
      Move-Item -Path $file.FullName -Destination $Directory
      Write-Host "Moving $($file.Name) to $($Directory)" -ForegroundColor Yellow
    }
    else {
      Write-Host "$($file.Name) is already in the correct folder." -ForegroundColor Green
    }

    $i++
    Write-Progress -Activity "Organizing photos" -Status "Progress: $i of $($files.Count)" -PercentComplete (($i / $files.Count) * 100)
  }
}

# Function to remove empty folders
function Remove-EmptyFolders {
  param (
    [String]
    $Path
  )

  $folders = Get-ChildItem -Path $Path -Recurse -Directory

  foreach ($folder in $folders) {
    if (-not (Get-ChildItem -Path $folder.FullName)) {
      Write-Host "Removing empty folder: $($folder.FullName)" -ForegroundColor Red
      Remove-Item -Path $folder.FullName -Force
    }
  }
}

# Main execution block
try {
  Move-Pictures -PicturePath $PicturePath -TargetPath $TargetPath

  # Uncomment to remove empty folders
  # Remove-EmptyFolders -Path $PicturePath
  # Remove-EmptyFolders -Path $TargetPath
}
catch {
  Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
}