#requires -version 5
<#
.SYNOPSIS
  This script organizes photos by date, moving them into folders structured by year and month.
.DESCRIPTION
  The script scans a directory for photo files, retrieves their date of capture (or last write time if the capture date is unavailable), and moves the files into a folder structure based on year and month. It can also remove empty folders after sorting.
.PARAMETER SourcePath
  The path to the folder containing the photos to be organized.
.PARAMETER TargetPath
  The path to the folder where the photos will be moved and organized by date.
.INPUTS
  Strings for SourcePath and TargetPath.
.OUTPUTS
  None. Moves files and organizes them into directories.
.NOTES
  Version:        1.0
  Author:         MrManjah
  Creation Date:  14/10/2024
  Purpose/Change: Initial script development
.EXAMPLE
  Move-Pictures -SourcePath "\\nas\Photos\john a trier" -TargetPath "\\nas\Photos\"
  
  This command moves photos from the source folder to the target folder, organizing them by year and month.
#>

#---------------------------------------------------------[Script Parameters]------------------------------------------------------

Param (
  [Parameter(Mandatory = $true,
    HelpMessage = "Path to the folder containing photos to be organized. Ex: \\nas\SourcePath\")]
  [String]$SourcePath,
  
  [Parameter(Mandatory = $true,
    HelpMessage = "Target folder where photos should be moved. Ex: \\nas\TargetPath\")]
  [String]$TargetPath
)


#---------------------------------------------------------[Initialisations]--------------------------------------------------------

# Set Error Action to Silently Continue
$ErrorActionPreference = 'SilentlyContinue'

# Import Modules & Snap-ins
# Example: Import-Module PSLogging (if needed)

#----------------------------------------------------------[Declarations]----------------------------------------------------------

# Global declarations, if any, go here (e.g., version, counters)

#-----------------------------------------------------------[Functions]------------------------------------------------------------

Function Get-DateTaken {
  Param ([String]$Path)
  Begin {
    Write-Host 'Retrieving DateTaken metadata...'
  }
  Process {
    Try {
      $shell = New-Object -COMObject Shell.Application
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
      return $returnvalue
    }
    Catch {
      Write-Host -BackgroundColor Red "Error retrieving DateTaken for $($Path): $($_.Exception.Message)"
      Break
    }
  }
  End {
    If ($?) {
      Write-Host 'DateTaken retrieved successfully.'
    }
  }
}

Function Move-Pictures {
  Param ([String]$SourcePath, [String]$TargetPath)
  Begin {
    Write-Host 'Starting to move pictures...'
  }
  Process {
    Try {
      $files = Get-ChildItem -Path $SourcePath -Recurse -File
      $i = 0

      foreach ($file in $files) {
        $DateTaken = Get-DateTaken -Path $file.FullName

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

        $Directory = Join-Path -Path $TargetPath -ChildPath "$($year)\$($month)"

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
    Catch {
      Write-Host -BackgroundColor Red "Error while moving pictures: $($_.Exception.Message)"
      Break
    }
  }
  End {
    If ($?) {
      Write-Host 'All pictures moved successfully.'
    }
  }
}

Function Remove-EmptyFolders {
  Param ([String]$Path)
  Begin {
    Write-Host 'Starting to remove empty folders...'
  }
  Process {
    Try {
      $folders = Get-ChildItem -Path $Path -Recurse -Directory

      foreach ($folder in $folders) {
        if (-not (Get-ChildItem -Path $folder.FullName)) {
          Write-Host "Removing empty folder: $($folder.FullName)" -ForegroundColor Red
          Remove-Item -Path $folder.FullName -Force
        }
      }
    }
    Catch {
      Write-Host -BackgroundColor Red "Error while removing folders: $($_.Exception.Message)"
      Break
    }
  }
  End {
    If ($?) {
      Write-Host 'Empty folders removed successfully.'
    }
  }
}

#-----------------------------------------------------------[Execution]------------------------------------------------------------

# Main execution block
try {
  Move-Pictures -SourcePath $SourcePath -TargetPath $TargetPath
  # Uncomment to remove empty folders
  # Remove-EmptyFolders -Path $SourcePath
  # Remove-EmptyFolders -Path $TargetPath
}
catch {
  Write-Host -BackgroundColor Red "Error during execution: $($_.Exception.Message)"
}
