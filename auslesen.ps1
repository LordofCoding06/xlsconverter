$directoryhistory = New-Object System.Collections.Generic.Stack[string]

function search-navigate-directory {
  param (
    [Parameter(Mandatory = $true)]
    [string]$directory
  )

  Clear-Host

  search-xls-files -directory $directory
  $directories = Get-ChildItem -Path $directory -Directory

  for ($i = 0; $i -lt $directories.Count; $i++) {
    Write-Host "[$i] $($directories[$i].FullName)"
  }

  $selection = Read-Host "Enter the index number of the desired subdirectory (or 'z' to go back, 'q' to quit)"

  if ($selection -eq "q") {
    return
  }

  if ($selection -eq "z") {
    if ($directoryhistory.Count -gt 0) {
      $previousdirectory = $directoryhistory.Pop()
      search-navigate-directory -directory $previousdirectory
    }
    else {
      Write-Host "You are already in the top-level directory."
      search-navigate-directory -directory $directory
    }
    return
  }

  if ($selection -lt 0 -or $selection -ge $directories.Count) {
    Write-Host "Invalid selection. Please enter a valid index number."
    search-navigate-directory -directory $directory
    return
  }

  $newdirectory = $directories[$selection].FullName
  $directoryhistory.Push($directory)
  search-navigate-directory -directory $newdirectory
}

function search-xls-files {
  param (
    [Parameter(Mandatory = $true)]
    [string]$directory
  )

  Clear-Host

  $files = Get-ChildItem -Path $directory -Filter "*.xls" -File

  if ($files.Count -gt 0) {
    Write-Host "Found .xls files in directory '$directory':"
    $files | foreach { Write-Host $_.Name }
  }
  else {
    Write-Host "Nothing found in directory:    $directory"
  }
}