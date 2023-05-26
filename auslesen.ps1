function Suche-XLS-Dateien {
  param (
    [Parameter(Mandatory = $true)]
    [string]$Verzeichnis
  )

  # Durchsucht das Verzeichnis nach .xls-Dateien
  $dateien = Get-ChildItem -Path $Verzeichnis -Filter "*.xls" -File

  if ($dateien.Count -gt 0) {
    Write-Host "Gefundene .xls-Dateien im Verzeichnis '$Verzeichnis':"
    $dateien | foreach { Write-Host $_.FullName }
  }

  # Durchsucht die Unterverzeichnisse
  $verzeichnisse = Get-ChildItem -Path $Verzeichnis -Directory

  foreach ($verzeichnis in $verzeichnisse) {
    Suche-XLS-Dateien -Verzeichnis $verzeichnis.FullName
  }
}

# Beispielaufruf
Suche-XLS-Dateien -Verzeichnis "C:\Pfad\Zum\Verzeichnis"