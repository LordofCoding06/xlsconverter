function Suche-Navigiere-Verzeichnis {
  param (
      [Parameter(Mandatory = $true)]
      [string]$Verzeichnis
  )

  # Durchsucht die Unterverzeichnisse rekursiv
  $verzeichnisse = Get-ChildItem -Path $Verzeichnis -Directory

  # Indizierte Ausgabe der Unterverzeichnisse
  for ($i = 0; $i -lt $verzeichnisse.Count; $i++) {
      Write-Host "[$i] $($verzeichnisse[$i].FullName)"
  }

  # Auswahl des Unterverzeichnisses
  $auswahl = Read-Host "Geben Sie die Indexnummer des gewünschten Unterverzeichnisses ein (oder 'q' zum Beenden):"

  if ($auswahl -eq "q") {
      return
  }

  # Validierung der Auswahl
  if ($auswahl -lt 0 -or $auswahl -ge $verzeichnisse.Count) {
      Write-Host "Ungültige Auswahl. Bitte geben Sie eine gültige Indexnummer ein."
      Suche-Navigiere-Verzeichnis -Verzeichnis $Verzeichnis
      return
  }

  # Navigieren in das ausgewählte Unterverzeichnis
  Suche-Navigiere-Verzeichnis -Verzeichnis $verzeichnisse[$auswahl].FullName
}

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
  else {
    Write-Host "Nichts gefunen im Verzeichnis:    $Verzeichnis"
  }
}

Clear-Host
# Suche-XLS-Dateien -Verzeichnis "Q:\AppTesting\QFTestFrameWork\QFTestDriver\Syrius\CEN_UNI"
Suche-Navigiere-Verzeichnis -Verzeichnis "Q:\AppTesting\QFTestFrameWork\QFTestDriver\Syrius\CEN_UNI"