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

# Beispielaufruf
Clear-Host
Suche-XLS-Dateien -Verzeichnis "Q:\AppTesting\QFTestFrameWork\QFTestDriver\Syrius\CEN_UNI"