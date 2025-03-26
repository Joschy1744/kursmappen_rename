# PowerShell Skript zur Umbenennung von PDF-Dateien basierend auf OCR-Text
$FolderPath = $PSScriptRoot


# Log-Funktion
function Write-Log {
    param ([string]$Message)
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$Timestamp $Message" | Out-File -Append -Encoding utf8 "umbenennung_ocr_log.txt"
    Write-Host $Message
}

# Funktion zur Erkennung des Halbjahres
function Get-Halbjahr {
    param ([string]$Text)
    if ($Text -match "2.?\s*Halbjahr") { return "2. Halbjahr" }
    elseif ($Text -match "1.?\s*Halbjahr") { return "1. Halbjahr" }
    else { return "Unbekanntes_Halbjahr" }
}

# Funktion zum Extrahieren des Schuljahrs
function Get-Schuljahr {
    param ([string]$Text)
    if ($Text -match "Schuljahr\s+(\d{4})/\d{4}") {
        return $Matches[1]
    }
    return "Unbekanntes_Jahr"
}

# Funktion zur Extraktion von Fach, Klasse und Kursbezeichnung
function Get-FachKlasseKurs {
    param ([string[]]$Lines, [string]$Halbjahr)
    
    # Finde die Zeile mit dem Halbjahr
    $HalbjahrIndex = $Lines | Select-String -Pattern $Halbjahr | Select-Object -First 1 | ForEach-Object { $_.LineNumber }
    
    # Extrahiere Fach, Klasse und Kursbezeichnung aus der nächsten Zeile nach dem Halbjahr
    if ($HalbjahrIndex -and ($HalbjahrIndex -lt $Lines.Count)) {
        $NextLine = $Lines[$HalbjahrIndex]  # Nächste Zeile nach dem Halbjahr
        Write-Log "Zeile nach Halbjahr: $NextLine"

        # Regulärer Ausdruck für Fach, Klasse und Kursbezeichnung
        if ($NextLine -match "(\w+)\s+([A-Za-z0-9]+)\s*\(([^)]+)\)") {
            return @{
                Fach = $Matches[1]
                Klasse = $Matches[2]
                Kursbezeichnung = $Matches[3]
            }
        }
    }
    
    # Rückgabe bei nicht gefundenem Ergebnis
    return @{ Fach = "Unbekannt"; Klasse = "Unbekannt"; Kursbezeichnung = "Unbekannt" }
}


# Funktion zur Texterkennung und Umbenennung der Datei
function Rename-Pdf { 
    param ([string]$PdfPath)
   
    Write-Log "📄 Verarbeite: $PdfPath"
    
    # Konvertiere erste PDF-Seite in ein Bild
    $ImagePath = "$PdfPath.png"
    & "pdftoppm" -png -f 1 -singlefile "$PdfPath" "$ImagePath"
    
    if (!(Test-Path "$ImagePath.png")) {
        Write-Log "❌ Fehler beim Konvertieren der PDF"
        return
    }
    
  # OCR ausführen
$OcrText = & "tesseract" "$ImagePath.png" stdout -l deu

# Zerlege den OCR-Text in Zeilen
$Lines = $OcrText -split "`n" | Where-Object { $_ -match "\S" }

# Konvertiere jede Zeile nachträglich in UTF-8
$Lines = $Lines | ForEach-Object {
    [System.Text.Encoding]::UTF8.GetString([System.Text.Encoding]::Default.GetBytes($_))
}

# Log-Ausgabe zum Überprüfen
#$Lines | ForEach-Object { Write-Log "Zeile: $_" }
# Log-Ausgabe
#$OcrText | fl


    # Halbjahr extrahieren
    $Halbjahr = Get-Halbjahr $OcrText
    Write-Log "📆 Erkanntes Halbjahr: $Halbjahr"

     # Schuljahr aus OCR-Text extrahieren
    $Schuljahr = Get-Schuljahr $OcrText
    Write-Log "📅 Erkanntes Schuljahr: $Schuljahr"
   
  # Extrahiere Fach, Klasse und Kursbezeichnung
$FachKlasseKurs = Get-FachKlasseKurs -Lines $Lines -Halbjahr $Halbjahr
Write-Log "📖 Fach: $($FachKlasseKurs.Fach)"
Write-Log "🏫 Klasse: $($FachKlasseKurs.Klasse)"
Write-Log "📌 Kursbezeichnung: $($FachKlasseKurs.Kursbezeichnung)"


    # Wenn Fach unbekannt ist, versuche es durch die Zeile vor der "bei"-Zeile zu finden
    if ($FachKlasseKurs.Fach -eq "Unbekannt") {
        $BeiIndex = $Lines | Select-String -Pattern "^bei\s" | Select-Object -First 1 | ForEach-Object { $_.LineNumber }
        if ($BeiIndex -gt 0) {
            # Die Zeile vor der "bei"-Zeile könnte das Fach enthalten
            $PrevLine = $Lines[$BeiIndex - 2]
            Write-Log "Zeile vor 'bei': $PrevLine"
        if ($PrevLine -match "(\w+)\s+([A-Za-z0-9]+)") {
                # Fach und Klasse extrahieren
                $FachKlasseKurs.Fach = $Matches[1]
                $FachKlasseKurs.Klasse = $Matches[2]
                Write-Log "📖 Fach nach 'bei' gefunden: $($FachKlasseKurs.Fach)"
                Write-Log "🏫 Klasse nach 'bei' gefunden: $($FachKlasseKurs.Klasse)"
            }
        }
    }



# Extrahiere Lehrkraft und Kürzel
$LehrkraftMatch = ($Lines | Where-Object { $_ -match "^bei\s" }) -match "bei\s+([A-Za-zÄÖÜäöüß\-]+\s+[A-Za-zÄÖÜäöüß\-]+(?:\s+[A-Za-zÄÖÜäöüß\-]+)*)\s*\(([^)]+)\)"
if ($LehrkraftMatch) {
    # Sicherstellen, dass Matches vorhanden sind
    if ($Matches.Count -ge 3) {
        $LehrkraftName = $Matches[1].Trim() -replace "\s", "_"
        $LehrkraftKuerzel = $Matches[2]
    } else {
        $LehrkraftName = "Unbekannt"
        $LehrkraftKuerzel = "Unbekannt"
    }
} else {
    $LehrkraftName = "Unbekannt"
    $LehrkraftKuerzel = "Unbekannt"
}

Write-Log "👤 Lehrkraft: $LehrkraftName"
Write-Log "🔠 Kürzel: $LehrkraftKuerzel"

# Wenn mehr als drei Werte unbekannt sind oder das Jahr/Name nicht erkannt wurden, bleibe beim alten Namen
    if ( $LehrkraftName -eq "Unbekannt" -or $LehrkraftKuerzel -eq "Unbekannt" -or $FachKlasseKurs.Fach -eq "Unbekannt" -or $FachKlasseKurs.Klasse -eq "Unbekannt") {
        Write-Log "⚠️ Ein oder mehrere wichtige Werte sind unbekannt, behalte den alten Dateinamen bei: $PdfPath"
        # Bild löschen
        Remove-Item "$ImagePath.png" -Force
        return
    }

# Neuen Dateinamen generieren
$NewName = "$Schuljahr`_$Halbjahr`_$LehrkraftKuerzel`_$LehrkraftName`_$($FachKlasseKurs.Fach)`_$($FachKlasseKurs.Klasse)`_$($FachKlasseKurs.Kursbezeichnung)" -replace "\s", "" -replace "\.", "_" -replace "\,", "_"

$NewPath = [System.IO.Path]::Combine((Split-Path -Parent $PdfPath), "$NewName.pdf")

# Überprüfen, ob die Datei schon existiert, und fortlaufend nummerieren
$Counter = 1
while (Test-Path $NewPath) {
    $NewPath = [System.IO.Path]::Combine((Split-Path -Parent $PdfPath), "$NewName" + "_$Counter.pdf")
    $Counter++
}

# Datei umbenennen
try {
    Rename-Item -Path $PdfPath -NewName $NewPath -Force
    Write-Log "✅ Umbenannt zu: $NewPath"
} catch {
    Write-Log "❌ Fehler beim Umbenennen: $_"
}


    # Bild löschen
    Remove-Item "$ImagePath.png" -Force
}


    
# Alle PDFs im Verzeichnis verarbeiten
Get-ChildItem -Path $FolderPath -Filter "*.pdf" | ForEach-Object { Rename-Pdf $_.FullName }

