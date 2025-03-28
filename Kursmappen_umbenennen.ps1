# PowerShell Skript zur Umbenennung von PDF-Dateien basierend auf OCR-Text
$FolderPath = $PSScriptRoot


# Log-Funktion
function Write-Log {
    param ([string]$Message)
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$Timestamp $Message" | Out-File -Append -Encoding utf8 "umbenennung_ocr_log.txt"
    Write-Host $Message
}


# Liste der bekannten Fächer (kann erweitert werden)
$FaecherListe = @("Klasenlehrer", "Tutor", "Englisch", "Deutsch", "Mathematik", "Geschichte", "Politik und Wirtschaft", "Politik", "Powi", "Ethik", "Spanisch", "Französisch", "Arbeitslehre", "Schreibförderung", "Sport", "Chemie", "Physik", "Biologie", "Erdkunde", "Geographie", "Kunst", "Musik", "Darstellendes Spiel", "Religion")

# Sucht in den ertsten 10 Zeilen nach dem Vorkommen des Faches
function Find-FachInTopLines {
    param (
        [string[]]$Lines
    )

    
    
    # Durchsuche die ersten 10 Zeilen nach einem Fach (auch als Teilwort)
    for ($i = 0; $i -lt [Math]::Min(10, $Lines.Count); $i++) {
        foreach ($Fach in $FaecherListe) {
            if ($Lines[$i] -match "($Fach)\w*") {  # \w* erlaubt Erweiterungen (z. B. Sporta → Sport)
                Write-Log "📖 Fach in den ersten 10 Zeilen gefunden: $Matches[1]"
                return $Matches[1]  # Gibt das Basisfach zurück
            }
        }
    }

    return "Unbekannt"
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

# Erkannte Muster für Kursbezeichnungen
# Beispiel	Erklärung
# 101G21	Muster 1: 3 Zahlen + 1-4 Buchstaben + 2 Zahlen
# -1AB20	Muster 2: -1 oder -2 + 1-4 Buchstaben + 2 Zahlen
# ÜGMA23	Muster 3: ÜG + 1-4 Buchstaben + 2-4 Zahlen
# Q1GK24	Muster 4: E1, E2, Q1, Q2, Q3, Q4 + 1-4 Buchstaben + 2 Zahlen

function Find-Kursbezeichnung {
    param (
        [string[]]$Lines
    )

    # Reguläre Ausdrücke für Kursbezeichnungen
    $Patterns = @(
        "\b(\d{3}[A-Za-z]{1,4}\d{2})\b",    # Muster 1: 101G21
        "\b(-[12][A-Za-z]{1,4}\d{2})\b",    # Muster 2: -1AB20 oder -2XYZ99
        "\b(ÜG[A-Za-z]{1,4}\d{2,4})\b",     # Muster 3: ÜGMA23
        "\b(E[12]|Q[1-4])[A-Za-z]{1,4}\d{2}\b" # Muster 4: E1GK24, Q2MA21
    )

    # Durchsuche die ersten 10 Zeilen nach einer Kursbezeichnung
    for ($i = 0; $i -lt [Math]::Min(10, $Lines.Count); $i++) {
        foreach ($Pattern in $Patterns) {
            if ($Lines[$i] -match $Pattern) {
                Write-Log "📌 Kursbezeichnung gefunden: $Matches[1]"
                return $Matches[1]  # Gibt den gefundenen Wert zurück
            }
        }
    }

    return "Unbekannt"
}

# Muster für Klassennamen
# Beispiel	Erklärung
# 10ah	Regulär: 2 Zahlen + 1 Buchstabe (a-f) + 1 Buchstabe (h, r, g)
# 05 a g	Mit Leerzeichen: 2 Zahlen + Leerzeichen + Buchstabe (a-f) + Leerzeichen + (h, r, g)
# O9ah	OCR-Fehler: O erkannt als 0


function Find-Klasse {
    param (
        [string[]]$Lines
    )

    # Regulärer Ausdruck für Klassennamen
    $Pattern = "\b([01O][5-9O]|10)\s?([a-fA-F])\s?([hrgHRG])\b"

    # Durchsuche die ersten 10 Zeilen nach einer Klasse
    for ($i = 0; $i -lt [Math]::Min(10, $Lines.Count); $i++) {
        if ($Lines[$i] -match $Pattern) {
            # Ersetze falsch erkannte "O" mit "0"
            $Klasse = "$($Matches[1])$($Matches[2])$($Matches[3])" -replace "O", "0"
            Write-Log "🏫 Klasse gefunden: $Klasse"
            return $Klasse  # Gibt den gefundenen Klassennamen zurück
        }
    }

    return "Unbekannt"
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
$Lines | ForEach-Object { Write-Log "Zeile: $_" }
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
        # Die Zeile vor der "bei"-Zeile könnte Fach und Klasse enthalten
        $PrevLine = $Lines[$BeiIndex - 2]
        Write-Log "Zeile vor 'bei': $PrevLine"

        # Regulärer Ausdruck zur Extraktion des Fachs, der Klassen und des Kurses
        if ($PrevLine -match "^(.+?)\s*\(([^)]+)\)\s*(\(([^)]+)\))?$") {
            $FachKlasseKurs.Fach = $Matches[1].Trim()
            $FachKlasseKurs.Klasse = $Matches[2].Trim()
            $FachKlasseKurs.Kursbezeichnung = if ($Matches[4]) { $Matches[4].Trim() } else { "Unbekannt" }
            Write-Log "📖 Fach nach 'bei' gefunden: $($FachKlasseKurs.Fach)"
            Write-Log "🏫 Klasse nach 'bei' gefunden: $($FachKlasseKurs.Klasse)"
            Write-Log "📌 Kursbezeichnung nach 'bei' gefunden: $($FachKlasseKurs.Kursbezeichnung)"
        }
    }
}

if ($FachKlasseKurs.Fach -eq "Unbekannt") {
    $FachKlasseKurs.Fach = Find-FachInTopLines -Lines $Lines
}

if ($FachKlasseKurs.Kursbezeichnung -eq "Unbekannt") {
    $FachKlasseKurs.Kursbezeichnung = Find-Kursbezeichnung -Lines $Lines
}

if ($FachKlasseKurs.Klasse -eq "Unbekannt") {
    $FachKlasseKurs.Klasse = Find-Klasse -Lines $Lines
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
    if (( $LehrkraftName -eq "Unbekannt" -or $LehrkraftKuerzel -eq "Unbekannt" ) -and $FachKlasseKurs.Fach -eq "Unbekannt" ) {
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

