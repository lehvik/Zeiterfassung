# ⏱️ Arbeitszeit Tracker - PWA

Eine Progressive Web App zur einfachen Erfassung deiner Arbeitszeiten mit Google Sheets Synchronisation.

## 🚀 Features

✅ **4 Tracking-Modi:**
- Normal (Pendeln ins Büro)
- HomeOffice
- Stempeln (nur Büro)
- Urlaub / Krank

✅ **PWA-Funktionen:**
- Installierbar auf allen Geräten (Android, iOS, Windows, Mac, Linux)
- Funktioniert offline
- Sieht aus wie eine native App

✅ **Intelligente Button-Logik:**
- Nächster Button wird hervorgehoben
- Geklickte Buttons zeigen die Zeit an
- Automatische Farbänderung

✅ **Google Sheets Integration:**
- Automatische Synchronisation
- Arbeitszeit-Berechnung
- Pendelzeit-Berechnung
- CSV-Export

## 📱 Installation

### Schritt 1: Hosting auf GitHub Pages

1. **GitHub Account erstellen** (falls noch nicht vorhanden)
   - Gehe zu https://github.com
   - Erstelle kostenlosen Account

2. **Neues Repository erstellen**
   - Klicke auf "New Repository"
   - Name: `timetracker` (oder beliebig)
   - Visibility: Public
   - Klicke "Create Repository"

3. **Dateien hochladen**
   - Klicke "uploading an existing file"
   - Ziehe alle 6 Dateien in das Fenster:
     - `timetracker.html`
     - `manifest.json`
     - `service-worker.js`
     - `icon-192.png`
     - `icon-512.png`
     - `Code.gs` (für deine Unterlagen, wird nicht auf GitHub hochgeladen)
   - Klicke "Commit changes"

4. **GitHub Pages aktivieren**
   - Gehe zu "Settings" (oben im Repository)
   - Linke Sidebar: "Pages"
   - Source: "Deploy from a branch"
   - Branch: "main" / Ordner: "/ (root)"
   - Klicke "Save"
   - ⏳ Warte 1-2 Minuten

5. **Deine App ist jetzt online!**
   - URL: `https://DEIN-USERNAME.github.io/timetracker/timetracker.html`
   - Beispiel: `https://max-mueller.github.io/timetracker/timetracker.html`

### Schritt 2: Google Apps Script einrichten

1. **Google Sheet erstellen**
   - Gehe zu https://drive.google.com
   - Neu → Google Tabellen → Leere Tabelle
   - Benenne sie z.B. "Meine Arbeitszeiten"

2. **Apps Script Editor öffnen**
   - Im Google Sheet: Menü "Erweiterungen" → "Apps Script"
   - Ein neuer Tab öffnet sich mit dem Apps Script Editor

3. **Code einfügen**
   - Lösche den vorhandenen Code (`function myFunction() { ... }`)
   - Öffne die Datei `Code.gs` (wurde mit der App bereitgestellt)
   - Kopiere den gesamten Inhalt
   - Füge ihn in den Apps Script Editor ein
   - Klicke auf das Disketten-Symbol (💾) zum Speichern
   - Gib dem Projekt einen Namen z.B. "Zeiterfassung Backend"

4. **Apps Script deployen (bereitstellen)**
   - Klicke oben rechts auf "Bereitstellen" → "Neue Bereitstellung"
   - Klicke auf "Typ auswählen" (⚙️) → Wähle "Web-App"
   - Einstellungen:
     - **Beschreibung**: "Zeiterfassung API" (optional)
     - **Ausführen als**: "Ich (deine-email@gmail.com)"
     - **Zugriff**: "Jeder"
   - Klicke "Bereitstellen"
   - **Wichtig**: Du musst der App Zugriff gewähren:
     - Klicke "Zugriff autorisieren"
     - Wähle dein Google-Konto
     - Klicke "Erweitert" → "Zu [Projektname] wechseln (unsicher)"
     - Klicke "Zulassen"
   - **Kopiere die Web-App-URL** (z.B. `https://script.google.com/macros/s/ABC.../exec`)
   - Diese URL brauchst du gleich!

### Schritt 3: App konfigurieren

1. Öffne deine App-URL im Browser
2. Du siehst den Setup-Assistenten
3. Gib ein:
   - **Apps Script Web-App-URL** (aus Schritt 2.4)
   - **Tabellenblatt-Name** (Standard: "Zeiterfassung")
4. Klicke "Verbindung testen & speichern"
5. ✅ Fertig!

### Schritt 4: App installieren

#### Auf Android:
1. Öffne die App-URL in Chrome
2. Tippe auf das Menü (⋮) → "Zum Startbildschirm hinzufügen"
3. Oder warte auf den Install-Dialog
4. Die App erscheint wie eine normale App auf deinem Homescreen

#### Auf Windows:
1. Öffne die App-URL in Chrome oder Edge
2. Klicke oben rechts auf das Install-Symbol (⊕)
3. Oder: Menü (⋮) → "App installieren"
4. Die App erscheint im Startmenü und kann an Taskleiste gepinnt werden

#### Auf iPhone/iPad:
1. Öffne die App-URL in Safari
2. Tippe unten auf "Teilen" (□↑)
3. "Zum Home-Bildschirm" → "Hinzufügen"
4. Die App erscheint auf deinem Homescreen

## 🎯 Nutzung

### Tracking
1. Wähle deinen Arbeitsmodus (Normal, HomeOffice, etc.)
2. Klicke die Buttons in der richtigen Reihenfolge
3. Die Zeit wird automatisch erfasst und zu Google Sheets synchronisiert

### Modi
- **Normal**: Alle 4 Zeiten (Losfahrt Zuhause → Ankunft Arbeit → Losfahrt Arbeit → Ankunft Zuhause)
- **HomeOffice**: Nur Arbeitsbeginn und -ende
- **Stempeln**: Nur Ankunft/Verlassen Büro
- **Urlaub/Krank**: Keine Zeiterfassung, nur Markierung

### Export
- Klicke auf "Daten als CSV exportieren"
- Die CSV-Datei kann direkt in Excel geöffnet werden

## 📊 Google Sheets Struktur

```
Datum | Modus | Losfahrt Zuhause | Ankunft Arbeit | Losfahrt Arbeit | Ankunft Zuhause | Arbeitszeit | Pendelzeit
```

Die Arbeitszeit und Pendelzeit werden automatisch berechnet.

## 🔒 Datenschutz

- Deine Apps Script URL wird nur lokal in deinem Browser gespeichert
- Die App kommuniziert direkt mit deinem Google Apps Script (läuft in deinem Google Account)
- Alle Daten bleiben in deinem Google Account
- Kein Drittanbieter-Server involviert
- Das Apps Script läuft unter deiner Identität und hat nur Zugriff auf dein Sheet

## 🛠️ Technologie

- **Frontend**: HTML5 / CSS3 / Vanilla JavaScript
- **Backend**: Google Apps Script (serverless)
- **Datenspeicherung**: Google Sheets
- **PWA**: Progressive Web App mit Service Worker
- **Offline**: LocalStorage für Konfiguration
- **Deployment**: GitHub Pages (kostenlos)

## 📝 Lizenz

Privates Projekt - Frei zur persönlichen Nutzung

## 🆘 Probleme?

**App lädt nicht:**
- Prüfe, ob alle 6 Dateien hochgeladen wurden (inkl. Code.gs)
- Warte 2-3 Minuten nach GitHub Pages Aktivierung

**"Verbindung fehlgeschlagen":**
- Prüfe die Apps Script Web-App-URL (muss mit `/exec` enden)
- Stelle sicher, dass du die App autorisiert hast
- Prüfe, dass "Zugriff: Jeder" eingestellt ist

**Daten werden nicht gespeichert:**
- Überprüfe die Internet-Verbindung
- Öffne das Google Sheet und prüfe, ob Daten ankommen
- Prüfe in Apps Script die "Ausführungen" (Menü links) für Fehler

**Apps Script Fehler "Unauthorized":**
- Gehe zu Apps Script Editor
- Klicke "Bereitstellen" → "Bereitstellungen verwalten"
- Prüfe: "Ausführen als" = "Ich" und "Zugriff" = "Jeder"
- Falls nötig: Neue Bereitstellung erstellen

## 🎉 Viel Erfolg!

Deine Zeiterfassung läuft jetzt professionell auf allen Geräten!
