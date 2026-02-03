# 🚀 Quick Start - Apps Script Setup

Eine Schritt-für-Schritt-Anleitung zum Einrichten des Google Apps Script Backends.

## ⏱️ Geschätzte Zeit: 10 Minuten

---

## Schritt 1: Google Sheet erstellen (2 Min)

1. Öffne https://drive.google.com
2. Klicke auf "Neu" → "Google Tabellen" → "Leere Tabelle"
3. Benenne die Tabelle: **"Meine Arbeitszeiten"**
4. Lass das Tab offen

✅ **Fertig!** Dein Sheet ist bereit.

---

## Schritt 2: Apps Script Code einfügen (3 Min)

1. **Im Google Sheet**, klicke oben auf:
   - **Erweiterungen** → **Apps Script**
   
2. Ein neuer Tab öffnet sich mit dem Apps Script Editor

3. **Lösche** den vorhandenen Code:
   ```javascript
   function myFunction() {
     // Code wird gelöscht
   }
   ```

4. **Öffne** die Datei `Code.gs` (die du zusammen mit der App heruntergeladen hast)

5. **Kopiere** den gesamten Inhalt der Datei

6. **Füge** den Code in den Apps Script Editor ein

7. **Speichere**:
   - Klicke auf das 💾 Disketten-Symbol ODER
   - Drücke `Strg+S` (Windows) / `Cmd+S` (Mac)

8. **Benenne das Projekt**:
   - Oben links auf "Unbenanntes Projekt" klicken
   - Neuer Name: **"Zeiterfassung Backend"**
   - Klicke außerhalb zum Speichern

✅ **Fertig!** Der Code ist eingefügt.

---

## Schritt 3: Apps Script deployen (5 Min)

### 3.1 Neue Bereitstellung erstellen

1. Klicke oben rechts auf **"Bereitstellen"** → **"Neue Bereitstellung"**

2. Klicke auf das **Zahnrad-Symbol** ⚙️ neben "Typ auswählen"

3. Wähle **"Web-App"**

4. Konfiguriere die Einstellungen:
   ```
   Beschreibung: Zeiterfassung API (optional)
   Ausführen als: Ich (deine-email@gmail.com)
   Zugriff: Jeder
   ```
   
   **⚠️ WICHTIG:** "Zugriff: Jeder" ist korrekt! 
   Das bedeutet nur, dass deine App (die du kennst) darauf zugreifen kann.

5. Klicke **"Bereitstellen"**

### 3.2 Zugriff autorisieren

Du siehst jetzt eine Warnung - das ist normal!

1. Klicke **"Zugriff autorisieren"**

2. Wähle dein **Google-Konto**

3. Google zeigt: "Diese App wurde nicht von Google geprüft"
   - Das ist OK, es ist DEINE App!
   - Klicke **"Erweitert"**
   - Klicke **"Zu [Zeiterfassung Backend] wechseln (unsicher)"**

4. Google fragt nach Berechtigungen:
   - "Google-Tabellen ansehen, bearbeiten, erstellen und löschen"
   - Das ist korrekt - deine App braucht Zugriff auf dein Sheet
   - Klicke **"Zulassen"**

### 3.3 URL kopieren

1. Nach erfolgreicher Autorisierung siehst du:
   ```
   Web-App
   URL: https://script.google.com/macros/s/ABC123.../exec
   ```

2. **Kopiere diese URL!** Du brauchst sie gleich für die App.

3. Die URL sieht ungefähr so aus:
   ```
   https://script.google.com/macros/s/AKfycby...ABC123.../exec
   ```
   
   **⚠️ WICHTIG:** Die URL muss mit `/exec` enden!

✅ **Fertig!** Dein Backend ist live!

---

## Schritt 4: App konfigurieren (1 Min)

1. Öffne deine **Zeiterfassung Web-App** im Browser
   (z.B. `https://dein-name.github.io/timetracker/timetracker.html`)

2. Du siehst den **Setup-Assistenten**

3. Füge ein:
   - **Apps Script Web-App-URL**: Die URL aus Schritt 3.3
   - **Tabellenblatt-Name**: `Zeiterfassung` (Standard)

4. Klicke **"Verbindung testen & speichern"**

5. Warte 2-3 Sekunden...

6. ✅ **"Verbindung erfolgreich!"**

---

## 🎉 Geschafft!

Deine App ist jetzt einsatzbereit!

**Was jetzt passiert:**
- Du klickst auf einen Tracking-Button
- Die App sendet die Daten an dein Apps Script
- Apps Script schreibt sie in dein Google Sheet
- Alles synchronisiert automatisch!

**Öffne dein Google Sheet** und klicke auf einen Button - du siehst die Daten live erscheinen!

---

## 🔧 Testen

1. Wähle einen **Arbeitsmodus** (z.B. "Normal")
2. Klicke **"Losfahrt von Zuhause"**
3. Öffne dein **Google Sheet**
4. Du siehst eine neue Zeile mit:
   - Heutiges Datum
   - Modus: "Normal"
   - Zeit in der "Losfahrt Zuhause" Spalte

✅ **Es funktioniert!**

---

## 🆘 Probleme?

### Fehler: "Verbindung fehlgeschlagen"

**Mögliche Ursachen:**

1. **URL falsch kopiert**
   - Prüfe: URL muss mit `/exec` enden
   - Keine Leerzeichen am Anfang/Ende

2. **"Zugriff: Jeder" nicht gesetzt**
   - Gehe zu Apps Script
   - Klicke "Bereitstellen" → "Bereitstellungen verwalten"
   - Prüfe die Einstellung
   - Falls falsch: Klicke ✏️ Bearbeiten und ändere es

3. **App nicht autorisiert**
   - Du hast möglicherweise "Abbrechen" geklickt
   - Lösung: Neue Bereitstellung erstellen und Zugriff erlauben

### Fehler: "Authorization required"

Du hast die App noch nicht autorisiert.

**Lösung:**
1. Gehe zu Apps Script Editor
2. Klicke auf eine Funktion (z.B. `doPost`)
3. Klicke oben auf "Ausführen"
4. Autorisiere die App wie in Schritt 3.2

### Daten kommen nicht im Sheet an

1. **Öffne Apps Script Editor**
2. Klicke links auf **"Ausführungen"** (Uhr-Symbol)
3. Siehst du Fehler?
   - Wenn ja: Screenshot machen und Fehler prüfen
   - Meist: Autorisierung fehlt

---

## 📱 Nächste Schritte

1. **Installiere die App auf deinem Handy**
   - Android: Chrome Menü → "Zum Startbildschirm"
   - iPhone: Safari Teilen → "Zum Home-Bildschirm"

2. **Installiere auf Windows**
   - Chrome/Edge: Klicke auf Install-Symbol (oben rechts)

3. **Teste alle Modi**
   - Normal, HomeOffice, Stempeln, Urlaub, Krank

4. **Exportiere Daten**
   - Klicke "Daten als CSV exportieren"
   - Öffne in Excel

---

## 💡 Tipps

- **Google Sheet formatieren:** Du kannst das Sheet nach Belieben formatieren, Spalten färben etc.
- **Mehrere Geräte:** Die gleiche Apps Script URL funktioniert auf allen Geräten
- **Backup:** Dein Google Sheet ist automatisch in Google Drive gesichert
- **Teilen:** Du kannst das Sheet mit anderen teilen (nur ansehen!)

---

**Viel Erfolg! 🚀**

Bei weiteren Fragen: Prüfe die `README.md` für detaillierte Infos.
