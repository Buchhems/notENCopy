# 📝 NotENCopy – Notenübertragung für LUSD

**NotENCopy** ist ein praktisches Windows-Tool zur schnellen und fehlerfreien Übertragung von Notenlisten aus Excel in den externen Notenclient (ENC) der LUSD-Datenbank des Landes Hessen.

![Screenshot](https://github.com/user-attachments/assets/f79d57a4-1e1c-4dc9-a76e-324a127b40b3)


---

## 🚀 Funktionen

- 📋 Kopiert Noten direkt aus der Zwischenablage in ENC
- 🔗 Verbindet sich automatisch mit dem geöffneten ENC-Fenster
- 🧮 Berechnet automatisch die Anzahl der Schüler und Spalten
- 🛠 GUI mit Buttons für Hilfe, Verbindung und Übertragung

---

## 🖥️ Benutzeroberfläche

| Element                  | Beschreibung                                      |
|--------------------------|--------------------------------------------------|
| **Mit ENC verbinden**    | Stellt Verbindung zum geöffneten ENC her         |
| **Zwischenablage >> ENC**| Überträgt kopierte Noten in die LUSD-Datenbank   |
| **Hilfe**                | Zeigt eine Anleitung zur Benutzung               |
| **Statusanzeige**        | Zeigt Version und Entwicklerinformationen        |

---

## 📦 Voraussetzungen

- Windows-PC mit **ENC-Client** der LUSD
- Excel-Datei mit Notenspalten

---

## 🧑‍🏫 Anleitung

1. **ENC starten** und einloggen  
2. In **NotENCopy** auf „Mit ENC verbinden“ klicken  
3. Notenspalten aus Excel kopieren  
4. In NotENCopy auf „Zwischenablage >> ENC“ klicken  
5. Die Noten werden automatisch in die richtige Zelle eingetragen

Hier nochmals in Videoform: https://youtu.be/UOjbGazMLwM


---

## ⚠️ Häufige Fehler

- **Zeilen überspringen**: Anzahl Spalten in Excel und ENC müssen übereinstimmen
- **Verrückte Spalten**: Noten müssen korrekt formatiert sein (z. B. zweistellig: `01`, `02`, …)
- **Keine Verbindung**: ENC-Fenster muss aktiv und sichtbar sein

---

## 🧩 Technische Details

- Programmiert in **AutoIt v3**
- GUI erstellt mit **Koda GUI Designer**
- Nutzt `ClipGet()` zur Zwischenablage-Auswertung
- Automatisiert Tastatureingaben mit `Send()`
- Fenstererkennung über `WinWaitActive()` und `WinGetHandle()`

---

## 📁 Dateien

| Datei              | Zweck                                      |
|--------------------|--------------------------------------------|
| `Notencopy.au3`    | Hauptskript                                |
| `Notencopy.ico`    | Icon für GUI                               |
| `Notencopy.jpg`    | Bild für GUI                               |

---

## 📜 Lizenz

Benutzung auf eigene Gefahr.  
Software ohne Gewähr.  
Keine Haftung für Datenverlust oder Fehlübertragungen.

---

## 🧠 Erweiterungsideen

- Automatische Validierung der Notenformate
- Integration mit Schulverwaltungssoftware

---

## 🔧 Kompilierung

Zur Erstellung einer ausführbaren Datei (`.exe`) muss die AU-Datei in AutoIt kompiliert werden.
