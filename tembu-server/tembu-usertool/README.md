# Tembu User Tool

Lokales Konsolen-Tool zur Verwaltung von `tembu-server`-Konfigurationsdateien. Läuft auf derselben Maschine wie der Server — bearbeitet Dateien direkt, kein Netzwerkzugriff nötig.

```
dotnet run --project tembu-usertool
```
oder die gebaute `tembu-usertool.exe` direkt starten.

## Secret

Beim Start fragt das Tool nach dem `License:Secret` (identisch mit `tembu-server`s `appsettings.Local.json`) — nötig für Menüpunkte [1]–[3], die damit `users.dat` ver-/entschlüsseln bzw. Lizenzschlüssel generieren. Eingabe versteckt. Kann mit der Umgebungsvariable `TEMBU_LICENSE_SECRET` übersprungen werden (z.B. für Skripting).

## Menüpunkte

**[1] Neue users.dat erstellen**
Legt eine verschlüsselte `users.dat` mit Kundenname und einer Liste lizenzierter E-Mail-Adressen an. Ohne `users.dat` sind alle Nutzer erlaubt (Solo-Betrieb).

**[2] Lizenzschlüssel generieren**
Erzeugt aus Kunden-E-Mail + Ablaufdatum den `License:Key` für `appsettings.json`.

**[3] Bestehende users.dat anzeigen**
Zeigt Kunde, Erstellungsdatum und lizenzierte Nutzer einer `users.dat` an.

**[4] KI-Einstellungen bearbeiten (`ai-settings.json`)**
Bearbeitet Provider, API-Key, Model und Endpoint für die KI-Analyse zentral für alle Nutzer des Servers.
- Enter bei jeder Frage = Wert unverändert lassen
- API-Key wird bei der Eingabe maskiert (`***`) und in der Übersicht abgekürzt angezeigt
- Provider per Nummernauswahl: Gemini, Claude (Anthropic), OpenAI, Groq, Ollama, LM Studio
- Bei Ollama/LM Studio wird statt API-Key ein Endpoint abgefragt (Standard: `http://localhost:11434` bzw. `http://localhost:1234`)
- Ein bereits laufender `tembu-server` übernimmt die Änderung sofort — kein Neustart nötig (Live-Reload)

Die Datei muss mit `tembu-server/Models/AiSettings.cs` kompatibel bleiben (Felder `Provider`/`ApiKey`/`Model`/`Endpoint` unter dem Schlüssel `"AI"`).

**[5] Erlaubte Server-Origins bearbeiten (`cors-settings.json`)**
Verwaltet, von welchen Web-Adressen aus der Browser Anfragen an `tembu-server` stellen darf (CORS).
- Origins per `[a]` hinzufügen, `[r]` + Nummer entfernen, `[s]` speichern
- Ein bereits laufender `tembu-server` übernimmt die Änderung sofort — kein Neustart nötig (Live-Reload)

Die Datei muss mit `tembu-server/Models/CorsSettings.cs` kompatibel bleiben (Feld `AllowedOrigins` unter dem Schlüssel `"Cors"`).
