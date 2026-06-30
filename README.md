<div align="center">
  <img src="assets/logo.svg" width="84" alt="Live Web Region" />
  <h1>Live Web Region für PowerPoint</h1>
  <p><b>Zeigt eine lokale HTML/JS/CSS-Datei live in einem Folienbereich an – interaktiv, auch im Präsentationsmodus.</b></p>
</div>

---

## Was es macht

Mit *Live Web Region* markierst du in PowerPoint eine Form und weist ihr eine **Website** zu –
entweder eine **Online-URL** (`https://…`) oder eine **lokale HTML-Datei**. Im
**Präsentationsmodus** wird in diesem Bereich ein echter Browser (Microsoft Edge **WebView2**)
eingeblendet, der die Seite **live und interaktiv** darstellt – inklusive JavaScript,
Animationen, Klicks und Eingaben. So lässt sich z. B. eine Web-App oder Demo-Seite mitten in
der Präsentation vorführen.

- ✅ Online-URL **oder** lokale HTML/JS/CSS-Datei (echter Browser, kein iframe-Sandbox-Limit)
- ✅ Voll interaktiv im Präsentationsmodus (Maus, Tastatur, JS)
- ✅ Mehrere Bereiche pro Folie
- ✅ Pfeiltasten / Bild↑↓ / Esc steuern die Präsentation weiter – auch während die Seite Fokus hat
- ✅ Multi-Monitor

> Plattform: **Windows-Desktop-PowerPoint (64-Bit)**. Kein Mac/Web/Mobile (technisch bedingt,
> da ein echter eingebetteter Browser genutzt wird).

## Installation (für Anwender)

Den Ordner **`dist/LiveWebRegion`** (aus einem Release) auf den Rechner kopieren, dann **eine**
der beiden Varianten:

- **Variante A – am einfachsten:** **`LiveWebRegionSetup.exe`** doppelklicken.
  Self-contained Installer, **pro Benutzer ohne Administratorrechte**. Deinstallation:
  `LiveWebRegionSetup.exe /uninstall`.
- **Variante B – Skript:** **`Install.cmd`** doppelklicken. (Deinstallation: `Uninstall.cmd`.)

In beiden Fällen: Fehlt die **WebView2-Runtime**, öffnet sich automatisch die Download-Seite –
danach das Setup erneut starten. Anschließend PowerPoint starten → Reiter **„Live Web"**.

> Hinweis: `.ps1`-Dateien sind PowerShell-Skripte, **keine** `.exe` – Windows öffnet sie beim
> Doppelklick nur im Editor. Deshalb gibt es die **`.exe`** und die **`.cmd`** als bequeme
> Doppelklick-Starter.

## Verwendung

1. **Link setzen** (Reiter *Live Web*): Ist eine Form markiert, wird sie genutzt; sonst wird
   automatisch ein Rechteck eingefügt. Anschließend **URL eingeben oder HTML-Datei wählen**.
2. Form nach Wunsch positionieren/skalieren.
3. **F5** – die Seite läuft live im Bereich.
   - **Interaktion:** Klicks/Tippen gehen an die Seite.
   - **Navigation:** Pfeiltasten, Bild↑/↓, **Esc** steuern die Präsentation – auch wenn die
     Seite gerade den Fokus hat.

Weitere Buttons: **Datei ändern**, **Link entfernen**, **Neu laden**, **Anleitung**, **Info**.

## Aus dem Quellcode bauen (für Entwickler)

Voraussetzungen: **.NET SDK 8/10**, Windows, PowerPoint x64, WebView2-Runtime.

```powershell
# Bauen + pro Benutzer registrieren (PowerPoint vorher schließen):
./scripts/build.ps1

# Verteilbares Paket nach dist/ erstellen:
./scripts/package.ps1

# Schneller Lade-Test (startet PowerPoint, prüft Diagnose-Beacons):
./scripts/test-load.ps1
```

Diagnose-Log: `%LOCALAPPDATA%\LiveWebRegion\addin.log`.

### Architektur (Kurzfassung)

- **Managed COM-Add-in** (C# / .NET Framework `net48`, x64), kein VSTO, kein Office.js.
- Lifecycle über `IDTExtensibility2`, UI über `IRibbonExtensibility` (Ribbon-XML).
- Anzeige: **WebView2** in einem randlosen, top-most WinForms-Overlay, das im Slideshow per
  eigener Punkt→Pixel-Umrechnung exakt über die getaggte Form gelegt wird.
- Bereich = Form mit Tag `LiveWebPath` (`Shape.Tags`).
- Tastatur im Slideshow: WebView2 `AcceleratorKeyPressed` reicht Navigationstasten an die Show.

Wichtige Implementierungsdetails (z. B. warum `UseWindowsForms` vermieden wird und Interface-
Pointer als `IntPtr` entgegengenommen werden) stehen als Kommentare im Quellcode.

## Lizenz

© 2026 Michael Burtscher. Alle Rechte vorbehalten. Siehe [LICENSE.txt](LICENSE.txt)
(proprietär; persönliche/Evaluierungs-Nutzung). Für kommerzielle Lizenzierung den Autor
kontaktieren.
