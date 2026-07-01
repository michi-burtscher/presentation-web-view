<div align="center">
  <img src="assets/logo.svg" width="84" alt="Live Web Region" />
  <h1>Live Web Region für PowerPoint</h1>
  <p><b>Zeigt eine lokale HTML/JS/CSS-Datei live in einem Folienbereich an – interaktiv, auch im Präsentationsmodus.</b></p>
</div>

### ⬇️ Download / Installation (ein Klick)

**[➡️ Installer herunterladen (LiveWebRegionSetup.exe)](https://github.com/michi-burtscher/presentation-web-view/releases/latest/download/LiveWebRegionSetup.exe)**
— danach doppelklicken, fertig (pro Benutzer, ohne Adminrechte).
Alle Versionen: **[Releases](https://github.com/michi-burtscher/presentation-web-view/releases/latest)**.

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
- ✅ Mehrere Frames pro Folie · Multi-Monitor · Presenter-View
- ✅ Pfeiltasten / Bild↑↓ / Esc steuern die Präsentation weiter – auch während die Seite Fokus hat
- ✅ Pro Frame: Interaktiv/klick-durch, Stummschalten, Zoom, Auto-Neuladen
- ✅ Portabel: relative Pfade **oder** HTML in die `.pptx` einbetten
- ✅ Bearbeiten über „Optionen“ · Fehler-/Offline-Fallback · In-App-Update

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

1. **Fenster erstellen** (Reiter *Live Web*): fügt einen *Live Web Frame* (Karte) ein und
   öffnet den Dialog – **URL eingeben oder HTML-Datei wählen**, plus Optionen.
2. Frame nach Wunsch positionieren/skalieren. Bearbeiten per **Doppelklick** oder **Optionen**.
3. **F5** – die Seite läuft live im Frame.
   - **Interaktion:** Klicks/Tippen gehen an die Seite (abschaltbar = klick-durch).
   - **Navigation:** Pfeiltasten, Bild↑/↓, **Esc** steuern die Präsentation – auch wenn die
     Seite gerade den Fokus hat.

Pro Frame einstellbar (Dialog): Interaktiv, Stummschalten, Zoom %, Auto-Neuladen, Einbetten.
Weitere Buttons: **Fenster entfernen**, **Neu laden**, **Update**, **Anleitung**, **Info**.

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

Freie Software unter der **MIT-Lizenz** © 2026 Michael Burtscher – siehe [LICENSE](LICENSE).
Kostenlos nutzbar, weitergebbar und veränderbar, ohne Gewährleistung.
