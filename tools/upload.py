"""GUI helper to publish Live Web Region releases to GitHub.

Run from the repository root:

    python tools/upload.py

Requirements:
  * git and the GitHub CLI ``gh`` on PATH, authenticated (``gh auth login``).
  * PowerShell (for the build via scripts/package.ps1).

Typical order (top to bottom): set version -> build -> write sha256 ->
create/upload GitHub release -> commit & push. Normal users never need this;
they only read the public update.json manifest and download the release asset.
"""

from __future__ import annotations

import hashlib
import json
import re
import shutil
import subprocess
import threading
import tkinter as tk
from pathlib import Path
from tkinter import scrolledtext, messagebox


def find_gh() -> str:
    """Locate the GitHub CLI even if it is not yet on this session's PATH."""
    found = shutil.which("gh")
    if found:
        return found
    for candidate in (
        r"C:\Program Files\GitHub CLI\gh.exe",
        r"C:\Program Files (x86)\GitHub CLI\gh.exe",
    ):
        if Path(candidate).exists():
            return candidate
    return "gh"  # fall back; will error clearly if missing


GH = find_gh()

ROOT = Path(__file__).resolve().parents[1]
OWNER = "michi-burtscher"
REPO = "presentation-web-view"
APPINFO_CS = ROOT / "src" / "LiveWebRegion" / "AppInfo.cs"
UPDATE_JSON = ROOT / "update.json"
SETUP_EXE = ROOT / "dist" / "LiveWebRegion" / "LiveWebRegionSetup.exe"
PACKAGE_PS1 = ROOT / "scripts" / "package.ps1"


def normalize_version(v: str) -> str:
    v = (v or "").strip()
    if not v:
        raise ValueError("Version ist leer.")
    if not v.startswith("v"):
        v = "v" + v
    if not re.fullmatch(r"v\d+(?:[._-][A-Za-z0-9]+)?", v):
        raise ValueError("Nutze Versionen wie v01, v02 oder v02-test.")
    return v


def exe_url(version: str) -> str:
    return f"https://github.com/{OWNER}/{REPO}/releases/download/{version}/LiveWebRegionSetup.exe"


def read_version() -> str:
    try:
        m = re.search(r'Version\s*=\s*"([^"]+)"', APPINFO_CS.read_text(encoding="utf-8"))
        if m:
            return m.group(1)
    except Exception:
        pass
    return "v01"


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        root.title("Live Web Region – Release Helper")
        root.geometry("820x620")

        top = tk.Frame(root)
        top.pack(fill="x", padx=10, pady=8)
        tk.Label(top, text="Version:").grid(row=0, column=0, sticky="w")
        self.version = tk.Entry(top, width=14)
        self.version.insert(0, read_version())
        self.version.grid(row=0, column=1, sticky="w", padx=(4, 20))
        tk.Label(top, text=f"Repo: {OWNER}/{REPO}").grid(row=0, column=2, sticky="w")

        tk.Label(top, text="Release-Notes (eine Zeile pro Punkt):").grid(row=1, column=0, columnspan=3, sticky="w", pady=(8, 0))
        self.notes = tk.Text(top, height=5, width=100)
        self.notes.grid(row=2, column=0, columnspan=3, sticky="we", pady=(2, 0))
        self.notes.insert("1.0", "Neue Version.")

        steps = tk.Frame(root)
        steps.pack(fill="x", padx=10, pady=6)
        buttons = [
            ("1 · Version setzen (AppInfo + update.json)", self.set_version),
            ("2 · Build (package.ps1)", self.build),
            ("3 · SHA-256 in update.json schreiben", self.write_sha),
            ("4 · GitHub Release + EXE hochladen", self.release),
            ("5 · Commit + Push (main)", self.commit_push),
            ("Alles nacheinander (1–5)", self.all_in_one),
        ]
        for i, (label, cb) in enumerate(buttons):
            b = tk.Button(steps, text=label, width=42, command=lambda c=cb: self.run_async(c))
            b.grid(row=i // 2, column=i % 2, padx=4, pady=4, sticky="we")
        self._step_buttons = steps.winfo_children()

        self.log = scrolledtext.ScrolledText(root, height=22)
        self.log.pack(fill="both", expand=True, padx=10, pady=(4, 10))
        self.busy = False

    # ---- helpers ----
    def out(self, text: str):
        self.log.insert("end", text.rstrip() + "\n")
        self.log.see("end")
        self.root.update_idletasks()

    def run(self, args, cwd=ROOT) -> str:
        self.out("$ " + " ".join(args))
        p = subprocess.run(args, cwd=cwd, text=True, capture_output=True, encoding="utf-8", errors="replace")
        output = (p.stdout or "") + (p.stderr or "")
        self.out(output.strip())
        if p.returncode != 0:
            raise RuntimeError(f"Befehl fehlgeschlagen (Code {p.returncode}): {' '.join(args)}")
        return output.strip()

    def run_async(self, func):
        if self.busy:
            messagebox.showinfo("Release", "Es läuft bereits eine Aktion.")
            return
        self.busy = True
        for b in self._step_buttons:
            b.config(state="disabled")

        def worker():
            try:
                func()
                self.out("== OK ==")
            except Exception as exc:
                self.out("FEHLER: " + str(exc))
                messagebox.showerror("Release", str(exc))
            finally:
                self.busy = False
                for b in self._step_buttons:
                    b.config(state="normal")

        threading.Thread(target=worker, daemon=True).start()

    def v(self) -> str:
        return normalize_version(self.version.get())

    def notes_list(self) -> list[str]:
        lines = [ln.strip(" -\t") for ln in self.notes.get("1.0", "end").splitlines()]
        return [ln for ln in lines if ln] or ["Neue Version."]

    # ---- steps ----
    def set_version(self):
        version = self.v()
        text = APPINFO_CS.read_text(encoding="utf-8")
        text, n = re.subn(r'(Version\s*=\s*")[^"]+(")', rf'\g<1>{version}\g<2>', text, count=1)
        if n != 1:
            raise RuntimeError("Version in AppInfo.cs nicht gefunden.")
        APPINFO_CS.write_text(text, encoding="utf-8")

        data = {
            "version": version,
            "notes": self.notes_list(),
            "exe_url": exe_url(version),
            "sha256": "",
            "page_url": f"https://github.com/{OWNER}/{REPO}/releases",
        }
        UPDATE_JSON.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        self.out(f"Version {version} in AppInfo.cs und update.json gesetzt.")

    def build(self):
        self.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(PACKAGE_PS1)])
        if not SETUP_EXE.exists():
            raise RuntimeError(f"Setup-EXE nicht gefunden: {SETUP_EXE}")
        self.out(f"Gebaut: {SETUP_EXE}")

    def write_sha(self):
        if not SETUP_EXE.exists():
            raise RuntimeError("Bitte zuerst bauen (Schritt 2).")
        sha = hashlib.sha256(SETUP_EXE.read_bytes()).hexdigest()
        data = json.loads(UPDATE_JSON.read_text(encoding="utf-8-sig"))
        data["sha256"] = sha
        data["exe_url"] = exe_url(self.v())
        UPDATE_JSON.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        self.out("SHA-256: " + sha)

    def release(self):
        version = self.v()
        if not SETUP_EXE.exists():
            raise RuntimeError("Bitte zuerst bauen (Schritt 2).")
        notes = "\n".join("- " + n for n in self.notes_list())
        asset = f"{SETUP_EXE}#LiveWebRegionSetup.exe"
        try:
            self.run([GH, "release", "create", version, asset, "--title", f"Live Web Region {version}", "--notes", notes])
        except RuntimeError:
            self.out("Release existiert evtl. schon – lade Asset mit --clobber hoch.")
            self.run([GH, "release", "upload", version, asset, "--clobber"])

    def commit_push(self):
        version = self.v()
        self.run(["git", "add", "-A"])
        # commit may be empty if nothing changed; ignore that specific failure
        p = subprocess.run(["git", "commit", "-m", f"release: {version}"], cwd=ROOT, text=True,
                           capture_output=True, encoding="utf-8", errors="replace")
        self.out((p.stdout or "") + (p.stderr or ""))
        self.run(["git", "push", "origin", "main"])

    def all_in_one(self):
        self.set_version()
        self.build()
        self.write_sha()
        self.release()
        self.commit_push()


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
