# sharepointcli

Kleines, eigenständiges Kommandozeilenwerkzeug (Go, eine Binärdatei, kein CGO), mit dem die
Harness eines Mitarbeiters (Claude Code, Codex o. ä.) **mit dem eigenen Microsoft-Login des
Mitarbeiters** über Microsoft Graph lesend auf SharePoint zugreift.

- **Delegierte Rechte:** Das Tool sieht genau das, was der angemeldete Mitarbeiter in SharePoint
  auch im Browser sehen darf – nicht mehr. Es gibt kein Dienstkonto und kein App-Secret.
- **Nur lesend** (Sites.Read.All, Files.Read.All).
- Ausgabe als Text oder mit `--json` maschinenlesbar.

> **Status 0.1.0:** Unit-Tests (Referenz-Parsing, Share-URL-Kodierung, docx/pptx/xlsx-Extraktion,
> Token-Cache) laufen grün. **Gegen einen echten Microsoft-365-Tenant ist das Tool noch zu
> verifizieren** – bisher stand kein Live-Tenant zur Verfügung.

## Installation

1. Passendes Archiv aus dem Release herunterladen:
   - Windows: `sharepointcli_0.1.0_windows_amd64.zip` (bzw. `_arm64`)
   - macOS: `sharepointcli_0.1.0_darwin_arm64.tar.gz` (Apple Silicon) / `_darwin_amd64` (Intel)
   - Linux: `sharepointcli_0.1.0_linux_amd64.tar.gz` / `_linux_arm64`
2. Prüfsumme vergleichen: `sha256sum -c SHA256SUMS --ignore-missing`
   (Windows: `Get-FileHash .\sharepointcli_0.1.0_windows_amd64.zip`)
3. Entpacken und in den `PATH` legen:
   - Linux/macOS: `tar xzf sharepointcli_*.tar.gz && mkdir -p ~/.local/bin && mv sharepointcli ~/.local/bin/`
     (macOS ggf. Quarantäne entfernen: `xattr -d com.apple.quarantine ~/.local/bin/sharepointcli`)
   - Windows: `sharepointcli.exe` z. B. nach `%LOCALAPPDATA%\Programs\sharepointcli\` kopieren und
     diesen Ordner über „Umgebungsvariablen für dieses Konto bearbeiten“ → `Path` hinzufügen.
4. Testen: `sharepointcli version`

## Einrichtung beim Kunden (einmalig, durch Kunden-Admin)

Das Tool braucht eine App-Registrierung im Entra ID (Azure AD) des Kunden. Sie enthält **kein
Secret**; sie legt nur fest, welche delegierten Rechte Mitarbeiter dem Tool geben dürfen.

1. <https://entra.microsoft.com> → **Identität → Anwendungen → App-Registrierungen → Neue Registrierung**.
2. Name: z. B. `sharepointcli`. Unterstützte Kontotypen: **Nur Konten in diesem Organisationsverzeichnis**
   (Single Tenant).
3. Umleitungs-URI: Plattform **Öffentlicher Client/nativ (mobil und Desktop)**, URI `http://localhost`.
   → **Registrieren**.
4. Auf der Übersichtsseite notieren:
   - **Anwendungs-ID (Client)** → später `--client-id`
   - **Verzeichnis-ID (Mandant)** → später `--tenant` (alternativ die Domain, z. B. `contoso.onmicrosoft.com`)
5. **Authentifizierung** → „Erweiterte Einstellungen“ → **Öffentliche Clientflows zulassen: Ja**
   (nötig für `--device-code`). Speichern.
6. **API-Berechtigungen → Berechtigung hinzufügen → Microsoft Graph → Delegierte Berechtigungen**:
   - `Sites.Read.All`
   - `Files.Read.All`
   - `offline_access`
   - `User.Read` (ist meist schon vorhanden; nötig für `whoami`)
   Keine *Anwendungsberechtigungen* hinzufügen.
7. **Administratorzustimmung für <Mandant> erteilen** klicken und bestätigen. Danach steht bei
   allen Rechten „Gewährt für <Mandant>“. Ohne diesen Schritt sehen Mitarbeiter beim Login
   „Administratorgenehmigung erforderlich“, da `Sites.Read.All`/`Files.Read.All` admin-pflichtig sind.
8. Client-ID und Tenant an die Mitarbeiter weitergeben (beides ist nicht geheim).

Optional: Unter **Unternehmensanwendungen → sharepointcli → Eigenschaften** „Zuweisung erforderlich“
aktivieren und nur bestimmte Benutzer/Gruppen zuweisen.

Hinweis: Die Rechte sind „*.Read.All“-Scopes, aber delegiert – Graph liefert trotzdem nur, was der
angemeldete Benutzer selbst sehen darf.

## Login

```bash
sharepointcli login --tenant contoso.onmicrosoft.com --client-id 11111111-2222-3333-4444-555555555555
```

- Öffnet den System-Browser, der Login landet über `http://localhost:<zufallsport>` wieder beim Tool.
- Ohne Browser (SSH, Container): `sharepointcli login --device-code` – Code auf
  <https://microsoft.com/devicelogin> eingeben.
- Tenant/Client-ID werden in `config.json` gespeichert; danach reicht `sharepointcli login`.
  Alternativ Umgebungsvariablen `SHAREPOINTCLI_TENANT` und `SHAREPOINTCLI_CLIENT_ID`.
- Alle weiteren Befehle holen das Token still aus dem Cache und erneuern es automatisch
  (Refresh-Token). Läuft es ab oder wird es widerrufen: erneut `sharepointcli login`.

## Befehle

Alle Befehle verstehen `--json` und `--help`.

| Befehl | Zweck |
|---|---|
| `login [--tenant] [--client-id] [--device-code]` | anmelden |
| `logout` | Token-Cache löschen |
| `whoami` | Name und UPN des angemeldeten Benutzers |
| `sites [suche]` | Sites suchen (Name, webUrl, id) |
| `drives <site-id>` | Dokumentbibliotheken einer Site |
| `ls <ref>` | Ordnerinhalt |
| `get <ref> [-o datei]` | Datei herunterladen (`-o -` = stdout) |
| `text <ref> [--max-chars N]` | Text extrahieren (Default 20000 Zeichen, 0 = unbegrenzt) |
| `resolve <url>` | SharePoint-URL → site/drive/item |
| `search <site-id> "<begriff>"` | Suche in der Standardbibliothek einer Site |
| `version` | Version |

### Referenzen

Ein Element wird so adressiert:

- `site=<site-id> drive=<drive-id> item=<item-id>` – als ein Argument (in Anführungszeichen) oder
  mehrere Argumente; `drive` + `item` genügt, `site` + `item` nutzt die Standardbibliothek.
- `path=/Ordner/Datei.docx` statt `item=` (relativ zur Bibliothekswurzel).
- Einzelne Flags: `--site`, `--drive`, `--item`, `--path`.
- `--url <sharepoint-url>` – beliebiger SharePoint-/OneDrive-Link (auch Freigabelinks); wird über
  `/shares/u!<base64url>/driveItem` aufgelöst.

### Beispiele

```bash
sharepointcli whoami
sharepointcli sites Vertrieb
sharepointcli sites --json
sharepointcli drives "contoso.sharepoint.com,2c7d…,9a1b…"
sharepointcli ls site="contoso.sharepoint.com,2c7d…,9a1b…"
sharepointcli ls drive=b!AbC… path=/Angebote/2026
sharepointcli ls --url "https://contoso.sharepoint.com/sites/Vertrieb/Shared%20Documents/Angebote"
sharepointcli resolve "https://contoso.sharepoint.com/:w:/r/sites/Vertrieb/Shared%20Documents/Angebot.docx"
sharepointcli get "drive=b!AbC… item=01XYZ…" -o angebot.docx
sharepointcli text --url "https://contoso.sharepoint.com/:p:/r/sites/Vertrieb/…/Pitch.pptx" --max-chars 5000
sharepointcli text drive=b!AbC… path=/Protokolle/2026-10.xlsx --json
sharepointcli search "contoso.sharepoint.com,2c7d…,9a1b…" "Rahmenvertrag"
```

`ls`, `search` geben pro Eintrag eine fertige Referenz (`drive=… item=…`) aus, die direkt an
`get`/`text` übergeben werden kann.

### Textextraktion

- `.docx`, `.pptx`, `.xlsx`: direkt aus dem Office-XML (nur Go-Standardbibliothek). pptx wird
  folienweise, xlsx blattweise (Tab-getrennt) ausgegeben.
- `.txt`, `.md`, `.csv` (und ähnliche Textformate): unverändert.
- `.pdf`: **in 0.1 nicht unterstützt** – Datei mit `get` laden.

### Fehlermeldungen

- **403:** „Kein Zugriff mit deinem Microsoft-Login (das ist kein Fehler des Tools)“ – der Benutzer hat
  auf das Element keine Rechte.
- **404 bei einem Element:** Item-IDs ändern sich, wenn Dateien über Bibliotheksgrenzen verschoben
  werden; mit `--url` neu auflösen.
- **401 / abgelaufen:** erneut `sharepointcli login`.

## Sicherheit

- Kein Secret, kein Dienstkonto: Zugriff ausschließlich mit dem Login des Mitarbeiters (OAuth 2.0
  Authorization Code + PKCE bzw. Device Code, über MSAL Go).
- Nur Leserechte angefragt; das Tool schreibt nichts nach SharePoint.
- Token-Cache: `<Benutzer-Konfigurationsordner>/sharepointcli/token_cache.json` mit Dateirechten 0600
  (Linux: `~/.config/sharepointcli/`, macOS: `~/Library/Application Support/sharepointcli/`,
  Windows: `%AppData%\sharepointcli\`). Er enthält Access- und Refresh-Token – wie ein Passwort
  behandeln. Unter Windows schützt die Benutzerprofil-ACL die Datei, nicht die Unix-Rechte.
  `sharepointcli logout` löscht ihn.
- Admins können den Zugriff jederzeit entziehen (App-Zuweisung entfernen, Sitzungen widerrufen,
  Unternehmensanwendung deaktivieren) und Anmeldungen im Entra-Anmeldeprotokoll nachvollziehen.
- Hinweis für Harness-Nutzung: Ausgaben von `text`/`get` landen im Kontext des KI-Modells. Nur
  Inhalte abrufen, die dort verarbeitet werden dürfen.

## Selbst bauen

Voraussetzung: Go ≥ 1.24.

```bash
CGO_ENABLED=0 go build -o sharepointcli .
go vet ./... && go test ./...
./build.sh 0.1.0      # alle Plattformen nach dist/, inkl. dist/SHA256SUMS
```

Einzige externe Abhängigkeit: `github.com/AzureAD/microsoft-authentication-library-for-go`.

## Lizenz

MIT, siehe [LICENSE](LICENSE).
