// sharepointcli: SharePoint-Zugriff über Microsoft Graph mit dem eigenen
// Microsoft-Login (delegierte Rechte).
package main

import (
	"context"
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"net/url"
	"os"
	"os/signal"
	"path/filepath"
	"strconv"
	"strings"
	"text/tabwriter"
)

var version = "dev"

const usage = `sharepointcli – SharePoint über Microsoft Graph mit deinem eigenen Microsoft-Login

Aufruf:
  sharepointcli <befehl> [argumente] [--json]

Befehle:
  login      Anmelden (Browser; --device-code ohne Browser)
  logout     Token-Cache löschen
  whoami     Angemeldeten Benutzer anzeigen
  sites      Sites suchen
  drives     Bibliotheken einer Site
  ls         Ordnerinhalt auflisten
  get        Datei herunterladen
  text       Text aus Datei extrahieren (docx/pptx/xlsx/txt/md/csv)
  resolve    SharePoint-URL in Referenz (site/drive/item) auflösen
  search     In der Standardbibliothek einer Site suchen
  version    Version anzeigen

Referenz (ls/get/text):
  site=<site-id> drive=<drive-id> item=<item-id>   (auch path=/Ordner/Datei.docx)
  oder einzeln: --site <id> --drive <id> --item <id> --path </pfad>
  oder: --url <sharepoint-url>

Globale Flags:
  --json     maschinenlesbare Ausgabe
  -h,--help  Hilfe (auch: sharepointcli <befehl> --help)
`

var cmdHelp = map[string]string{
	"login": `sharepointcli login [--tenant <tenant-id|domain>] [--client-id <app-id>] [--device-code]
  Meldet dich mit deinem Microsoft-Konto an. Tenant und Client-ID werden gespeichert.
Beispiele:
  sharepointcli login --tenant contoso.onmicrosoft.com --client-id 11111111-2222-3333-4444-555555555555
  sharepointcli login --device-code
  (Alternativ: Umgebungsvariablen SHAREPOINTCLI_TENANT, SHAREPOINTCLI_CLIENT_ID)`,
	"logout": `sharepointcli logout
  Löscht den lokalen Token-Cache.
Beispiel:
  sharepointcli logout`,
	"whoami": `sharepointcli whoami [--json]
  Zeigt Name und UPN des angemeldeten Benutzers.
Beispiel:
  sharepointcli whoami --json`,
	"sites": `sharepointcli sites [suche] [--json]
  Sucht Sites, auf die du Zugriff hast. Ohne Suchbegriff: alle (*).
Beispiele:
  sharepointcli sites Vertrieb
  sharepointcli sites --json`,
	"drives": `sharepointcli drives <site-id> [--json]
  Listet die Dokumentbibliotheken einer Site.
Beispiel:
  sharepointcli drives contoso.sharepoint.com,1234...,5678...`,
	"ls": `sharepointcli ls <ref> [--json]
  Listet den Inhalt eines Ordners (ohne item/path: Wurzel der Bibliothek).
Beispiele:
  sharepointcli ls site=<site-id>
  sharepointcli ls drive=<drive-id> path=/Projekte/2026
  sharepointcli ls "drive=<drive-id> item=<item-id>"
  sharepointcli ls --url "https://contoso.sharepoint.com/sites/Vertrieb/Shared%20Documents/Angebote"`,
	"get": `sharepointcli get <ref> [-o datei] [--json]
  Lädt eine Datei herunter (Standard: Dateiname im aktuellen Verzeichnis; -o - = stdout).
Beispiele:
  sharepointcli get drive=<drive-id> item=<item-id>
  sharepointcli get --url "https://contoso.sharepoint.com/:w:/r/sites/Vertrieb/..." -o angebot.docx`,
	"text": `sharepointcli text <ref> [--max-chars N] [--json]
  Extrahiert Text aus .docx/.pptx/.xlsx/.txt/.md/.csv (Default --max-chars 20000, 0 = unbegrenzt).
  PDF wird in 0.1 nicht unterstützt (Datei mit get laden).
Beispiele:
  sharepointcli text --url "https://contoso.sharepoint.com/:w:/r/sites/Vertrieb/..."
  sharepointcli text drive=<drive-id> path=/Protokolle/2026-10.docx --max-chars 5000`,
	"resolve": `sharepointcli resolve <url> [--json]
  Löst eine SharePoint-/OneDrive-URL in site/drive/item auf (über /shares).
Beispiel:
  sharepointcli resolve "https://contoso.sharepoint.com/sites/Vertrieb/Shared%20Documents/Angebot.docx"`,
	"search": `sharepointcli search <site-id> "<begriff>" [--json]
  Sucht in der Standardbibliothek der Site.
Beispiel:
  sharepointcli search <site-id> "Rahmenvertrag"`,
	"version": `sharepointcli version
Beispiel:
  sharepointcli version`,
}

// opts: geparste Flags. Flags dürfen vor und nach Positionsargumenten stehen.
type opts struct {
	json, deviceCode, help bool
	tenant, clientID       string
	site, drive, item      string
	path, url, out         string
	maxChars               int
	args                   []string
}

func parseArgs(argv []string) (opts, error) {
	o := opts{maxChars: 20000}
	valueFlags := map[string]*string{
		"tenant": &o.tenant, "client-id": &o.clientID, "site": &o.site, "drive": &o.drive,
		"item": &o.item, "path": &o.path, "url": &o.url, "o": &o.out, "output": &o.out,
	}
	var maxChars string
	for i := 0; i < len(argv); i++ {
		a := argv[i]
		if a == "--" {
			o.args = append(o.args, argv[i+1:]...)
			break
		}
		if !strings.HasPrefix(a, "-") || a == "-" {
			o.args = append(o.args, a)
			continue
		}
		name := strings.TrimLeft(a, "-")
		val, hasVal := "", false
		if k, v, ok := strings.Cut(name, "="); ok {
			name, val, hasVal = k, v, true
		}
		switch name {
		case "json":
			o.json = true
			continue
		case "device-code":
			o.deviceCode = true
			continue
		case "h", "help":
			o.help = true
			continue
		}
		dst, ok := valueFlags[name]
		if name == "max-chars" {
			dst, ok = &maxChars, true
		}
		if !ok {
			return o, fmt.Errorf("unbekanntes Flag %q (siehe --help)", a)
		}
		if !hasVal {
			if i+1 >= len(argv) {
				return o, fmt.Errorf("Flag %q braucht einen Wert", a)
			}
			i++
			val = argv[i]
		}
		*dst = val
	}
	if maxChars != "" {
		n, err := strconv.Atoi(maxChars)
		if err != nil || n < 0 {
			return o, fmt.Errorf("--max-chars erwartet eine Zahl ≥ 0")
		}
		o.maxChars = n
	}
	return o, nil
}

func main() {
	ctx, stop := signal.NotifyContext(context.Background(), os.Interrupt)
	defer stop()
	if err := run(ctx, os.Args[1:], os.Stdout); err != nil {
		fmt.Fprintln(os.Stderr, "Fehler:", err)
		os.Exit(1)
	}
}

func run(ctx context.Context, argv []string, w io.Writer) error {
	if len(argv) == 0 {
		fmt.Fprint(w, usage)
		return nil
	}
	cmd := argv[0]
	if cmd == "-h" || cmd == "--help" || cmd == "help" {
		if len(argv) > 1 && cmdHelp[argv[1]] != "" {
			fmt.Fprintln(w, cmdHelp[argv[1]])
			return nil
		}
		fmt.Fprint(w, usage)
		return nil
	}
	if cmd == "--version" {
		cmd = "version"
	}
	o, err := parseArgs(argv[1:])
	if err != nil {
		return err
	}
	if o.help {
		if h, ok := cmdHelp[cmd]; ok {
			fmt.Fprintln(w, h)
			return nil
		}
		fmt.Fprint(w, usage)
		return nil
	}
	switch cmd {
	case "version":
		if o.json {
			return emitJSON(w, map[string]string{"version": version})
		}
		fmt.Fprintln(w, "sharepointcli", version)
		return nil
	case "login":
		return cmdLogin(ctx, w, o)
	case "logout":
		if err := logout(ctx); err != nil {
			return err
		}
		fmt.Fprintln(w, "Abgemeldet (Token-Cache gelöscht).")
		return nil
	case "whoami":
		return cmdWhoami(ctx, w, o)
	case "sites":
		return cmdSites(ctx, w, o)
	case "drives":
		return cmdDrives(ctx, w, o)
	case "ls":
		return cmdLs(ctx, w, o)
	case "get":
		return cmdGet(ctx, w, o)
	case "text":
		return cmdText(ctx, w, o)
	case "resolve":
		return cmdResolve(ctx, w, o)
	case "search":
		return cmdSearch(ctx, w, o)
	}
	return fmt.Errorf("unbekannter Befehl %q (siehe sharepointcli --help)", cmd)
}

func emitJSON(w io.Writer, v any) error {
	enc := json.NewEncoder(w)
	enc.SetIndent("", "  ")
	enc.SetEscapeHTML(false)
	return enc.Encode(v)
}

func table(w io.Writer, header []string, rows [][]string) {
	tw := tabwriter.NewWriter(w, 0, 0, 2, ' ', 0)
	fmt.Fprintln(tw, strings.Join(header, "\t"))
	for _, r := range rows {
		fmt.Fprintln(tw, strings.Join(r, "\t"))
	}
	tw.Flush()
}

func cmdLogin(ctx context.Context, w io.Writer, o opts) error {
	c, err := resolveConfig(o.tenant, o.clientID)
	if err != nil {
		return err
	}
	res, err := login(ctx, c, o.deviceCode)
	if err != nil {
		if !o.deviceCode {
			return fmt.Errorf("Login fehlgeschlagen: %v\nOhne Browser (z. B. per SSH): sharepointcli login --device-code", err)
		}
		return fmt.Errorf("Login fehlgeschlagen: %v", err)
	}
	if err := saveConfig(c); err != nil {
		return err
	}
	if o.json {
		return emitJSON(w, map[string]string{"user": res.Account.PreferredUsername, "tenant": c.Tenant, "client_id": c.ClientID})
	}
	fmt.Fprintf(w, "Angemeldet als %s\n", res.Account.PreferredUsername)
	return nil
}

func cmdWhoami(ctx context.Context, w io.Writer, o opts) error {
	g, err := newGraph(ctx)
	if err != nil {
		return err
	}
	var me struct {
		DisplayName       string `json:"displayName"`
		UserPrincipalName string `json:"userPrincipalName"`
		Mail              string `json:"mail"`
		ID                string `json:"id"`
	}
	if err := g.getJSON(ctx, "/me?$select=id,displayName,userPrincipalName,mail", false, &me); err != nil {
		return err
	}
	if o.json {
		return emitJSON(w, me)
	}
	fmt.Fprintf(w, "Name: %s\nUPN:  %s\n", me.DisplayName, me.UserPrincipalName)
	return nil
}

func cmdSites(ctx context.Context, w io.Writer, o opts) error {
	q := strings.Join(o.args, " ")
	if q == "" {
		q = "*"
	}
	g, err := newGraph(ctx)
	if err != nil {
		return err
	}
	sites, err := getAll[site](ctx, g, "/sites?search="+url.QueryEscape(q)+"&$select=id,name,displayName,webUrl", false, 500)
	if err != nil {
		return err
	}
	if o.json {
		return emitJSON(w, nonNil(sites))
	}
	rows := make([][]string, 0, len(sites))
	for _, s := range sites {
		n := s.DisplayName
		if n == "" {
			n = s.Name
		}
		rows = append(rows, []string{n, s.WebURL, s.ID})
	}
	table(w, []string{"NAME", "WEBURL", "ID"}, rows)
	return nil
}

func cmdDrives(ctx context.Context, w io.Writer, o opts) error {
	siteID := o.site
	if len(o.args) > 0 {
		siteID = strings.TrimPrefix(o.args[0], "site=")
	}
	if siteID == "" {
		return errors.New("site-id fehlt: sharepointcli drives <site-id>")
	}
	g, err := newGraph(ctx)
	if err != nil {
		return err
	}
	drives, err := getAll[drive](ctx, g, "/sites/"+url.PathEscape(siteID)+"/drives?$select=id,name,driveType,webUrl", false, 500)
	if err != nil {
		return err
	}
	if o.json {
		return emitJSON(w, nonNil(drives))
	}
	rows := make([][]string, 0, len(drives))
	for _, d := range drives {
		rows = append(rows, []string{d.Name, d.WebURL, d.ID})
	}
	table(w, []string{"NAME", "WEBURL", "ID"}, rows)
	return nil
}

func nonNil[T any](s []T) []T {
	if s == nil {
		return []T{}
	}
	return s
}

// refFromOpts kombiniert Positionsargumente und Einzel-Flags; URLs werden
// über /shares aufgelöst.
func refFromOpts(ctx context.Context, g *graphClient, o opts) (itemRef, error) {
	r, err := parseRef(o.args)
	if err != nil {
		return r, err
	}
	r = r.merge(itemRef{Site: o.site, Drive: o.drive, Item: o.item, Path: o.path, URL: o.url})
	if r.URL != "" {
		it, err := resolveURL(ctx, g, r.URL)
		if err != nil {
			return r, err
		}
		return itemRef{Site: it.ParentReference.SiteID, Drive: it.ParentReference.DriveID, Item: it.ID}, nil
	}
	if r.Site == "" && r.Drive == "" {
		return r, errors.New("Referenz fehlt: site=<id> drive=<id> item=<id> oder --url <sharepoint-url>")
	}
	return r, nil
}

func resolveURL(ctx context.Context, g *graphClient, u string) (driveItem, error) {
	var it driveItem
	err := g.getJSON(ctx, "/shares/"+encodeShareURL(u)+"/driveItem", false, &it)
	var ge *graphError
	if errors.As(err, &ge) && ge.Status == 404 {
		return it, errors.New("URL nicht auflösbar (404) – existiert die Datei noch und ist es eine SharePoint-/OneDrive-URL?")
	}
	return it, err
}

func cmdResolve(ctx context.Context, w io.Writer, o opts) error {
	u := o.url
	if len(o.args) > 0 {
		u = o.args[0]
	}
	if u == "" {
		return errors.New("URL fehlt: sharepointcli resolve <url>")
	}
	g, err := newGraph(ctx)
	if err != nil {
		return err
	}
	it, err := resolveURL(ctx, g, u)
	if err != nil {
		return err
	}
	ref := itemRef{Site: it.ParentReference.SiteID, Drive: it.ParentReference.DriveID, Item: it.ID}
	if o.json {
		return emitJSON(w, map[string]any{"site": ref.Site, "drive": ref.Drive, "item": ref.Item, "name": it.Name, "webUrl": it.WebURL, "folder": it.Folder != nil, "ref": ref.String()})
	}
	fmt.Fprintln(w, ref.String())
	fmt.Fprintf(w, "# %s\n", it.Name)
	return nil
}

func kind(it driveItem) string {
	if it.Folder != nil {
		return "ordner"
	}
	return "datei"
}

func printItems(w io.Writer, o opts, items []driveItem) error {
	if o.json {
		return emitJSON(w, nonNil(items))
	}
	rows := make([][]string, 0, len(items))
	for _, it := range items {
		size := strconv.FormatInt(it.Size, 10)
		if it.Folder != nil {
			size = "-"
		}
		ref := itemRef{Drive: it.ParentReference.DriveID, Item: it.ID}.String()
		rows = append(rows, []string{kind(it), it.Name, size, it.LastModifiedDateTime, ref})
	}
	table(w, []string{"TYP", "NAME", "GRÖSSE", "GEÄNDERT", "REF"}, rows)
	return nil
}

const itemSelect = "$select=id,name,size,webUrl,lastModifiedDateTime,folder,file,parentReference"

func cmdLs(ctx context.Context, w io.Writer, o opts) error {
	g, err := newGraph(ctx)
	if err != nil {
		return err
	}
	r, err := refFromOpts(ctx, g, o)
	if err != nil {
		return err
	}
	p, err := r.childrenPath()
	if err != nil {
		return err
	}
	items, err := getAll[driveItem](ctx, g, p+"?"+itemSelect+"&$top=200", r.Item != "" || r.Path != "", 5000)
	if err != nil {
		return err
	}
	return printItems(w, o, items)
}

func cmdSearch(ctx context.Context, w io.Writer, o opts) error {
	if len(o.args) < 2 {
		return errors.New(`Aufruf: sharepointcli search <site-id> "<begriff>"`)
	}
	siteID := strings.TrimPrefix(o.args[0], "site=")
	term := strings.ReplaceAll(strings.Join(o.args[1:], " "), "'", "''")
	g, err := newGraph(ctx)
	if err != nil {
		return err
	}
	p := "/sites/" + url.PathEscape(siteID) + "/drive/root/search(q='" + url.PathEscape(term) + "')?" + itemSelect
	items, err := getAll[driveItem](ctx, g, p, false, 1000)
	if err != nil {
		return err
	}
	return printItems(w, o, items)
}

// fetchFile lädt Metadaten + Inhalt.
func fetchMeta(ctx context.Context, g *graphClient, o opts) (driveItem, string, error) {
	var it driveItem
	r, err := refFromOpts(ctx, g, o)
	if err != nil {
		return it, "", err
	}
	p, err := r.itemPath(false)
	if err != nil {
		return it, "", err
	}
	if err := g.getJSON(ctx, p+"?"+itemSelect, true, &it); err != nil {
		return it, "", err
	}
	if it.Folder != nil {
		return it, "", fmt.Errorf("%q ist ein Ordner – dafür ls verwenden", it.Name)
	}
	return it, p, nil
}

func cmdGet(ctx context.Context, w io.Writer, o opts) error {
	g, err := newGraph(ctx)
	if err != nil {
		return err
	}
	it, p, err := fetchMeta(ctx, g, o)
	if err != nil {
		return err
	}
	resp, err := g.do(ctx, p+"/content", true)
	if err != nil {
		return err
	}
	defer resp.Body.Close()
	if o.out == "-" {
		_, err = io.Copy(w, resp.Body)
		return err
	}
	dst := o.out
	if dst == "" {
		dst = filepath.Base(filepath.Clean("/" + it.Name))
		if dst == "/" || dst == "." {
			dst = it.ID
		}
	}
	f, err := os.Create(dst)
	if err != nil {
		return err
	}
	n, err := io.Copy(f, resp.Body)
	if cerr := f.Close(); err == nil {
		err = cerr
	}
	if err != nil {
		return err
	}
	if o.json {
		return emitJSON(w, map[string]any{"file": dst, "bytes": n, "name": it.Name, "item": it.ID, "drive": it.ParentReference.DriveID})
	}
	fmt.Fprintf(w, "%s gespeichert (%d Bytes)\n", dst, n)
	return nil
}

func cmdText(ctx context.Context, w io.Writer, o opts) error {
	g, err := newGraph(ctx)
	if err != nil {
		return err
	}
	it, p, err := fetchMeta(ctx, g, o)
	if err != nil {
		return err
	}
	if strings.EqualFold(filepath.Ext(it.Name), ".pdf") {
		return errPDF
	}
	const maxDownload = 200 << 20
	if it.Size > maxDownload {
		return fmt.Errorf("Datei zu groß für text (%d Bytes) – mit get laden", it.Size)
	}
	resp, err := g.do(ctx, p+"/content", true)
	if err != nil {
		return err
	}
	defer resp.Body.Close()
	data, err := io.ReadAll(io.LimitReader(resp.Body, maxDownload))
	if err != nil {
		return err
	}
	txt, err := extractText(it.Name, data)
	if err != nil {
		return err
	}
	txt, cut := truncate(txt, o.maxChars)
	if o.json {
		return emitJSON(w, map[string]any{"name": it.Name, "item": it.ID, "drive": it.ParentReference.DriveID, "truncated": cut, "text": txt})
	}
	fmt.Fprint(w, txt)
	if !strings.HasSuffix(txt, "\n") {
		fmt.Fprintln(w)
	}
	if cut {
		fmt.Fprintf(os.Stderr, "[gekürzt auf %d Zeichen – mehr mit --max-chars]\n", o.maxChars)
	}
	return nil
}
