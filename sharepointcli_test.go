package main

import (
	"archive/zip"
	"bytes"
	"context"
	"encoding/base64"
	"errors"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

func mkZip(t *testing.T, files map[string]string) []byte {
	t.Helper()
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	for name, content := range files {
		f, err := zw.Create(name)
		if err != nil {
			t.Fatal(err)
		}
		if _, err := f.Write([]byte(content)); err != nil {
			t.Fatal(err)
		}
	}
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	return buf.Bytes()
}

func TestParseRef(t *testing.T) {
	cases := []struct {
		in   []string
		want itemRef
		err  bool
	}{
		{[]string{"site=S1 drive=D1 item=I1"}, itemRef{Site: "S1", Drive: "D1", Item: "I1"}, false},
		{[]string{"site=S1", "drive=D1", "item=I1"}, itemRef{Site: "S1", Drive: "D1", Item: "I1"}, false},
		{[]string{"site=contoso.sharepoint.com,abc,def"}, itemRef{Site: "contoso.sharepoint.com,abc,def"}, false},
		{[]string{"drive=b!xYz=="}, itemRef{Drive: "b!xYz=="}, false},
		{[]string{"drive=D", "path=/A/B.docx"}, itemRef{Drive: "D", Path: "/A/B.docx"}, false},
		{[]string{"drive=D", "/A/B"}, itemRef{Drive: "D", Path: "/A/B"}, false},
		{[]string{"https://contoso.sharepoint.com/x"}, itemRef{URL: "https://contoso.sharepoint.com/x"}, false},
		{[]string{"foo"}, itemRef{}, true},
		{[]string{"bar=1"}, itemRef{}, true},
		{[]string{"item="}, itemRef{}, true},
	}
	for _, c := range cases {
		got, err := parseRef(c.in)
		if c.err {
			if err == nil {
				t.Errorf("%v: Fehler erwartet", c.in)
			}
			continue
		}
		if err != nil || got != c.want {
			t.Errorf("%v: got %+v err %v, want %+v", c.in, got, err, c.want)
		}
	}
}

func TestRefRoundTrip(t *testing.T) {
	r := itemRef{Site: "s", Drive: "d", Item: "i"}
	got, err := parseRef([]string{r.String()})
	if err != nil || got != r {
		t.Fatalf("round trip: %+v %v", got, err)
	}
}

func TestParseArgsFlags(t *testing.T) {
	o, err := parseArgs([]string{"drive=D", "--item", "I", "--json", "--max-chars=5", "-o", "x.docx"})
	if err != nil {
		t.Fatal(err)
	}
	if !o.json || o.item != "I" || o.maxChars != 5 || o.out != "x.docx" || len(o.args) != 1 {
		t.Fatalf("%+v", o)
	}
	if _, err := parseArgs([]string{"--nope"}); err == nil {
		t.Fatal("unbekanntes Flag muss Fehler sein")
	}
	if o, _ := parseArgs(nil); o.maxChars != 20000 {
		t.Fatal("default max-chars")
	}
}

func TestItemPath(t *testing.T) {
	cases := []struct {
		r        itemRef
		root     bool
		want     string
		children string
	}{
		{itemRef{Drive: "D", Item: "I"}, false, "/drives/D/items/I", "/drives/D/items/I/children"},
		{itemRef{Site: "S", Item: "I"}, false, "/sites/S/drive/items/I", ""},
		{itemRef{Drive: "D"}, true, "/drives/D/root", "/drives/D/root/children"},
		{itemRef{Drive: "D", Path: "/Ordner A/ä.docx"}, false, "/drives/D/root:/Ordner%20A/%C3%A4.docx:", "/drives/D/root:/Ordner%20A/%C3%A4.docx:/children"},
	}
	for _, c := range cases {
		got, err := c.r.itemPath(c.root)
		if err != nil || got != c.want {
			t.Errorf("%+v: got %q %v want %q", c.r, got, err, c.want)
		}
		if c.children != "" {
			if ch, _ := c.r.childrenPath(); ch != c.children {
				t.Errorf("children %q want %q", ch, c.children)
			}
		}
	}
	if _, err := (itemRef{Drive: "D"}).itemPath(false); err == nil {
		t.Error("ohne item/path darf get nicht die Wurzel nehmen")
	}
	if _, err := (itemRef{Item: "I"}).itemPath(false); err == nil {
		t.Error("ohne drive/site Fehler erwartet")
	}
}

func TestEncodeShareURL(t *testing.T) {
	// Beispiel aus der Microsoft-Graph-Doku (shares-get).
	in := "https://onedrive.live.com/redir?resid=1231244193912!12&authKey=1201919!12921!1"
	want := "u!aHR0cHM6Ly9vbmVkcml2ZS5saXZlLmNvbS9yZWRpcj9yZXNpZD0xMjMxMjQ0MTkzOTEyITEyJmF1dGhLZXk9MTIwMTkxOSExMjkyMSEx"
	if got := encodeShareURL(in); got != want {
		t.Fatalf("got %s\nwant %s", got, want)
	}
	// Kein Padding, URL-sicheres Alphabet.
	got := encodeShareURL("https://contoso.sharepoint.com/sites/a?b=~~~>>>")
	if strings.ContainsAny(got[2:], "=+/") {
		t.Fatalf("nicht url-sicher: %s", got)
	}
	dec, err := base64.RawURLEncoding.DecodeString(got[2:])
	if err != nil || string(dec) != "https://contoso.sharepoint.com/sites/a?b=~~~>>>" {
		t.Fatalf("decode: %q %v", dec, err)
	}
}

func TestExtractDocx(t *testing.T) {
	doc := `<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
<w:p><w:r><w:t>Hallo</w:t></w:r><w:r><w:t xml:space="preserve"> Welt</w:t></w:r></w:p>
<w:p><w:r><w:t>Zweite</w:t><w:tab/><w:t>Zeile &amp; Ümlaut</w:t></w:r></w:p>
</w:body></w:document>`
	got, err := extractText("a.docx", mkZip(t, map[string]string{"word/document.xml": doc}))
	if err != nil {
		t.Fatal(err)
	}
	if got != "Hallo Welt\nZweite\tZeile & Ümlaut\n" {
		t.Fatalf("got %q", got)
	}
}

func TestExtractPptx(t *testing.T) {
	slide := func(s string) string {
		return `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>` + s + `</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`
	}
	z := mkZip(t, map[string]string{
		"ppt/slides/slide10.xml":            slide("Zehn"),
		"ppt/slides/slide2.xml":             slide("Zwei"),
		"ppt/slides/slide1.xml":             slide("Eins"),
		"ppt/slides/_rels/slide1.xml.rels": "<x/>",
	})
	got, err := extractText("p.PPTX", z)
	if err != nil {
		t.Fatal(err)
	}
	want := "--- Folie 1 ---\nEins\n--- Folie 2 ---\nZwei\n--- Folie 3 ---\nZehn\n"
	if got != want {
		t.Fatalf("got %q", got)
	}
}

func TestExtractXlsx(t *testing.T) {
	ss := `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3" uniqueCount="3">
<si><t>Name</t></si><si><t>Betrag</t></si><si><r><t>Mül</t></r><r><t>ler</t></r><rPh><t>X</t></rPh></si></sst>`
	sheet := `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>
<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row>
<row r="2"><c r="A2" t="s"><v>2</v></c><c r="C2"><v>42.5</v></c></row>
<row r="3"><c r="A3" t="inlineStr"><is><t>inline</t></is></c></row>
</sheetData></worksheet>`
	got, err := extractText("t.xlsx", mkZip(t, map[string]string{"xl/sharedStrings.xml": ss, "xl/worksheets/sheet1.xml": sheet}))
	if err != nil {
		t.Fatal(err)
	}
	want := "--- Blatt 1 ---\nName\tBetrag\nMüller\t\t42.5\ninline\n"
	if got != want {
		t.Fatalf("got %q want %q", got, want)
	}
}

func TestExtractPlainAndPDF(t *testing.T) {
	if got, _ := extractText("a.md", []byte("\xEF\xBB\xBF# Titel")); got != "# Titel" {
		t.Fatalf("got %q", got)
	}
	if _, err := extractText("a.pdf", nil); !errors.Is(err, errPDF) {
		t.Fatalf("pdf: %v", err)
	}
	if _, err := extractText("a.docx", []byte("kein zip")); err == nil {
		t.Fatal("kaputte docx muss Fehler sein")
	}
	if _, err := extractText("a.exe", nil); err == nil {
		t.Fatal("exe muss Fehler sein")
	}
}

func TestTruncate(t *testing.T) {
	if s, cut := truncate("äöüß", 2); s != "äö" || !cut {
		t.Fatalf("%q %v", s, cut)
	}
	if s, cut := truncate("abc", 0); s != "abc" || cut {
		t.Fatal("0 = unbegrenzt")
	}
	if s, cut := truncate("abc", 3); s != "abc" || cut {
		t.Fatal("exakt")
	}
}

func TestGraphErrorsGerman(t *testing.T) {
	if !strings.Contains((&graphError{Status: 403}).Error(), "kein Fehler des Tools") {
		t.Fatal("403")
	}
	if !strings.Contains((&graphError{Status: 404, IsItem: true}).Error(), "--url neu auflösen") {
		t.Fatal("404 item")
	}
}

func TestFileCacheAndConfig(t *testing.T) {
	d := t.TempDir()
	t.Setenv("SHAREPOINTCLI_CONFIG_DIR", d)
	t.Setenv("SHAREPOINTCLI_TENANT", "")
	t.Setenv("SHAREPOINTCLI_CLIENT_ID", "")
	fc := &fileCache{path: filepath.Join(d, "token_cache.json")}
	if err := fc.Export(context.Background(), marsh("abc"), struct{ PartitionKey string }{}); err != nil {
		t.Fatal(err)
	}
	st, err := os.Stat(fc.path)
	if err != nil {
		t.Fatal(err)
	}
	if st.Mode().Perm() != 0o600 {
		t.Fatalf("perm %v", st.Mode().Perm())
	}
	var u unmarsh
	if err := fc.Replace(context.Background(), &u, struct{ PartitionKey string }{}); err != nil || string(u) != "abc" {
		t.Fatalf("replace %q %v", u, err)
	}
	if _, err := resolveConfig("", ""); err == nil {
		t.Fatal("ohne config Fehler erwartet")
	}
	if err := saveConfig(config{Tenant: "t", ClientID: "c"}); err != nil {
		t.Fatal(err)
	}
	c, err := resolveConfig("", "c2")
	if err != nil || c.Tenant != "t" || c.ClientID != "c2" {
		t.Fatalf("%+v %v", c, err)
	}
}

type marsh string

func (m marsh) Marshal() ([]byte, error) { return []byte(m), nil }

type unmarsh []byte

func (u *unmarsh) Unmarshal(b []byte) error { *u = append([]byte(nil), b...); return nil }
