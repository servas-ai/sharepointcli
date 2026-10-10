package main

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"errors"
	"fmt"
	"io"
	"path"
	"regexp"
	"sort"
	"strconv"
	"strings"
	"unicode/utf8"
)

var errPDF = errors.New("PDF-Text nicht unterstützt in 0.1, Datei mit get laden")

// extractText wählt den Extraktor anhand der Dateiendung.
func extractText(name string, data []byte) (string, error) {
	ext := strings.ToLower(path.Ext(name))
	switch ext {
	case ".txt", ".md", ".csv", ".tsv", ".json", ".xml", ".html", ".htm", ".log", ".yaml", ".yml":
		if !utf8.Valid(data) {
			return strings.ToValidUTF8(string(data), "�"), nil
		}
		return string(bytes.TrimPrefix(data, []byte("\xEF\xBB\xBF"))), nil
	case ".docx", ".docm":
		return extractDocx(data)
	case ".pptx", ".pptm":
		return extractPptx(data)
	case ".xlsx", ".xlsm":
		return extractXlsx(data)
	case ".pdf":
		return "", errPDF
	}
	return "", fmt.Errorf("Dateityp %q wird von text nicht unterstützt (docx, pptx, xlsx, txt, md, csv); Datei mit get laden", ext)
}

func openZip(data []byte) (map[string]*zip.File, error) {
	zr, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		return nil, fmt.Errorf("kein gültiges Office-Dokument (ZIP): %w", err)
	}
	m := make(map[string]*zip.File, len(zr.File))
	for _, f := range zr.File {
		m[f.Name] = f
	}
	return m, nil
}

func readZipFile(f *zip.File) ([]byte, error) {
	rc, err := f.Open()
	if err != nil {
		return nil, err
	}
	defer rc.Close()
	return io.ReadAll(io.LimitReader(rc, 256<<20))
}

// xmlText sammelt CharData der Elemente textLocal; paraLocal erzeugt Zeilenumbrüche.
func xmlText(data []byte, textLocal, paraLocal string) (string, error) {
	dec := xml.NewDecoder(bytes.NewReader(data))
	var sb strings.Builder
	inText := 0
	for {
		tok, err := dec.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return "", err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case textLocal:
				inText++
			case "tab":
				sb.WriteByte('\t')
			case "br", "cr":
				sb.WriteByte('\n')
			}
		case xml.EndElement:
			if t.Name.Local == textLocal && inText > 0 {
				inText--
			} else if t.Name.Local == paraLocal {
				sb.WriteByte('\n')
			}
		case xml.CharData:
			if inText > 0 {
				sb.Write(t)
			}
		}
	}
	return sb.String(), nil
}

func extractDocx(data []byte) (string, error) {
	files, err := openZip(data)
	if err != nil {
		return "", err
	}
	f, ok := files["word/document.xml"]
	if !ok {
		return "", errors.New("word/document.xml fehlt – keine gültige .docx")
	}
	b, err := readZipFile(f)
	if err != nil {
		return "", err
	}
	return xmlText(b, "t", "p")
}

var numRe = regexp.MustCompile(`(\d+)\.xml$`)

func sortedParts(files map[string]*zip.File, prefix string) []string {
	var names []string
	for n := range files {
		if strings.HasPrefix(n, prefix) && strings.HasSuffix(n, ".xml") && !strings.Contains(strings.TrimPrefix(n, prefix), "/") {
			names = append(names, n)
		}
	}
	num := func(s string) int {
		m := numRe.FindStringSubmatch(s)
		if m == nil {
			return 0
		}
		i, _ := strconv.Atoi(m[1])
		return i
	}
	sort.Slice(names, func(i, j int) bool { return num(names[i]) < num(names[j]) })
	return names
}

func extractPptx(data []byte) (string, error) {
	files, err := openZip(data)
	if err != nil {
		return "", err
	}
	slides := sortedParts(files, "ppt/slides/slide")
	if len(slides) == 0 {
		return "", errors.New("keine Folien gefunden – keine gültige .pptx")
	}
	var sb strings.Builder
	for i, s := range slides {
		b, err := readZipFile(files[s])
		if err != nil {
			return "", err
		}
		t, err := xmlText(b, "t", "p")
		if err != nil {
			return "", err
		}
		fmt.Fprintf(&sb, "--- Folie %d ---\n%s\n", i+1, strings.TrimRight(t, "\n"))
	}
	return sb.String(), nil
}

func extractXlsx(data []byte) (string, error) {
	files, err := openZip(data)
	if err != nil {
		return "", err
	}
	var shared []string
	if f, ok := files["xl/sharedStrings.xml"]; ok {
		b, err := readZipFile(f)
		if err != nil {
			return "", err
		}
		shared, err = parseSharedStrings(b)
		if err != nil {
			return "", err
		}
	}
	sheets := sortedParts(files, "xl/worksheets/sheet")
	if len(sheets) == 0 {
		return "", errors.New("keine Tabellenblätter gefunden – keine gültige .xlsx")
	}
	var sb strings.Builder
	for i, s := range sheets {
		b, err := readZipFile(files[s])
		if err != nil {
			return "", err
		}
		rows, err := parseSheet(b, shared)
		if err != nil {
			return "", err
		}
		fmt.Fprintf(&sb, "--- Blatt %d ---\n", i+1)
		for _, r := range rows {
			sb.WriteString(strings.Join(r, "\t"))
			sb.WriteByte('\n')
		}
	}
	return sb.String(), nil
}

func parseSharedStrings(b []byte) ([]string, error) {
	dec := xml.NewDecoder(bytes.NewReader(b))
	var out []string
	var cur strings.Builder
	inSI, inT, inRPh := false, false, false
	for {
		tok, err := dec.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return nil, err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "si":
				inSI = true
				cur.Reset()
			case "t":
				inT = true
			case "rPh":
				inRPh = true
			}
		case xml.EndElement:
			switch t.Name.Local {
			case "si":
				out = append(out, cur.String())
				inSI = false
			case "t":
				inT = false
			case "rPh":
				inRPh = false
			}
		case xml.CharData:
			if inSI && inT && !inRPh {
				cur.Write(t)
			}
		}
	}
	return out, nil
}

// colIndex: "C7" -> 2
func colIndex(ref string) int {
	n := 0
	for _, c := range ref {
		if c >= 'A' && c <= 'Z' {
			n = n*26 + int(c-'A'+1)
		} else {
			break
		}
	}
	return n - 1
}

func parseSheet(b []byte, shared []string) ([][]string, error) {
	dec := xml.NewDecoder(bytes.NewReader(b))
	var rows [][]string
	var row []string
	var cellType, cellRef string
	var val strings.Builder
	inV, inIS, inT := false, false, false
	for {
		tok, err := dec.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return nil, err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "row":
				row = nil
			case "c":
				cellType, cellRef = "", ""
				val.Reset()
				for _, a := range t.Attr {
					switch a.Name.Local {
					case "t":
						cellType = a.Value
					case "r":
						cellRef = a.Value
					}
				}
			case "v":
				inV = true
			case "is":
				inIS = true
			case "t":
				inT = true
			}
		case xml.EndElement:
			switch t.Name.Local {
			case "v":
				inV = false
			case "is":
				inIS = false
			case "t":
				inT = false
			case "c":
				s := val.String()
				if cellType == "s" {
					if i, err := strconv.Atoi(strings.TrimSpace(s)); err == nil && i >= 0 && i < len(shared) {
						s = shared[i]
					}
				}
				idx := len(row)
				if cellRef != "" {
					if ci := colIndex(cellRef); ci >= idx {
						idx = ci
					}
				}
				for len(row) < idx {
					row = append(row, "")
				}
				row = append(row, s)
			case "row":
				rows = append(rows, row)
			}
		case xml.CharData:
			if inV || (inIS && inT) {
				val.Write(t)
			}
		}
	}
	return rows, nil
}

// truncate begrenzt auf maxChars Zeichen (Runen).
func truncate(s string, maxChars int) (string, bool) {
	if maxChars <= 0 {
		return s, false
	}
	n := 0
	for i := range s {
		if n == maxChars {
			return s[:i], true
		}
		n++
	}
	return s, false
}
