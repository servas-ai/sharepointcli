package main

import (
	"encoding/base64"
	"fmt"
	"net/url"
	"strings"
)

// itemRef adressiert ein DriveItem: per IDs, per Pfad oder per SharePoint-URL.
type itemRef struct {
	Site  string `json:"site,omitempty"`
	Drive string `json:"drive,omitempty"`
	Item  string `json:"item,omitempty"`
	Path  string `json:"path,omitempty"`
	URL   string `json:"url,omitempty"`
}

func (r itemRef) String() string {
	var p []string
	if r.Site != "" {
		p = append(p, "site="+r.Site)
	}
	if r.Drive != "" {
		p = append(p, "drive="+r.Drive)
	}
	if r.Item != "" {
		p = append(p, "item="+r.Item)
	}
	if r.Path != "" {
		p = append(p, "path="+r.Path)
	}
	if r.URL != "" {
		p = append(p, "url="+r.URL)
	}
	return strings.Join(p, " ")
}

// parseRef liest "site=… drive=… item=… path=…" aus einem oder mehreren
// Argumenten. Ein nacktes https://-Argument gilt als URL, ein Argument mit
// führendem "/" als Pfad.
func parseRef(args []string) (itemRef, error) {
	var r itemRef
	for _, a := range args {
		for _, tok := range strings.Fields(a) {
			if strings.HasPrefix(tok, "https://") || strings.HasPrefix(tok, "http://") {
				r.URL = tok
				continue
			}
			if strings.HasPrefix(tok, "/") {
				r.Path = tok
				continue
			}
			k, v, ok := strings.Cut(tok, "=")
			if !ok || v == "" {
				return r, fmt.Errorf("ungültige Referenz %q – erwartet site=<id> drive=<id> item=<id> (oder --url)", tok)
			}
			switch strings.ToLower(k) {
			case "site":
				r.Site = v
			case "drive":
				r.Drive = v
			case "item":
				r.Item = v
			case "path":
				r.Path = v
			case "url":
				r.URL = v
			default:
				return r, fmt.Errorf("unbekannter Referenzschlüssel %q (erlaubt: site, drive, item, path, url)", k)
			}
		}
	}
	return r, nil
}

func (r itemRef) merge(o itemRef) itemRef {
	if o.Site != "" {
		r.Site = o.Site
	}
	if o.Drive != "" {
		r.Drive = o.Drive
	}
	if o.Item != "" {
		r.Item = o.Item
	}
	if o.Path != "" {
		r.Path = o.Path
	}
	if o.URL != "" {
		r.URL = o.URL
	}
	return r
}

// encodeShareURL kodiert eine Freigabe-/SharePoint-URL für /shares/{id}:
// "u!" + base64url ohne Padding.
func encodeShareURL(u string) string {
	return "u!" + base64.RawURLEncoding.EncodeToString([]byte(strings.TrimSpace(u)))
}

// escapePath kodiert jedes Pfadsegment einzeln.
func escapePath(p string) string {
	segs := strings.Split(strings.Trim(p, "/"), "/")
	for i, s := range segs {
		segs[i] = url.PathEscape(s)
	}
	return strings.Join(segs, "/")
}

// itemPath liefert den Graph-Pfad des Items (ohne Suffix). URL muss vorher
// aufgelöst sein. allowRoot: ohne item/path wird das Wurzelverzeichnis genutzt.
func (r itemRef) itemPath(allowRoot bool) (string, error) {
	var base string
	switch {
	case r.Drive != "":
		base = "/drives/" + url.PathEscape(r.Drive)
	case r.Site != "":
		base = "/sites/" + url.PathEscape(r.Site) + "/drive"
	default:
		return "", fmt.Errorf("Referenz unvollständig: drive=<id> oder site=<id> nötig (oder --url)")
	}
	switch {
	case r.Item != "":
		return base + "/items/" + url.PathEscape(r.Item), nil
	case r.Path != "" && strings.Trim(r.Path, "/") != "":
		return base + "/root:/" + escapePath(r.Path) + ":", nil
	case allowRoot:
		return base + "/root", nil
	}
	return "", fmt.Errorf("Referenz unvollständig: item=<id> oder path=/… nötig")
}

// childrenPath: bei Pfad-Adressierung braucht Graph "root:/p:/children".
func (r itemRef) childrenPath() (string, error) {
	p, err := r.itemPath(true)
	if err != nil {
		return "", err
	}
	return p + "/children", nil
}
