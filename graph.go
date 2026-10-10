package main

import (
	"context"
	"encoding/json"
	"fmt"
	"io"
	"net/http"
	"strings"
	"time"
)

const graphBase = "https://graph.microsoft.com/v1.0"

type graphClient struct {
	token string
	http  *http.Client
}

type graphError struct {
	Status  int
	Code    string
	Message string
	IsItem  bool
}

func (e *graphError) Error() string {
	switch e.Status {
	case 401:
		return "Nicht angemeldet oder Token ungültig – bitte erneut: sharepointcli login"
	case 403:
		return "Kein Zugriff mit deinem Microsoft-Login (das ist kein Fehler des Tools)"
	case 404:
		if e.IsItem {
			return "Nicht gefunden (404). Hinweis: Die ID ändert sich beim Verschieben über Bibliotheksgrenzen; mit --url neu auflösen"
		}
		return "Nicht gefunden (404)"
	case 429:
		return "Zu viele Anfragen (429) – bitte kurz warten und erneut versuchen"
	}
	msg := fmt.Sprintf("Graph-Fehler %d", e.Status)
	if e.Code != "" {
		msg += " " + e.Code
	}
	if e.Message != "" {
		msg += ": " + e.Message
	}
	return msg
}

func newGraph(ctx context.Context) (*graphClient, error) {
	tok, err := accessToken(ctx)
	if err != nil {
		return nil, err
	}
	return &graphClient{token: tok, http: &http.Client{Timeout: 10 * time.Minute}}, nil
}

func (g *graphClient) do(ctx context.Context, path string, isItem bool) (*http.Response, error) {
	u := path
	if !strings.HasPrefix(u, "https://") {
		u = graphBase + path
	}
	req, err := http.NewRequestWithContext(ctx, http.MethodGet, u, nil)
	if err != nil {
		return nil, err
	}
	req.Header.Set("Authorization", "Bearer "+g.token)
	req.Header.Set("Accept", "application/json")
	resp, err := g.http.Do(req)
	if err != nil {
		return nil, fmt.Errorf("Netzwerkfehler: %w", err)
	}
	if resp.StatusCode >= 300 {
		defer resp.Body.Close()
		ge := &graphError{Status: resp.StatusCode, IsItem: isItem}
		var body struct {
			Error struct {
				Code    string `json:"code"`
				Message string `json:"message"`
			} `json:"error"`
		}
		b, _ := io.ReadAll(io.LimitReader(resp.Body, 1<<16))
		if json.Unmarshal(b, &body) == nil {
			ge.Code, ge.Message = body.Error.Code, body.Error.Message
		}
		return nil, ge
	}
	return resp, nil
}

func (g *graphClient) getJSON(ctx context.Context, path string, isItem bool, out any) error {
	resp, err := g.do(ctx, path, isItem)
	if err != nil {
		return err
	}
	defer resp.Body.Close()
	return json.NewDecoder(resp.Body).Decode(out)
}

// getAll folgt @odata.nextLink bis max Einträge.
func getAll[T any](ctx context.Context, g *graphClient, path string, isItem bool, max int) ([]T, error) {
	var all []T
	for path != "" && len(all) < max {
		var page struct {
			Value    []T    `json:"value"`
			NextLink string `json:"@odata.nextLink"`
		}
		if err := g.getJSON(ctx, path, isItem, &page); err != nil {
			return nil, err
		}
		all = append(all, page.Value...)
		path = page.NextLink
	}
	if len(all) > max {
		all = all[:max]
	}
	return all, nil
}

type parentRef struct {
	DriveID string `json:"driveId"`
	SiteID  string `json:"siteId"`
	Path    string `json:"path,omitempty"`
}

type driveItem struct {
	ID                   string     `json:"id"`
	Name                 string     `json:"name"`
	WebURL               string     `json:"webUrl"`
	Size                 int64      `json:"size"`
	LastModifiedDateTime string     `json:"lastModifiedDateTime"`
	Folder               *struct {
		ChildCount int `json:"childCount"`
	} `json:"folder,omitempty"`
	File *struct {
		MimeType string `json:"mimeType"`
	} `json:"file,omitempty"`
	ParentReference parentRef `json:"parentReference"`
}

type site struct {
	ID          string `json:"id"`
	Name        string `json:"name"`
	DisplayName string `json:"displayName"`
	WebURL      string `json:"webUrl"`
}

type drive struct {
	ID        string `json:"id"`
	Name      string `json:"name"`
	DriveType string `json:"driveType"`
	WebURL    string `json:"webUrl"`
}
