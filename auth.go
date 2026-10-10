package main

import (
	"context"
	"encoding/json"
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"strings"

	"github.com/AzureAD/microsoft-authentication-library-for-go/apps/cache"
	"github.com/AzureAD/microsoft-authentication-library-for-go/apps/public"
)

// Delegierte Graph-Rechte. offline_access/openid/profile ergänzt MSAL selbst.
var graphScopes = []string{
	"https://graph.microsoft.com/User.Read",
	"https://graph.microsoft.com/Sites.Read.All",
	"https://graph.microsoft.com/Files.Read.All",
}

type config struct {
	Tenant   string `json:"tenant"`
	ClientID string `json:"client_id"`
}

func configDir() (string, error) {
	if d := os.Getenv("SHAREPOINTCLI_CONFIG_DIR"); d != "" {
		return d, nil
	}
	base, err := os.UserConfigDir()
	if err != nil {
		return "", err
	}
	return filepath.Join(base, "sharepointcli"), nil
}

func ensureConfigDir() (string, error) {
	d, err := configDir()
	if err != nil {
		return "", err
	}
	if err := os.MkdirAll(d, 0o700); err != nil {
		return "", err
	}
	return d, nil
}

func loadConfig() (config, error) {
	var c config
	d, err := configDir()
	if err != nil {
		return c, err
	}
	b, err := os.ReadFile(filepath.Join(d, "config.json"))
	if err != nil && !errors.Is(err, os.ErrNotExist) {
		return c, err
	}
	if len(b) > 0 {
		if err := json.Unmarshal(b, &c); err != nil {
			return c, fmt.Errorf("config.json beschädigt: %w", err)
		}
	}
	return c, nil
}

func saveConfig(c config) error {
	d, err := ensureConfigDir()
	if err != nil {
		return err
	}
	b, _ := json.MarshalIndent(c, "", "  ")
	return writeFile0600(filepath.Join(d, "config.json"), b)
}

func writeFile0600(path string, data []byte) error {
	tmp := path + ".tmp"
	if err := os.WriteFile(tmp, data, 0o600); err != nil {
		return err
	}
	_ = os.Chmod(tmp, 0o600)
	return os.Rename(tmp, path)
}

// fileCache implementiert cache.ExportReplace: der komplette MSAL-Cache liegt
// in einer Datei mit Rechten 0600.
type fileCache struct{ path string }

func (f *fileCache) Replace(ctx context.Context, u cache.Unmarshaler, _ cache.ReplaceHints) error {
	if err := ctx.Err(); err != nil {
		return err
	}
	b, err := os.ReadFile(f.path)
	if errors.Is(err, os.ErrNotExist) {
		return nil
	}
	if err != nil {
		return err
	}
	return u.Unmarshal(b)
}

func (f *fileCache) Export(ctx context.Context, m cache.Marshaler, _ cache.ExportHints) error {
	if err := ctx.Err(); err != nil {
		return err
	}
	b, err := m.Marshal()
	if err != nil {
		return err
	}
	if err := os.MkdirAll(filepath.Dir(f.path), 0o700); err != nil {
		return err
	}
	return writeFile0600(f.path, b)
}

func cachePath() (string, error) {
	d, err := configDir()
	if err != nil {
		return "", err
	}
	return filepath.Join(d, "token_cache.json"), nil
}

// resolveConfig: Flags > Umgebungsvariablen > gespeicherte config.json.
func resolveConfig(tenant, clientID string) (config, error) {
	c, err := loadConfig()
	if err != nil {
		return c, err
	}
	if v := os.Getenv("SHAREPOINTCLI_TENANT"); v != "" {
		c.Tenant = v
	}
	if v := os.Getenv("SHAREPOINTCLI_CLIENT_ID"); v != "" {
		c.ClientID = v
	}
	if tenant != "" {
		c.Tenant = tenant
	}
	if clientID != "" {
		c.ClientID = clientID
	}
	if c.Tenant == "" || c.ClientID == "" {
		return c, errors.New("Tenant und Client-ID fehlen. Einmalig: sharepointcli login --tenant <tenant-id|domain> --client-id <app-id>")
	}
	return c, nil
}

func newClient(c config) (public.Client, error) {
	cp, err := cachePath()
	if err != nil {
		return public.Client{}, err
	}
	authority := "https://login.microsoftonline.com/" + strings.TrimSpace(c.Tenant)
	return public.New(c.ClientID, public.WithAuthority(authority), public.WithCache(&fileCache{path: cp}))
}

func login(ctx context.Context, c config, deviceCode bool) (public.AuthResult, error) {
	client, err := newClient(c)
	if err != nil {
		return public.AuthResult{}, err
	}
	if deviceCode {
		dc, err := client.AcquireTokenByDeviceCode(ctx, graphScopes)
		if err != nil {
			return public.AuthResult{}, err
		}
		fmt.Fprintln(os.Stderr, dc.Result.Message)
		return dc.AuthenticationResult(ctx)
	}
	fmt.Fprintln(os.Stderr, "Öffne Browser für den Microsoft-Login … (ohne Browser: --device-code)")
	return client.AcquireTokenInteractive(ctx, graphScopes, public.WithRedirectURI("http://localhost"))
}

// accessToken holt ein Token still aus dem Cache (inkl. Refresh).
func accessToken(ctx context.Context) (string, error) {
	c, err := resolveConfig("", "")
	if err != nil {
		return "", err
	}
	client, err := newClient(c)
	if err != nil {
		return "", err
	}
	accts, err := client.Accounts(ctx)
	if err != nil {
		return "", err
	}
	if len(accts) == 0 {
		return "", errors.New("nicht eingeloggt – bitte zuerst: sharepointcli login")
	}
	res, err := client.AcquireTokenSilent(ctx, graphScopes, public.WithSilentAccount(accts[0]))
	if err != nil {
		return "", fmt.Errorf("Anmeldung abgelaufen oder widerrufen – bitte erneut: sharepointcli login (%v)", err)
	}
	return res.AccessToken, nil
}

func logout(ctx context.Context) error {
	if c, err := resolveConfig("", ""); err == nil {
		if client, err := newClient(c); err == nil {
			if accts, err := client.Accounts(ctx); err == nil {
				for _, a := range accts {
					_ = client.RemoveAccount(ctx, a)
				}
			}
		}
	}
	cp, err := cachePath()
	if err != nil {
		return err
	}
	if err := os.Remove(cp); err != nil && !errors.Is(err, os.ErrNotExist) {
		return err
	}
	return nil
}
