#!/usr/bin/env bash
# Cross-Compile aller Release-Artefakte nach dist/.
set -euo pipefail
cd "$(dirname "$0")"
VERSION="${1:-${VERSION:-0.1.0}}"
export CGO_ENABLED=0
rm -rf dist && mkdir -p dist
go vet ./...
go test ./...
for target in windows/amd64 windows/arm64 darwin/amd64 darwin/arm64 linux/amd64 linux/arm64; do
  os="${target%/*}"; arch="${target#*/}"
  name="sharepointcli_${VERSION}_${os}_${arch}"
  stage="dist/${name}"
  mkdir -p "$stage"
  bin="sharepointcli"; [ "$os" = windows ] && bin="sharepointcli.exe"
  GOOS="$os" GOARCH="$arch" go build -trimpath -ldflags "-s -w -X main.version=${VERSION}" -o "$stage/$bin" .
  cp README.md LICENSE "$stage/"
  if [ "$os" = windows ]; then
    (cd "$stage" && zip -q -X "../${name}.zip" "$bin" README.md LICENSE)
  else
    tar -C "$stage" -czf "dist/${name}.tar.gz" "$bin" README.md LICENSE
  fi
  rm -rf "$stage"
  echo "gebaut: ${name}"
done
(cd dist && sha256sum *.zip *.tar.gz > SHA256SUMS)
cat dist/SHA256SUMS
