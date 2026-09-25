#!/usr/bin/env bash
set -euo pipefail

: "${AZURE_CLIENT_ID:?Defina AZURE_CLIENT_ID}"
: "${AZURE_APP_ID_URI:?Defina AZURE_APP_ID_URI}"

if ! [[ "$AZURE_CLIENT_ID" =~ ^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$ ]]; then
  echo "AZURE_CLIENT_ID deve ser um GUID valido." >&2
  exit 1
fi

if ! [[ "$AZURE_APP_ID_URI" =~ ^api:// ]]; then
  echo "AZURE_APP_ID_URI deve comecar com api://" >&2
  exit 1
fi

script_dir="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
workdir="$(mktemp -d)"
trap 'rm -rf "$workdir"' EXIT

sed \
  -e "s|__AZURE_CLIENT_ID__|${AZURE_CLIENT_ID}|g" \
  -e "s|__AZURE_APP_ID_URI__|${AZURE_APP_ID_URI}|g" \
  "$script_dir/manifest.template.json" > "$workdir/manifest.json"

cp "$script_dir/color.png" "$script_dir/outline.png" "$workdir/"

output="$script_dir/Corpal-Solicitacoes-Teams-1.0.29.zip"
rm -f "$output"
(
  cd "$workdir"
  zip -q "$output" manifest.json color.png outline.png
)

echo "Pacote gerado: $output"
