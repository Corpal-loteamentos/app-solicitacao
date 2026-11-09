#!/usr/bin/env bash
set -euo pipefail

DESTINATION="${1:-/opt/corpal-teams-app}"
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

sudo install -d -o root -g www-data -m 0750 "$DESTINATION"

sudo rsync -a --delete \
  --exclude '.gitignore' \
  --exclude 'manifest.template.json' \
  --exclude 'build-teams-package.sh' \
  --exclude 'build-teams-package.ps1' \
  --exclude 'deploy-to-server.sh' \
  --exclude 'README.md' \
  --exclude 'Corpal-Solicitacoes-Teams-*.zip' \
  "$SCRIPT_DIR/" "$DESTINATION/"

sudo chown -R root:www-data "$DESTINATION"
sudo find "$DESTINATION" -type d -exec chmod 0750 {} \;
sudo find "$DESTINATION" -type f -exec chmod 0640 {} \;

echo "Frontend publicado em: $DESTINATION"
