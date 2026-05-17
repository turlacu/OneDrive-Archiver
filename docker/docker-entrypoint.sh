#!/bin/sh
set -eu

PUID="${PUID:-99}"
PGID="${PGID:-100}"
UMASK="${UMASK:-022}"
APP_USER="appuser"
APP_GROUP="appgroup"

umask "$UMASK"

if ! awk -F: -v gid="$PGID" '$3 == gid { found=1 } END { exit found ? 0 : 1 }' /etc/group; then
  addgroup -g "$PGID" "$APP_GROUP"
else
  APP_GROUP="$(awk -F: -v gid="$PGID" '$3 == gid { print $1; exit }' /etc/group)"
fi

if ! awk -F: -v uid="$PUID" '$3 == uid { found=1 } END { exit found ? 0 : 1 }' /etc/passwd; then
  adduser -D -H -u "$PUID" -G "$APP_GROUP" "$APP_USER"
else
  APP_USER="$(awk -F: -v uid="$PUID" '$3 == uid { print $1; exit }' /etc/passwd)"
fi

if [ -n "${SERVER_DOWNLOAD_ROOT:-}" ]; then
  mkdir -p "$SERVER_DOWNLOAD_ROOT"
  chown "$PUID:$PGID" "$SERVER_DOWNLOAD_ROOT"
fi

if [ -n "${APP_DATA_DIR:-}" ]; then
  mkdir -p "$APP_DATA_DIR"
  chown "$PUID:$PGID" "$APP_DATA_DIR"
fi

exec su-exec "$PUID:$PGID" "$@"
