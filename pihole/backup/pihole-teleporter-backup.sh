#!/usr/bin/env bash
set -euo pipefail

TARGET_DIR="/mnt/nas/pihole-backups"
RETENTION_DAYS=30

# Ensure NAS is mounted
if ! mountpoint -q ${TARGET_DIR}; then
  echo "ERROR: NAS mount not present: /mnt/abs-nas01/pihole" >&2
  exit 1
fi

mkdir -p "${TARGET_DIR}"

# Teleporter export (Pi-hole v6)
# Prints the created filename on stdout
OUT_FILE="$(pihole-FTL --teleporter | tail -n 1 | tr -d '\r')"
if [[ -z "${OUT_FILE}" ]]; then
  echo "ERROR: Teleporter did not return a filename" >&2
  exit 1
fi

if [[ ! -f "${OUT_FILE}" ]]; then
  echo "ERROR: Teleporter file not found: ${OUT_FILE}" >&2
  exit 1
fi

# Copy to NAS with timestamp prefix
TS="$(date +%F_%H-%M-%S)"
DEST="${TARGET_DIR}/${OUT_FILE}"

cp -f "${OUT_FILE}" "${DEST}"

# Basic integrity check (works for .tar.gz / .tgz)
gzip -t "${DEST}" >/dev/null 2>&1 || true

# Cleanup local export file (optional)
rm -f "${OUT_FILE}" || true

# Retention on NAS
find "${TARGET_DIR}" -maxdepth 1 -type f -name 'pi-hole_pihole-teleporter_*' -mtime +"${RETENTION_DAYS}" -print -delete

echo "OK: ${DEST}"
