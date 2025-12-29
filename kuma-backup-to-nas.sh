#!/usr/bin/env bash
set -euo pipefail

# ------------------------------------------------------------------------------
# Uptime Kuma Backup to NAS (unprivileged LXC with bind-mounted CIFS share)
# ------------------------------------------------------------------------------
# - Uptime Kuma runs directly in this container (not Docker)
# - The NAS share is bind-mounted under /mnt/nas/kuma-backups
# - Backups include app data and config
# ------------------------------------------------------------------------------

NAS_ROOT="/mnt/nas/kuma-backups"
RETENTION_DAYS=30
KUMA_DATA_DIR="/opt/uptime-kuma/data/"
LOG_FILE="/var/log/kuma-backup-to-nas.log"

log() {
  printf "%s %s\n" "$(date -Is)" "$*" | tee -a "${LOG_FILE}"
}

fail() {
  log "ERROR: $*"
  exit 1
}

# Ensure NAS is mounted
mountpoint -q "${NAS_ROOT}" || fail "NAS path not mounted: ${NAS_ROOT}"

# Ensure source data exists
[[ -d "${KUMA_DATA_DIR}" ]] || fail "KUMA_DATA_DIR not found: ${KUMA_DATA_DIR}"

# Create subfolders if not existing
mkdir -p "${NAS_ROOT}/archives" "${NAS_ROOT}/logs"

TS="$(date +%F_%H-%M-%S)"
HOST="$(hostname -s)"
ARCHIVE="${NAS_ROOT}/archives/kuma_${HOST}_${TS}.tgz"
NAS_LOG="${NAS_ROOT}/logs/kuma_${HOST}.log"

log "Creating archive ${ARCHIVE}"
tar -czf "${ARCHIVE}" \
  --numeric-owner \
  -C / \
  "${KUMA_DATA_DIR#/}"

log "Validating archive gzip integrity"
gzip -t "${ARCHIVE}" >/dev/null || fail "Archive integrity check failed"

# Retention cleanup
{
  printf "%s\n" "$(date -Is) INFO: Retention cleanup (>${RETENTION_DAYS} days)"
  find "${NAS_ROOT}/archives" -type f -name "kuma_${HOST}_*.tgz" -mtime +"${RETENTION_DAYS}" -print -delete
} >> "${NAS_LOG}" 2>&1 || true

log "OK: ${ARCHIVE}"
