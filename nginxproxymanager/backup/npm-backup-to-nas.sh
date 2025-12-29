#!/usr/bin/env bash
set -euo pipefail

# Target on NAS (bind-mounted into the container)
NAS_ROOT="/mnt/nas/npm-backups"
RETENTION_DAYS=60

# Adjust these paths to your NPM locations
NPM_DATA_DIR="/data"
NPM_LE_DIR="/etc/letsencrypt"

# Optional: set explicitly if you know it (leave empty for auto-detect)
SERVICE_NAME=""

LOG_FILE="/var/log/npm-backup-to-nas.log"

log() {
  printf "%s %s\n" "$(date -Is)" "$*" | tee -a "${LOG_FILE}"
}

fail() {
  log "ERROR: $*"
  exit 1
}

has_systemd() {
  command -v systemctl >/dev/null 2>&1 && [[ -d /run/systemd/system ]]
}

# Validate NAS mount
mountpoint -q "${NAS_ROOT}" || fail "NAS path not mounted: ${NAS_ROOT}"
mkdir -p "${NAS_ROOT}/archives" "${NAS_ROOT}/logs"

# Validate source paths
[[ -d "${NPM_DATA_DIR}" ]] || fail "NPM_DATA_DIR not found: ${NPM_DATA_DIR}"
[[ -f "${NPM_DATA_DIR}/database.sqlite" ]] || fail "database.sqlite not found under: ${NPM_DATA_DIR}"
[[ -d "${NPM_LE_DIR}" ]] || fail "NPM_LE_DIR not found: ${NPM_LE_DIR}"

# Detect service name if not set
if [[ -z "${SERVICE_NAME}" ]] && has_systemd; then
  SERVICE_NAME="$(systemctl list-units --type=service --all --no-pager 2>/dev/null \
    | awk '{print $1}' \
    | grep -Ei 'nginx.*proxy.*manager|nginxproxymanager|npm' \
    | head -n 1 || true)"
fi

TS="$(date +%F_%H-%M-%S)"
HOST="$(hostname -s)"
ARCHIVE="${NAS_ROOT}/archives/npm_${HOST}_${TS}.tgz"
NAS_LOG="${NAS_ROOT}/logs/npm_${HOST}.log"

SERVICE_STOPPED="no"

stop_service() {
  if has_systemd && [[ -n "${SERVICE_NAME}" ]]; then
    if systemctl list-units --type=service --all --no-pager 2>/dev/null | awk '{print $1}' | grep -qx "${SERVICE_NAME}"; then
      log "Stopping service: ${SERVICE_NAME}"
      systemctl stop "${SERVICE_NAME}"
      return 0
    fi
  fi
  return 1
}

start_service() {
  if has_systemd && [[ -n "${SERVICE_NAME}" ]]; then
    log "Starting service: ${SERVICE_NAME}"
    systemctl start "${SERVICE_NAME}"
  fi
}

if stop_service; then
  SERVICE_STOPPED="yes"
else
  log "No systemd service detected or matching service not found. Proceeding without stop."
fi

log "Creating archive: ${ARCHIVE}"
tar -czf "${ARCHIVE}" \
  --numeric-owner \
  -C / \
  "${NPM_DATA_DIR#/}" \
  "${NPM_LE_DIR#/}"

if [[ "${SERVICE_STOPPED}" == "yes" ]]; then
  start_service
fi

log "Validating archive gzip integrity"
gzip -t "${ARCHIVE}" >/dev/null

# Write last log lines to NAS log
{
  printf "%s\n" "$(date -Is) INFO: Backup written to ${ARCHIVE}";
  tail -n 200 "${LOG_FILE}" || true;
  printf "%s\n" "$(date -Is) INFO: Retention cleanup (>${RETENTION_DAYS} days)";
} >> "${NAS_LOG}" 2>&1 || true

# Retention
find "${NAS_ROOT}/archives" -type f -name "npm_${HOST}_*.tgz" -mtime +"${RETENTION_DAYS}" -print -delete >> "${NAS_LOG}" 2>&1 || true

log "OK: ${ARCHIVE}"
