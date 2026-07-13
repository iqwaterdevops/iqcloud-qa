#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
LOG_DIR="${ROOT_DIR}/logs"
mkdir -p "${LOG_DIR}"

TIMESTAMP="$(date '+%Y-%m-%d_%H-%M-%S')"
LOG_FILE="${LOG_DIR}/playwright-${TIMESTAMP}.log"

cd "${ROOT_DIR}"

echo "[$(date)] Starting Playwright test run" | tee -a "${LOG_FILE}"
if npm test 2>&1 | tee -a "${LOG_FILE}"; then
  echo "[$(date)] Playwright test run completed successfully" | tee -a "${LOG_FILE}"
else
  echo "[$(date)] Playwright test run failed" | tee -a "${LOG_FILE}"
  exit 1
fi
