#!/usr/bin/env bash
# AEDI-S PostgreSQL Setup
# ------------------------
# One-shot installer + provisioner for the hub-side Postgres database.
#
# Run once when you want to promote the JSONL audit log (data/events.jsonl)
# to a proper relational store. The JSONL path keeps working regardless —
# Postgres is an additive, optional backend.
#
#   sudo ./scripts/setup-postgres.sh
#
# What it does (idempotent):
#   1. apt install postgresql postgresql-contrib (if missing)
#   2. Ensure the service is enabled + running
#   3. Create role `aedis` with password (AEDIS_DB_PASSWORD env or prompt)
#   4. Create database `aedis_hub` owned by aedis
#   5. Create core tables: events, role_changes, nodes_snapshot
#   6. Write connection URL to .env.local as AEDIS_DATABASE_URL
#
# Safe to re-run: all CREATE statements use IF NOT EXISTS / DO $$.
set -euo pipefail

if [[ $EUID -ne 0 ]]; then
  echo "This script needs sudo (apt install + service control)."
  echo "Re-run: sudo $0"
  exit 1
fi

REAL_USER="${SUDO_USER:-${USER}}"
WORKSPACE_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
ENV_FILE="${WORKSPACE_DIR}/.env.local"

DB_NAME="${AEDIS_DB_NAME:-aedis_hub}"
DB_USER="${AEDIS_DB_USER:-aedis}"
DB_PASS="${AEDIS_DB_PASSWORD:-}"

echo "── AEDI-S PostgreSQL Setup ──────────────────────────────"
echo "  workspace : ${WORKSPACE_DIR}"
echo "  db name   : ${DB_NAME}"
echo "  db user   : ${DB_USER}"
echo ""

# 1. Install if missing
if ! command -v psql >/dev/null 2>&1; then
  echo "▶ Installing postgresql + postgresql-contrib..."
  apt-get update -qq
  apt-get install -y postgresql postgresql-contrib
else
  echo "✓ psql already installed ($(psql --version))"
fi

# 2. Enable + start service
systemctl enable postgresql >/dev/null 2>&1 || true
systemctl start postgresql
echo "✓ postgresql service: $(systemctl is-active postgresql)"

# 3. Password — generate if not provided
if [[ -z "${DB_PASS}" ]]; then
  DB_PASS="$(tr -dc 'A-Za-z0-9' </dev/urandom | head -c 24 || echo "aedis-$(date +%s)")"
  echo "▶ Generated DB password (saved to .env.local)"
fi

# 4. Role + database
sudo -u postgres psql -v ON_ERROR_STOP=1 <<SQL
DO \$\$
BEGIN
  IF NOT EXISTS (SELECT 1 FROM pg_roles WHERE rolname = '${DB_USER}') THEN
    CREATE ROLE ${DB_USER} LOGIN PASSWORD '${DB_PASS}';
  ELSE
    ALTER ROLE ${DB_USER} WITH PASSWORD '${DB_PASS}';
  END IF;
END
\$\$;
SQL

# CREATE DATABASE cannot run inside a transaction — separate call.
if ! sudo -u postgres psql -tAc "SELECT 1 FROM pg_database WHERE datname='${DB_NAME}'" | grep -q 1; then
  sudo -u postgres createdb -O "${DB_USER}" "${DB_NAME}"
  echo "✓ Created database ${DB_NAME}"
else
  echo "✓ Database ${DB_NAME} already exists"
fi

# 5. Schema
sudo -u postgres psql -d "${DB_NAME}" -v ON_ERROR_STOP=1 <<SQL
CREATE TABLE IF NOT EXISTS events (
  id          BIGSERIAL PRIMARY KEY,
  ts          TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  kind        TEXT NOT NULL,
  payload     JSONB NOT NULL
);
CREATE INDEX IF NOT EXISTS events_kind_idx ON events (kind);
CREATE INDEX IF NOT EXISTS events_ts_idx   ON events (ts DESC);

CREATE TABLE IF NOT EXISTS role_changes (
  id        BIGSERIAL PRIMARY KEY,
  ts        TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  node_id   SMALLINT NOT NULL,
  prev_role TEXT,
  new_role  TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS nodes_snapshot (
  ts        TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  node_id   SMALLINT NOT NULL,
  role      TEXT,
  rssi_dbm  REAL,
  active    BOOLEAN,
  PRIMARY KEY (ts, node_id)
);

GRANT ALL PRIVILEGES ON ALL TABLES IN SCHEMA public TO ${DB_USER};
GRANT ALL PRIVILEGES ON ALL SEQUENCES IN SCHEMA public TO ${DB_USER};
SQL
echo "✓ Schema installed (events, role_changes, nodes_snapshot)"

# 6. .env.local
DB_URL="postgres://${DB_USER}:${DB_PASS}@127.0.0.1:5432/${DB_NAME}"
touch "${ENV_FILE}"
chown "${REAL_USER}:${REAL_USER}" "${ENV_FILE}"
if grep -q "^AEDIS_DATABASE_URL=" "${ENV_FILE}"; then
  sed -i "s|^AEDIS_DATABASE_URL=.*|AEDIS_DATABASE_URL=${DB_URL}|" "${ENV_FILE}"
else
  printf "\n# AEDI-S Postgres (installed %s)\nAEDIS_DATABASE_URL=%s\n" \
    "$(date -Iseconds)" "${DB_URL}" >> "${ENV_FILE}"
fi
chmod 600 "${ENV_FILE}"

echo ""
echo "── Done ─────────────────────────────────────────────────"
echo "  Connection URL written to ${ENV_FILE}"
echo "  Verify:  sudo -u postgres psql -d ${DB_NAME} -c '\\dt'"
echo "  Connect: psql '${DB_URL}'"
echo ""
echo "  The hub's Monitor → Storage tab will now show 'PostgreSQL: online'."
