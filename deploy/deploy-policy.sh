#!/usr/bin/env bash
# Deploys the admin-owned versioned policy.json (Outlook and Teams recipients,
# Calendar attendees + AI disclosure).
#
# Builds policy.json from parameters, atomically writes it to the admin-owned system location
# (Linux: /etc/microsoft-mcp/policy.json, macOS: /Library/Application Support/...)
# as root:root mode 644, so a user-level LLM with file access can neither rewrite
# nor delete it. The MCP server reads this file as its ONLY policy source
# (Messaging__* env vars are ignored by design) and refuses to start with a
# user-writable restrictive policy.
#
# Usage (as root, e.g. via sudo):
#   sudo ./deploy-policy.sh --domains firma.de,tochter.firma.de \
#     --disclosure-text "Hinweis: Dieser Entwurf wurde von einer KI erstellt und muss vor dem Versand geprüft werden."
#   sudo ./deploy-policy.sh --addresses partner@example.com \
#     --disclosure-text "Hinweis: Dieser Entwurf wurde von einer KI erstellt und muss vor dem Versand geprüft werden."
#   sudo ./deploy-policy.sh --domains firma.de --calendar-no-restrict \
#     --disclosure-text "Hinweis: Dieser Entwurf wurde von einer KI erstellt und muss vor dem Versand geprüft werden."
#   sudo ./deploy-policy.sh --domains firma.de --disclosure-text "..." --path /custom/policy.json
set -euo pipefail

DOMAINS=""
ADDRESSES=""
DISCLOSURE_TEXT=""
REQUIRE_INTERNAL=true
DISCLOSURE_ENABLED=true
POLICY_PATH=""
CALENDAR_DOMAINS=""
CALENDAR_ADDRESSES=""
CALENDAR_REQUIRE_INTERNAL=""
CALENDAR_DOMAINS_SET=false
CALENDAR_ADDRESSES_SET=false

usage() {
  echo "Usage: sudo $0 [--domains a.de,b.de] [--addresses a@b.example,c@d.example] [--calendar-domains a.de,b.de] [--calendar-addresses a@b.example] [--calendar-restrict|--calendar-no-restrict] --disclosure-text \"...\" [--no-restrict] [--no-disclosure] [--path FILE]"
  exit 2
}

while [ $# -gt 0 ]; do
  case "$1" in
    --domains) DOMAINS="${2:?}"; shift 2 ;;
    --addresses) ADDRESSES="${2:?}"; shift 2 ;;
    --calendar-domains) CALENDAR_DOMAINS="${2:?}"; CALENDAR_DOMAINS_SET=true; shift 2 ;;
    --calendar-addresses) CALENDAR_ADDRESSES="${2:?}"; CALENDAR_ADDRESSES_SET=true; shift 2 ;;
    --calendar-restrict) CALENDAR_REQUIRE_INTERNAL=true; shift ;;
    --calendar-no-restrict) CALENDAR_REQUIRE_INTERNAL=false; shift ;;
    --disclosure-text) DISCLOSURE_TEXT="${2:?}"; shift 2 ;;
    --no-restrict) REQUIRE_INTERNAL=false; shift ;;
    --no-disclosure) DISCLOSURE_ENABLED=false; shift ;;
    --path) POLICY_PATH="${2:?}"; shift 2 ;;
    -h|--help) usage ;;
    *) echo "Unknown argument: $1" >&2; usage ;;
  esac
done

if [ "$CALENDAR_DOMAINS_SET" = false ]; then
  CALENDAR_DOMAINS="$DOMAINS"
fi
if [ "$CALENDAR_ADDRESSES_SET" = false ]; then
  CALENDAR_ADDRESSES="$ADDRESSES"
fi
if [ -z "$CALENDAR_REQUIRE_INTERNAL" ]; then
  CALENDAR_REQUIRE_INTERNAL="$REQUIRE_INTERNAL"
fi

if [ "$REQUIRE_INTERNAL" = true ] && [ -z "$DOMAINS" ] && [ -z "$ADDRESSES" ]; then
  echo "Error: --domains or --addresses is required unless --no-restrict is given." >&2; exit 1
fi
if [ "$CALENDAR_REQUIRE_INTERNAL" = true ] && [ -z "$CALENDAR_DOMAINS" ] && [ -z "$CALENDAR_ADDRESSES" ]; then
  echo "Error: --calendar-domains or --calendar-addresses is required unless --calendar-no-restrict is given." >&2; exit 1
fi
if [ "$DISCLOSURE_ENABLED" = true ] && [ -z "$DISCLOSURE_TEXT" ]; then
  echo "Error: --disclosure-text is required unless --no-disclosure is given." >&2; exit 1
fi
if [ "$(id -u)" -ne 0 ]; then
  echo "Error: run as root (e.g. via sudo) so the file becomes admin-owned." >&2; exit 1
fi
if ! command -v python3 >/dev/null; then
  echo "Error: python3 is required to build the JSON safely." >&2; exit 1
fi

if [ -z "$POLICY_PATH" ]; then
  case "$(uname -s)" in
    Darwin) POLICY_PATH="/Library/Application Support/microsoft-mcp/policy.json" ;;
    *)      POLICY_PATH="/etc/microsoft-mcp/policy.json" ;;
  esac
fi

mkdir -p "$(dirname "$POLICY_PATH")"
TEMP_PATH="$(mktemp "${POLICY_PATH}.tmp.XXXXXX")"
trap 'rm -f "$TEMP_PATH"' EXIT
DOMAINS="$DOMAINS" ADDRESSES="$ADDRESSES" DISCLOSURE_TEXT="$DISCLOSURE_TEXT" \
REQUIRE_INTERNAL="$REQUIRE_INTERNAL" DISCLOSURE_ENABLED="$DISCLOSURE_ENABLED" \
CALENDAR_DOMAINS="$CALENDAR_DOMAINS" CALENDAR_ADDRESSES="$CALENDAR_ADDRESSES" \
CALENDAR_REQUIRE_INTERNAL="$CALENDAR_REQUIRE_INTERNAL" \
POLICY_PATH="$TEMP_PATH" python3 - <<'EOF'
import json, os

domains = [d.strip() for d in os.environ["DOMAINS"].split(",") if d.strip()]
addresses = [a.strip() for a in os.environ["ADDRESSES"].split(",") if a.strip()]
calendar_domains = [d.strip() for d in os.environ["CALENDAR_DOMAINS"].split(",") if d.strip()]
calendar_addresses = [a.strip() for a in os.environ["CALENDAR_ADDRESSES"].split(",") if a.strip()]
policy = {
  "version": 1,
  "outlook": {
    "requireInternalRecipients": os.environ["REQUIRE_INTERNAL"].lower() == "true",
    "allowedRecipientDomains": domains,
    "allowedRecipientAddresses": addresses,
    "aiDisclosureEnabled": os.environ["DISCLOSURE_ENABLED"].lower() == "true",
    "aiDisclosureText": os.environ["DISCLOSURE_TEXT"],
  },
  "calendar": {
    "requireInternalAttendees": os.environ["CALENDAR_REQUIRE_INTERNAL"].lower() == "true",
    "allowedAttendeeDomains": calendar_domains,
    "allowedAttendeeAddresses": calendar_addresses,
  },
  "teams": {
    "requireInternalRecipients": os.environ["REQUIRE_INTERNAL"].lower() == "true",
    "allowedRecipientDomains": domains,
    "allowedRecipientAddresses": addresses,
    "aiDisclosureEnabled": os.environ["DISCLOSURE_ENABLED"].lower() == "true",
    "aiDisclosureText": os.environ["DISCLOSURE_TEXT"],
  },
}
with open(os.environ["POLICY_PATH"], "w", encoding="utf-8") as f:
    json.dump(policy, f, ensure_ascii=False, indent=2)
    f.write("\n")
EOF

if [ "$(uname -s)" = "Darwin" ]; then
  chown root:wheel "$TEMP_PATH"
else
  chown root:root "$TEMP_PATH"
fi
chmod 644 "$TEMP_PATH"
mv -f "$TEMP_PATH" "$POLICY_PATH"
trap - EXIT

if [ "$(uname -s)" = "Darwin" ]; then
  chown root:wheel "$POLICY_PATH"
else
  chown root:root "$POLICY_PATH"
fi
chmod 644 "$POLICY_PATH"
echo "Wrote $POLICY_PATH (root-owned, 644):"
cat "$POLICY_PATH"

# Verify: the invoking (non-root) user must NOT be able to write the file.
if [ -n "${SUDO_USER:-}" ] && [ "$SUDO_USER" != "root" ]; then
  if sudo -u "$SUDO_USER" test -w "$POLICY_PATH"; then
    echo "Protection FAILED: $SUDO_USER can still write $POLICY_PATH." >&2; exit 1
  fi
  echo "Verified: $SUDO_USER cannot write $POLICY_PATH."
else
  echo "Warning: could not determine a non-root user; verify writability manually." >&2
fi
