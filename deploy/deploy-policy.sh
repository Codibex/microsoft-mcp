#!/usr/bin/env bash
# Deploys the admin-owned recipient policy.json (domains/addresses + AI disclosure).
#
# Builds policy.json from parameters, writes it to the admin-owned system location
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
#   sudo ./deploy-policy.sh --domains firma.de --disclosure-text "..." --path /custom/policy.json
set -euo pipefail

DOMAINS=""
ADDRESSES=""
DISCLOSURE_TEXT=""
REQUIRE_INTERNAL=true
DISCLOSURE_ENABLED=true
POLICY_PATH=""

usage() {
  echo "Usage: sudo $0 [--domains a.de,b.de] [--addresses a@b.example,c@d.example] --disclosure-text \"...\" [--no-restrict] [--no-disclosure] [--path FILE]"
  exit 2
}

while [ $# -gt 0 ]; do
  case "$1" in
    --domains) DOMAINS="${2:?}"; shift 2 ;;
    --addresses) ADDRESSES="${2:?}"; shift 2 ;;
    --disclosure-text) DISCLOSURE_TEXT="${2:?}"; shift 2 ;;
    --no-restrict) REQUIRE_INTERNAL=false; shift ;;
    --no-disclosure) DISCLOSURE_ENABLED=false; shift ;;
    --path) POLICY_PATH="${2:?}"; shift 2 ;;
    -h|--help) usage ;;
    *) echo "Unknown argument: $1" >&2; usage ;;
  esac
done

if [ "$REQUIRE_INTERNAL" = true ] && [ -z "$DOMAINS" ] && [ -z "$ADDRESSES" ]; then
  echo "Error: --domains or --addresses is required unless --no-restrict is given." >&2; exit 1
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
DOMAINS="$DOMAINS" ADDRESSES="$ADDRESSES" DISCLOSURE_TEXT="$DISCLOSURE_TEXT" \
REQUIRE_INTERNAL="$REQUIRE_INTERNAL" DISCLOSURE_ENABLED="$DISCLOSURE_ENABLED" \
POLICY_PATH="$POLICY_PATH" python3 - <<'EOF'
import json, os

domains = [d.strip() for d in os.environ["DOMAINS"].split(",") if d.strip()]
addresses = [a.strip() for a in os.environ["ADDRESSES"].split(",") if a.strip()]
policy = {
    "requireInternalRecipients": os.environ["REQUIRE_INTERNAL"].lower() == "true",
    "allowedRecipientDomains": domains,
    "allowedRecipientAddresses": addresses,
    "aiDisclosureEnabled": os.environ["DISCLOSURE_ENABLED"].lower() == "true",
    "aiDisclosureText": os.environ["DISCLOSURE_TEXT"],
}
with open(os.environ["POLICY_PATH"], "w", encoding="utf-8") as f:
    json.dump(policy, f, ensure_ascii=False, indent=2)
    f.write("\n")
EOF

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
