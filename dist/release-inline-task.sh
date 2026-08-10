#!/usr/bin/env bash
set -Eeuo pipefail
trap 'c=$?; echo "[ERROR] exit=$c line=$LINENO cmd=${BASH_COMMAND:-na}"; exit $c' ERR
require(){ [ -n "${!1:-}" ] || { echo "##[error] $1 is required"; exit 2; }; }

# -------------------- Configuration (edit if needed) --------------------
TOOLS_PROJECT="ShonizCollection"       # project that contains the shared repo
TOOLS_REPO="SharedTemplates"           # repository that holds release-komodo.sh
TOOLS_BRANCH="main"                    # branch where the script lives
SCRIPT_PATH="release-komodo.sh"        # path to the script inside the repo
# ----------------------------------------------------------------------
: "${AZP_TOKEN:=$(AZP_TOKEN)}"
: "${KOMODO_API_KEY:=$(KOMODO_API_KEY)}"
: "${KOMODO_API_SECRET:=$(KOMODO_API_SECRET)}"
: "${KOMODO_ADDRESS:=https://komodo.buluttakin.com}"
require AZP_TOKEN
# pick token: explicit AZP_TOKEN (secret) preferred, fallback to pipeline System.AccessToken
if [ -n "${AZP_TOKEN:-}" ]; then
  TOK="$AZP_TOKEN"
elif [ -n "${SYSTEM_ACCESSTOKEN:-}" ]; then
  TOK="$SYSTEM_ACCESSTOKEN"
else
  echo "##[error] AZP_TOKEN not found. Either:"
  echo "  * Provide AZP_TOKEN as a pipeline/release variable (secret), or"
  echo "  * Enable 'Allow scripts to access the OAuth token' for the job and retry."
  exit 1
fi

COLL_BASE="https://azure.buluttakin.com"

# Build URL parts
TOOLS_PROJECT_URLENC="${TOOLS_PROJECT// /%20}"
# NOTE: for Azure DevOps Server (on-prem TFS) the path format is the same but make sure COLLECTION is included
CLONE_URL="${COLL_BASE}/${TOOLS_PROJECT_URLENC}/_git/${TOOLS_REPO}"
echo "Clone URL: ${CLONE_URL}"

# create temp workdir and cleanup on exit
WORKDIR="$(mktemp -d /tmp/sharedscript.XXXXXX)"
cleanup(){ rc=$?; rm -rf "$WORKDIR" >/dev/null 2>&1 || true; exit $rc; }
trap cleanup EXIT

_auth_raw="$(printf '%s:%s' "pat" "$TOK")"
if command -v base64 >/dev/null 2>&1; then
  AUTH_HDR="$(printf '%s' "$_auth_raw" | base64 | tr -d '\n')"
else
  echo "##[warning] base64 not available; auth header may fail"
  AUTH_HDR=""
fi
unset _auth_raw

# try shallow git clone first (fast). If it fails, fallback to REST API single-file download.
if command -v git >/dev/null 2>&1; then
  # Keep the secret-derived header out of xtrace and process arguments.
  GIT_CONFIG_COUNT=1 \
  GIT_CONFIG_KEY_0=http.extraHeader \
  GIT_CONFIG_VALUE_0="Authorization: Basic ${AUTH_HDR}" \
    git clone --depth 1 --branch "$TOOLS_BRANCH" "$CLONE_URL" "$WORKDIR" || GIT_CLONE_FAIL=1
else
  echo "##[warning] git not found on agent; skipping git clone"
  GIT_CLONE_FAIL=1
fi

TARGET="$WORKDIR/$SCRIPT_PATH"

if [ -z "${GIT_CLONE_FAIL:-}" ] && [ -f "$TARGET" ]; then
  echo "##[section]Script found in git clone: $TARGET"
else
  echo "##[warning] git clone did not produce script; falling back to REST API fetch"

  fetch_path="/${SCRIPT_PATH#/}"
  # percent-encode $ as %24 to avoid shell expansion
  api_url="${COLL_BASE}/${TOOLS_PROJECT_URLENC}/_apis/git/repositories/${TOOLS_REPO}/items?path=${fetch_path}&versionDescriptor.version=${TOOLS_BRANCH}&api-version=6.0&%24format=text"
  echo "REST items URL: ${api_url}"

  if [ -n "$AUTH_HDR" ]; then
    # Authorization: Basic <base64(pat:TOKEN)>
    http_code=0
    if command -v curl >/dev/null 2>&1; then
      printf 'header = "Authorization: Basic %s"\n' "$AUTH_HDR" |
        curl --config - -sS -f "$api_url" -o "$WORKDIR/$(basename "$SCRIPT_PATH")" || http_code=$?
    elif command -v wget >/dev/null 2>&1; then
      wget_config="$WORKDIR/.wgetrc"
      (umask 077; printf 'header = Authorization: Basic %s\n' "$AUTH_HDR" > "$wget_config")
      WGETRC="$wget_config" wget -O "$WORKDIR/$(basename "$SCRIPT_PATH")" "$api_url" || http_code=$?
      : > "$wget_config"
      rm -f "$wget_config"
    else
      echo "##[error] Neither curl nor wget available to fetch script"; exit 2
    fi

    if [ "$http_code" != "0" ]; then
      echo "##[error] Failed to download script via REST API (curl/wget exit=$http_code)"; exit 3
    fi

    TARGET="$WORKDIR/$(basename "$SCRIPT_PATH")"
    if [ ! -f "$TARGET" ]; then
      echo "##[error] REST download did not produce file: $TARGET"; exit 4
    fi
    echo "##[section]Script downloaded to: $TARGET"
  else
    echo "##[error] Could not build auth header for REST API fallback (base64 missing)"; exit 5
  fi
fi

# Run the script in a subshell for minimal env leakage. Ensure executable.
chmod +x "$TARGET"
echo "##[section]Executing shared script..."
export AZP_TOKEN="$TOK"
export KOMODO_API_KEY="$KOMODO_API_KEY"
export KOMODO_API_SECRET="$KOMODO_API_SECRET"
export KOMODO_ADDRESS="$KOMODO_ADDRESS"
export SYSTEM_ACCESSTOKEN="$TOK"

# execute and capture exit code
( set -x; "$TARGET" )
RC=$?
echo "##[section]Shared script exited with code: $RC"

# cleanup and exit (cleanup trap handles removal)
exit "$RC"
