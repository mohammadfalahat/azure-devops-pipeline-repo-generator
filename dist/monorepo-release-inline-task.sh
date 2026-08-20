#!/usr/bin/env bash
set -Eeuo pipefail
trap 'code=$?; echo "##[error] MR release failed at line $LINENO (exit=$code)"; exit "$code"' ERR

require() { [ -n "${!1:-}" ] || { echo "##[error] $1 is required"; exit 2; }; }
: "${AZP_TOKEN:=$(AZP_TOKEN)}"
: "${KOMODO_API_KEY:=$(KOMODO_API_KEY)}"
: "${KOMODO_API_SECRET:=$(KOMODO_API_SECRET)}"
: "${KOMODO_ADDRESS:=https://komodo.buluttakin.com}"
require AZP_TOKEN
require KOMODO_API_KEY
require KOMODO_API_SECRET
require SYSTEM_DEFAULTWORKINGDIRECTORY

for command_name in node curl jq; do
  command -v "$command_name" >/dev/null 2>&1 || {
    echo "##[error] $command_name is required on the Release agent"; exit 2;
  }
done

manifest="$(find "$SYSTEM_DEFAULTWORKINGDIRECTORY" -type f -name manifest.json -path '*/mr-drop/*' -print -quit)"
[ -f "$manifest" ] || { echo "##[error] mr-drop/manifest.json was not downloaded"; exit 3; }
artifact_dir="$(dirname "$manifest")"

manifest_value() {
  node -e '
    const fs = require("fs");
    const value = JSON.parse(fs.readFileSync(process.argv[1], "utf8"))[process.argv[2]];
    if (value === undefined || value === null) process.exit(4);
    process.stdout.write(String(value));
  ' "$manifest" "$1"
}

build_id="$(manifest_value buildId)"
project_key="$(manifest_value project)"
environment="$(manifest_value environment)"
collection_uri="$(manifest_value collectionUri)"
azure_project_id="$(manifest_value azureProjectId)"
komodo_server="$(manifest_value komodoServer)"
komodo_stack="$(manifest_value komodoStack)"
deployment_root="$(manifest_value deploymentRoot)"
static_container="$(manifest_value staticContainer)"
bff_container="$(manifest_value bffContainer || true)"
bff_project="$(manifest_value bffProject || true)"
bff_entry="$(manifest_value bffEntry)"
compose_repository="$(manifest_value composeRepository)"
compose_path="$(manifest_value composePath)"

[[ "$build_id" =~ ^[0-9]+$ ]] || { echo "##[error] Unsafe build ID"; exit 4; }
[[ "$project_key" =~ ^[a-z0-9][a-z0-9-]*$ ]] || { echo "##[error] Unsafe project key"; exit 4; }
[[ "$environment" =~ ^[a-z0-9][a-z0-9-]*$ ]] || { echo "##[error] Unsafe environment"; exit 4; }
[[ "$komodo_server" =~ ^[A-Za-z0-9][A-Za-z0-9._:-]*$ ]] || { echo "##[error] Unsafe Komodo server"; exit 4; }
[[ "$komodo_stack" =~ ^[A-Za-z0-9][A-Za-z0-9._-]*$ ]] || { echo "##[error] Unsafe Komodo Stack"; exit 4; }
[[ "$deployment_root" =~ ^/mnt/graid/projects/[A-Za-z0-9._/-]+/monorepo/[a-z0-9][a-z0-9-]*$ ]] || {
  echo "##[error] Deployment root is outside the managed monorepo path"; exit 4;
}
[[ "$static_container" =~ ^[a-z0-9][a-z0-9_-]*$ ]] || {
  echo "##[error] Unsafe static container name"; exit 4;
}
[[ -z "$bff_container" || "$bff_container" =~ ^[a-z0-9][a-z0-9_-]*$ ]] || {
  echo "##[error] Unsafe BFF container name"; exit 4;
}
[[ -z "$bff_project" || "$bff_project" =~ ^[A-Za-z0-9][A-Za-z0-9._-]*$ ]] || {
  echo "##[error] Unsafe BFF project name"; exit 4;
}
[[ "$bff_entry" =~ ^[A-Za-z0-9][A-Za-z0-9._/-]*$ && "$bff_entry" != *..* ]] || {
  echo "##[error] Unsafe BFF entry path"; exit 4;
}
[[ "$compose_repository" =~ ^[A-Za-z0-9][A-Za-z0-9._-]*$ ]] || {
  echo "##[error] Unsafe Compose repository"; exit 4;
}
[[ "$compose_path" =~ ^/[A-Za-z0-9._/-]+/compose\.yml$ && "$compose_path" != *..* ]] || {
  echo "##[error] Unsafe Compose path"; exit 4;
}

auth_b64="$(printf 'pat:%s' "$AZP_TOKEN" | base64 | tr -d '\r\n')"
artifact_url="${collection_uri%/}/${azure_project_id}/_apis/build/builds/${build_id}/artifacts?artifactName=mr-drop&%24format=zip&api-version=6.0"
bff_changed=0
if [ -n "$bff_project" ] && awk -F '\t' -v name="$bff_project" '$1 == name && $2 == "bff" { found=1 } END { exit !found }' "$artifact_dir/modules.tsv"; then
  bff_changed=1
fi

workdir="$(mktemp -d /tmp/mr-release.XXXXXX)"
cleanup() {
  code=$?
  trap - EXIT
  find "$workdir" -type f -exec sh -c ': > "$1"' _ {} \; >/dev/null 2>&1 || true
  rm -rf "$workdir" >/dev/null 2>&1 || true
  unset auth_b64
  exit "$code"
}
trap cleanup EXIT

komodo_call() {
  endpoint="$1"
  request_file="$2"
  response_file="$3"
  status="$({
    printf 'header = "X-Api-Key: %s"\n' "$KOMODO_API_KEY"
    printf 'header = "X-Api-Secret: %s"\n' "$KOMODO_API_SECRET"
    printf 'header = "Content-Type: application/json"\n'
  } | curl --config - -sS -o "$response_file" -w '%{http_code}' \
    --data-binary "@$request_file" "${KOMODO_ADDRESS%/}/$endpoint")"
  if ! [[ "$status" =~ ^2[0-9][0-9]$ ]]; then
    message="$(jq -r '.error // .message // "Komodo request failed"' "$response_file" 2>/dev/null || true)"
    echo "##[error] Komodo $endpoint returned HTTP $status: $message" >&2
    return 1
  fi
}

run_terminal() {
  phase="$1"
  command_text="$2"
  request_file="$workdir/terminal-$phase.json"
  terminal_log="$workdir/terminal-$phase.log"
  export MR_TERMINAL_PHASE="$phase" MR_TERMINAL_COMMAND="$command_text" MR_TERMINAL_SERVER="$komodo_server" MR_TERMINAL_BUILD="$build_id"
  node - "$request_file" <<'NODE'
const fs = require('fs');
fs.writeFileSync(process.argv[2], JSON.stringify({
  server: process.env.MR_TERMINAL_SERVER,
  terminal: `ado-mr-${process.env.MR_TERMINAL_BUILD}-${process.env.MR_TERMINAL_PHASE}`,
  command: process.env.MR_TERMINAL_COMMAND
}));
NODE
  unset MR_TERMINAL_COMMAND
  if ! {
    printf 'header = "X-Api-Key: %s"\n' "$KOMODO_API_KEY"
    printf 'header = "X-Api-Secret: %s"\n' "$KOMODO_API_SECRET"
    printf 'header = "Content-Type: application/json"\n'
  } | curl --config - -fsSN --data-binary "@$request_file" "${KOMODO_ADDRESS%/}/terminal/execute" | tee "$terminal_log"; then
    return 1
  fi
  grep -q '__KOMODO_EXIT_CODE__:0' "$terminal_log"
}

deploy_stack() {
  attempt_name="$1"
  execute_request="$workdir/deploy-$attempt_name-request.json"
  update_file="$workdir/deploy-$attempt_name-update.json"
  jq -cn --arg stack "$komodo_stack" '{type:"DeployStack",params:{stack:$stack}}' > "$execute_request"
  komodo_call execute "$execute_request" "$update_file" || return 1

  update_id="$(jq -r '._id["$oid"] // .id["$oid"] // .id // empty' "$update_file")"
  status="$(jq -r '.status // empty' "$update_file")"
  polls=0
  while [ "$status" != Complete ]; do
    [ -n "$update_id" ] || { echo "##[error] DeployStack returned no update ID" >&2; return 1; }
    polls=$((polls + 1))
    [ "$polls" -le 300 ] || { echo "##[error] Timed out waiting for Komodo Stack deployment" >&2; return 1; }
    sleep 2
    read_request="$workdir/deploy-$attempt_name-read.json"
    jq -cn --arg id "$update_id" '{type:"GetUpdate",params:{id:$id}}' > "$read_request"
    komodo_call read "$read_request" "$update_file" || return 1
    status="$(jq -r '.status // empty' "$update_file")"
  done
  if ! jq -e '.success == true' "$update_file" >/dev/null; then
    echo "##[error] Komodo DeployStack failed for $komodo_stack" >&2
    jq -r '.logs[-10:][]? | if type == "object" then ((.stage // "Komodo") + ": " + ((.message // .stdout // .stderr // "operation failed") | tostring)) else tostring end' "$update_file" >&2 || true
    return 1
  fi
  echo "##[section]Komodo Stack deployed from ADO Git: $komodo_stack ($compose_repository$compose_path)"
}

export MR_REMOTE_ROOT="$deployment_root"
export MR_REMOTE_BUILD_ID="$build_id"
export MR_REMOTE_ARTIFACT_URL="$artifact_url"
export MR_REMOTE_AUTH_B64="$auth_b64"
export MR_REMOTE_BFF_CONTAINER="$bff_container"
export MR_REMOTE_BFF_CHANGED="$bff_changed"
export MR_REMOTE_PROJECT="$project_key"
export MR_REMOTE_ENVIRONMENT="$environment"
export MR_REMOTE_STATIC_CONTAINER="$static_container"

prepare_command="$(node <<'NODE'
const q = (value) => `'${String(value).replaceAll("'", `'\"'\"'`)}'`;
const script = `set -Eeuo pipefail
root=${q(process.env.MR_REMOTE_ROOT)}
build_id=${q(process.env.MR_REMOTE_BUILD_ID)}
artifact_url=${q(process.env.MR_REMOTE_ARTIFACT_URL)}
auth_b64=${q(process.env.MR_REMOTE_AUTH_B64)}
command -v curl >/dev/null 2>&1
command -v unzip >/dev/null 2>&1
tmp=$(mktemp -d /tmp/mr-deploy.XXXXXX)
trap 'rm -rf "$tmp"' EXIT
printf 'header = "Authorization: Basic %s"\\n' "$auth_b64" | curl --config - -fsSL "$artifact_url" -o "$tmp/drop.zip"
unzip -q "$tmp/drop.zip" -d "$tmp/drop"
manifest=$(find "$tmp/drop" -type f -name manifest.json -path '*/mr-drop/*' -print -quit)
[ -f "$manifest" ]
artifact_dir=$(dirname "$manifest")
nginx_config_source="$artifact_dir/runtime/nginx/default.conf"
[ -s "$nginx_config_source" ] || { echo "Missing packaged Nginx runtime configuration" >&2; exit 12; }
release="$root/releases/$build_id"
mkdir -p "$root/releases"
if [ -d "$root/current" ] && [ ! -L "$root/current" ]; then
  legacy="$root/releases/legacy-before-$build_id"
  rm -rf "$legacy"
  mv "$root/current" "$legacy"
  ln -s "releases/legacy-before-$build_id" "$root/current"
fi
previous_inventory="$tmp/previous-inventory.tsv"
if [ -f "$root/current/inventory.tsv" ]; then cp "$root/current/inventory.tsv" "$previous_inventory"; fi
rm -rf "$release"
mkdir -p "$release/root" "$release/modules"
if [ -L "$root/current" ] || [ -d "$root/current" ]; then
  cp -alL "$root/current/." "$release/" 2>/dev/null || cp -aL "$root/current/." "$release/"
fi
while IFS=$(printf '\\t') read -r name kind route; do
  [ -n "$name" ] || continue
  case "$name" in (*[!A-Za-z0-9._-]*|'') echo "Unsafe module name: $name" >&2; exit 11;; esac
  source="$artifact_dir/modules/$name"
  [ -d "$source" ] || { echo "Missing packaged module: $name" >&2; exit 12; }
  if [ "$kind" = shell ]; then
    rm -rf "$release/root"
    mkdir -p "$release/root"
    cp -a "$source/." "$release/root/"
  else
    rm -rf "$release/modules/$name"
    mkdir -p "$release/modules/$name"
    cp -a "$source/." "$release/modules/$name/"
  fi
done < "$artifact_dir/modules.tsv"
if [ -f "$previous_inventory" ]; then
  while IFS=$(printf '\\t') read -r old_name old_kind old_route; do
    [ -n "$old_name" ] || continue
    if ! awk -F '\\t' -v name="$old_name" '$1 == name { found=1 } END { exit !found }' "$artifact_dir/inventory.tsv"; then
      echo "MR orphan retained for manual review: $old_name"
    fi
  done < "$previous_inventory"
fi
for inventory_file in "$previous_inventory" "$artifact_dir/inventory.tsv"; do
  [ -f "$inventory_file" ] || continue
  while IFS=$(printf '\\t') read -r module_name module_kind module_route; do
    [ "$module_kind" = static ] || continue
    case "$module_name" in (*[!A-Za-z0-9._-]*|'') echo "Unsafe static module name: $module_name" >&2; exit 11;; esac
    [ -d "$release/modules/$module_name" ] || continue
    rm -rf "$release/root/$module_name"
    ln -s "../modules/$module_name" "$release/root/$module_name"
  done < "$inventory_file"
done
cp "$artifact_dir/inventory.tsv" "$release/inventory.tsv"
cp "$manifest" "$release/manifest.json"
docker run --rm --entrypoint nginx -v "$nginx_config_source:/etc/nginx/conf.d/default.conf:ro" nginx:1.27-alpine -t
rollback_state="$root/.mr-rollback/$build_id"
rm -rf "$rollback_state"
mkdir -p "$rollback_state"
if previous=$(readlink "$root/current" 2>/dev/null); then printf '%s' "$previous" > "$rollback_state/current-link"; else : > "$rollback_state/no-current"; fi
nginx_config_target="$root/runtime/nginx/default.conf"
mkdir -p "$root/runtime/nginx"
if [ -f "$nginx_config_target" ]; then cp "$nginx_config_target" "$rollback_state/nginx-default.conf"; else : > "$rollback_state/no-nginx"; fi
if [ -f "$nginx_config_target" ]; then cat "$nginx_config_source" > "$nginx_config_target"; else cp "$nginx_config_source" "$nginx_config_target"; fi
chmod 0644 "$nginx_config_target"
link="$root/.current-$build_id"
rm -f "$link"
ln -s "releases/$build_id" "$link"
mv -Tf "$link" "$root/current"
echo "MR deployment $build_id prepared; Komodo Stack will activate the Git-managed Compose"`;
process.stdout.write(`bash -lc ${q(script)}`);
NODE
)"

rollback_command="$(node <<'NODE'
const q = (value) => `'${String(value).replaceAll("'", `'\"'\"'`)}'`;
const script = `set -Eeuo pipefail
root=${q(process.env.MR_REMOTE_ROOT)}
build_id=${q(process.env.MR_REMOTE_BUILD_ID)}
state="$root/.mr-rollback/$build_id"
[ -d "$state" ] || { echo "No MR rollback state exists"; exit 0; }
if [ -s "$state/current-link" ]; then
  rollback="$root/.rollback-$build_id"
  rm -f "$rollback"
  ln -s "$(cat "$state/current-link")" "$rollback"
  mv -Tf "$rollback" "$root/current"
else
  rm -f "$root/current"
fi
target="$root/runtime/nginx/default.conf"
if [ -f "$state/nginx-default.conf" ]; then cat "$state/nginx-default.conf" > "$target"; else rm -f "$target"; fi
echo "MR current symlink and runtime Nginx configuration rolled back"`;
process.stdout.write(`bash -lc ${q(script)}`);
NODE
)"

post_deploy_command="$(node <<'NODE'
const q = (value) => `'${String(value).replaceAll("'", `'\"'\"'`)}'`;
const script = `set -Eeuo pipefail
project_key=${q(process.env.MR_REMOTE_PROJECT)}
environment=${q(process.env.MR_REMOTE_ENVIRONMENT)}
bff_container=${q(process.env.MR_REMOTE_BFF_CONTAINER || '')}
bff_changed=${q(process.env.MR_REMOTE_BFF_CHANGED || '0')}
static_container=${q(process.env.MR_REMOTE_STATIC_CONTAINER)}
docker inspect "$static_container" >/dev/null 2>&1
docker exec "$static_container" nginx -t
docker exec "$static_container" nginx -s reload
if [ "$bff_changed" = 1 ] && [ -n "$bff_container" ] && docker inspect "$bff_container" >/dev/null 2>&1; then
  docker restart "$bff_container"
fi
echo "MR runtime validated after Komodo Stack deployment"`;
process.stdout.write(`bash -lc ${q(script)}`);
NODE
)"

cleanup_command="$(node <<'NODE'
const q = (value) => `'${String(value).replaceAll("'", `'\"'\"'`)}'`;
const script = `set -Eeuo pipefail
root=${q(process.env.MR_REMOTE_ROOT)}
build_id=${q(process.env.MR_REMOTE_BUILD_ID)}
rm -rf "$root/.mr-rollback/$build_id"`;
process.stdout.write(`bash -lc ${q(script)}`);
NODE
)"

if ! run_terminal prepare "$prepare_command"; then
  run_terminal rollback-prepare "$rollback_command" || true
  echo "##[error] MR artifact preparation failed" >&2
  exit 20
fi
unset MR_REMOTE_AUTH_B64 auth_b64

if ! deploy_stack current; then
  run_terminal rollback-deploy "$rollback_command" || true
  deploy_stack previous-after-deploy-failure || true
  echo "##[error] Komodo Stack deployment failed; runtime state was rolled back" >&2
  exit 21
fi

if ! run_terminal validate "$post_deploy_command"; then
  run_terminal rollback-validation "$rollback_command" || true
  deploy_stack previous-after-validation-failure || true
  echo "##[error] MR runtime validation failed; previous state was restored through Komodo Stack" >&2
  exit 22
fi

run_terminal cleanup "$cleanup_command"
echo "##[section]MR artifact $build_id deployed as Komodo GitOps Stack $komodo_stack on $komodo_server"
