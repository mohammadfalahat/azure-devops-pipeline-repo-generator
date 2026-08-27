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

for command_name in git curl jq awk base64; do
  command -v "$command_name" >/dev/null 2>&1 || {
    echo "##[error] $command_name is required on the Release agent"; exit 2;
  }
done

manifest="$(find "$SYSTEM_DEFAULTWORKINGDIRECTORY" -type f -name manifest.json -path '*/mr-drop/*' -print -quit)"
[ -f "$manifest" ] || { echo "##[error] mr-drop/manifest.json was not downloaded"; exit 3; }

manifest_value() {
  jq -er --arg key "$1" '
    if has($key) and .[$key] != null
    then .[$key] | tostring
    else error("manifest field is missing: " + $key)
    end
  ' "$manifest"
}

manifest_optional() {
  jq -r --arg key "$1" 'if has($key) and .[$key] != null then .[$key] | tostring else "" end' "$manifest"
}

schema_version="$(manifest_value schemaVersion)"
deployment_mode="$(manifest_value deploymentMode)"
build_id="$(manifest_value buildId)"
collection_uri="$(manifest_value collectionUri)"
azure_project_id="$(manifest_value azureProjectId)"
komodo_stack="$(manifest_value komodoStack)"
compose_repository="$(manifest_value composeRepository)"
compose_path="$(manifest_value composePath)"
service="$(manifest_value service)"
static_container="$(manifest_value staticContainer)"
bff_container="$(manifest_optional bffContainer)"
static_image="$(manifest_value staticImage)"
bff_image="$(manifest_optional bffImage)"
static_repository="$(manifest_value staticImageRepository)"
bff_repository="$(manifest_value bffImageRepository)"

[ "$schema_version" = 2 ] || { echo "##[error] Unsupported MR manifest schema: $schema_version"; exit 4; }
[ "$deployment_mode" = immutable-images ] || { echo "##[error] Unsupported MR deployment mode: $deployment_mode"; exit 4; }
[[ "$build_id" =~ ^[0-9]+$ ]] || { echo "##[error] Unsafe build ID"; exit 4; }
[[ "$komodo_stack" =~ ^[A-Za-z0-9][A-Za-z0-9._-]*$ ]] || { echo "##[error] Unsafe Komodo Stack"; exit 4; }
[[ "$compose_repository" =~ ^[A-Za-z0-9][A-Za-z0-9._-]*$ ]] || { echo "##[error] Unsafe Compose repository"; exit 4; }
[[ "$service" =~ ^[a-z0-9][a-z0-9-]*$ ]] || { echo "##[error] Unsafe MR service key"; exit 4; }
[[ "$compose_path" =~ ^/[A-Za-z0-9._/-]+/compose\.yml$ && "$compose_path" != *..* ]] || {
  echo "##[error] Unsafe Compose path"; exit 4;
}
[[ "$static_container" =~ ^[a-z0-9][a-z0-9_-]*$ ]] || { echo "##[error] Unsafe static container"; exit 4; }
[[ -z "$bff_container" || "$bff_container" =~ ^[a-z0-9][a-z0-9_-]*$ ]] || {
  echo "##[error] Unsafe BFF container"; exit 4;
}
image_pattern='^[a-z0-9.-]+(:[0-9]+)?(/[a-z0-9._-]+)+(:[A-Za-z0-9_.-]+|@sha256:[a-f0-9]{64})$'
repository_pattern='^[a-z0-9.-]+(:[0-9]+)?(/[a-z0-9._-]+)+$'
[[ "$static_image" =~ $image_pattern ]] || { echo "##[error] Unsafe static image reference"; exit 4; }
[[ -z "$bff_image" || "$bff_image" =~ $image_pattern ]] || { echo "##[error] Unsafe BFF image reference"; exit 4; }
[[ "$static_repository" =~ $repository_pattern ]] || { echo "##[error] Unsafe static image repository"; exit 4; }
[[ "$bff_repository" =~ $repository_pattern ]] || { echo "##[error] Unsafe BFF image repository"; exit 4; }
if [ -n "$bff_image" ] && [ -z "$bff_container" ]; then
  echo "##[error] The manifest has a BFF image but no BFF Compose service"; exit 4
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

auth_b64="$(printf 'pat:%s' "$AZP_TOKEN" | (base64 -w0 2>/dev/null || base64 | tr -d '\r\n'))"
git_authenticated() {
  GIT_CONFIG_COUNT=1 \
  GIT_CONFIG_KEY_0=http.extraHeader \
  GIT_CONFIG_VALUE_0="Authorization: Basic ${auth_b64}" \
    git "$@"
}

repo_url="${collection_uri%/}/${azure_project_id}/_git/${compose_repository}"
repo_dir="$workdir/repository"
echo "##[section]Cloning Compose repository: $compose_repository/main"
git_authenticated clone --quiet --depth 1 --branch main "$repo_url" "$repo_dir"
git -C "$repo_dir" config user.name az-release
git -C "$repo_dir" config user.email azrelease@local
compose_file="$repo_dir/${compose_path#/}"
[ -s "$compose_file" ] || { echo "##[error] Compose file was not found: $compose_repository$compose_path"; exit 5; }
env_path="${compose_path%/compose.yml}/.env"
env_file="$repo_dir/${env_path#/}"
static_tag_key="${service//-/_}"
bff_tag_key="${static_tag_key}_bff"
bff_profile="mr-${service}-bff"
[[ "$static_tag_key" =~ ^[a-z0-9_]+$ ]] || { echo "##[error] Unsafe static .env key"; exit 5; }
[[ "$bff_profile" =~ ^[a-zA-Z0-9][a-zA-Z0-9_.-]+$ ]] || { echo "##[error] Unsafe BFF Compose profile"; exit 5; }

compose_backup="$workdir/compose.before"
env_backup="$workdir/env.before"
cp "$compose_file" "$compose_backup"
env_existed=false
if [ -f "$env_file" ]; then
  cp "$env_file" "$env_backup"
  env_existed=true
fi

read_service_image() {
  local service="$1"
  awk -v service="$service" '
    $0 == "  " service ":" { inside=1; next }
    inside && /^  [^[:space:]#][^:]*:[[:space:]]*$/ { exit }
    inside && /^[[:space:]]+image:[[:space:]]*/ {
      sub(/^[[:space:]]+image:[[:space:]]*/, "")
      sub(/[[:space:]]+#.*$/, "")
      gsub(/^["'\'' ]+|["'\'' ]+$/, "")
      print
      exit
    }
  ' "$compose_file"
}

update_service_image() {
  local service="$1"
  local image="$2"
  local output="$workdir/compose.updated"
  if ! awk -v service="$service" -v image="$image" '
    $0 == "  " service ":" { inside=1; found=1; print; next }
    inside && /^  [^[:space:]#][^:]*:[[:space:]]*$/ { inside=0 }
    inside && /^[[:space:]]+image:[[:space:]]*/ && !updated {
      print "    image: " image
      updated=1
      next
    }
    { print }
    END { if (!found || !updated) exit 42 }
  ' "$compose_file" > "$output"; then
    echo "##[error] Compose service or image field was not found: $service" >&2
    return 1
  fi
  mv "$output" "$compose_file"
}

image_tag_for_repository() {
  local image="$1"
  local repository="$2"
  local tag
  case "$image" in
    "${repository}:"*) tag="${image#"${repository}:"}" ;;
    *) echo "##[error] Image does not belong to the expected repository: $image" >&2; return 1 ;;
  esac
  [[ "$tag" =~ ^[A-Za-z0-9_.-]+$ ]] || {
    echo "##[error] Image has an unsafe or unsupported tag: $image" >&2
    return 1
  }
  printf '%s' "$tag"
}

update_env_tag() {
  local key="$1"
  local value="$2"
  local output="$workdir/env.updated"
  mkdir -p "$(dirname "$env_file")"
  touch "$env_file"
  awk -v key="$key" -v value="$value" '
    BEGIN { updated=0 }
    $0 ~ "^[[:space:]]*" key "[[:space:]]*[:=]" {
      if (!updated) {
        print key ":" value
        updated=1
      }
      next
    }
    { print }
    END {
      if (!updated) print key ":" value
    }
  ' "$env_file" > "$output"
  mv "$output" "$env_file"
}

read_env_value() {
  local key="$1"
  awk -v key="$key" '
    $0 ~ "^[[:space:]]*" key "[[:space:]]*[:=]" {
      value=$0
      sub("^[[:space:]]*" key "[[:space:]]*[:=][[:space:]]*", "", value)
      sub(/[[:space:]]+$/, "", value)
      print value
      exit
    }
  ' "$env_file"
}

update_env_variable() {
  local key="$1"
  local value="$2"
  local output="$workdir/env.variable.updated"
  mkdir -p "$(dirname "$env_file")"
  touch "$env_file"
  awk -v key="$key" -v value="$value" '
    BEGIN { updated=0 }
    $0 ~ "^[[:space:]]*" key "[[:space:]]*[:=]" {
      if (!updated) {
        print key "=" value
        updated=1
      }
      next
    }
    { print }
    END {
      if (!updated) print key "=" value
    }
  ' "$env_file" > "$output"
  mv "$output" "$env_file"
}

remove_env_value() {
  local key="$1"
  local output="$workdir/env.variable.removed"
  [ -f "$env_file" ] || return 0
  awk -v key="$key" '
    $0 !~ "^[[:space:]]*" key "[[:space:]]*[:=]" { print }
  ' "$env_file" > "$output"
  mv "$output" "$env_file"
}

merge_env_csv_item() {
  local key="$1"
  local item="$2"
  local enabled="$3"
  local current merged
  current="$(read_env_value "$key")"
  merged="$(awk -v current="$current" -v item="$item" -v enabled="$enabled" '
    BEGIN {
      output=""
      count=split(current, values, ",")
      for (i=1; i<=count; i++) {
        value=values[i]
        gsub(/^[[:space:]]+|[[:space:]]+$/, "", value)
        if (value == "" || value == item || seen[value]++) continue
        output = output (output == "" ? "" : ",") value
      }
      if (enabled == "true") output = output (output == "" ? "" : ",") item
      print output
    }
  ')"
  if [ -n "$merged" ]; then
    update_env_variable "$key" "$merged"
  else
    remove_env_value "$key"
  fi
}

current_static_image="$(read_service_image "$static_container")"
[ -n "$current_static_image" ] || { echo "##[error] Static Compose image was not found"; exit 5; }
static_tag="$(image_tag_for_repository "$static_image" "$static_repository")"
current_bff_image=""
bff_tag=""
if [ -n "$bff_image" ]; then
  current_bff_image="$(read_service_image "$bff_container")"
  [ -n "$current_bff_image" ] || { echo "##[error] BFF Compose image was not found"; exit 5; }
  bff_tag="$(image_tag_for_repository "$bff_image" "$bff_repository")"
fi

printf -v static_compose_image '%s:${%s}' "$static_repository" "$static_tag_key"
update_service_image "$static_container" "$static_compose_image"
update_env_tag "$static_tag_key" "$static_tag"
if [ -n "$bff_image" ]; then
  printf -v bff_compose_image '%s:${%s}' "$bff_repository" "$bff_tag_key"
  update_service_image "$bff_container" "$bff_compose_image"
  update_env_tag "$bff_tag_key" "$bff_tag"
  merge_env_csv_item COMPOSE_PROFILES "$bff_profile" true
else
  merge_env_csv_item COMPOSE_PROFILES "$bff_profile" false
fi

compose_changed=false
git -C "$repo_dir" add -A -- "${compose_path#/}" "${env_path#/}"
if ! git -C "$repo_dir" diff --cached --quiet; then
  compose_changed=true
  git -C "$repo_dir" commit --quiet -m "release(monorepo): deploy build ${build_id}"
  echo "##[section]Publishing immutable image tags to ${env_path}"
  git_authenticated -C "$repo_dir" push --quiet origin HEAD:main
else
  echo "##[section]Compose .env already references build ${build_id}; deployment is idempotent"
fi

request_number=0
komodo_call() {
  endpoint="$1"
  payload="$2"
  request_number=$((request_number + 1))
  request_file="$workdir/komodo-request-$request_number.json"
  response_file="$workdir/komodo-response-$request_number.json"
  printf '%s' "$payload" > "$request_file"
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
  printf '%s' "$response_file"
}

deploy_stack() {
  attempt="$1"
  response="$(komodo_call execute "$(jq -cn --arg stack "$komodo_stack" '{type:"DeployStack",params:{stack:$stack,services:[]}}')")" || return 1
  update_id="$(jq -r '
    def to_id:
      if type == "object" then .["$oid"] // empty
      elif type == "string" then .
      else empty
      end;
    (.update // .response // .data // .) |
    ((._id | to_id) // (.id | to_id) // "")
  ' "$response")"
  update_status="$(jq -r '(.update // .response // .data // .).status // ""' "$response")"
  polls=0
  while [ "$update_status" != Complete ]; do
    [ -n "$update_id" ] || { echo "##[error] DeployStack returned no Update ID" >&2; return 1; }
    polls=$((polls + 1))
    [ "$polls" -le 300 ] || { echo "##[error] Timed out waiting for Komodo Stack deployment" >&2; return 1; }
    sleep 2
    response="$(komodo_call read "$(jq -cn --arg id "$update_id" '{type:"GetUpdate",params:{id:$id}}')")" || return 1
    update_status="$(jq -r '(.update // .response // .data // .).status // ""' "$response")"
  done
  if ! jq -e '(.update // .response // .data // .).success == true' "$response" >/dev/null; then
    echo "##[error] Komodo DeployStack failed during $attempt for $komodo_stack" >&2
    jq -r '
      def nonempty($label; $value):
        if (($value // "") | tostring | length) > 0
        then $label + ":\n" + ($value | tostring)
        else empty
        end;
      (.update // .response // .data // .) as $update
      | (if (($update.other_data // "") | tostring | length) > 0
         then "Komodo other_data:\n" + ($update.other_data | tostring)
         else empty
         end),
        ($update.logs[-15:][]?
          | if type == "object"
            then "---- " + (.stage // "Komodo" | tostring),
              nonempty("message"; .message),
              nonempty("command"; .command),
              nonempty("stdout"; .stdout),
              nonempty("stderr"; .stderr)
            else tostring
            end)
    ' "$response" >&2 || true
    return 1
  fi
  echo "##[section]Komodo Stack deployment completed: $komodo_stack ($attempt)"
}

if deploy_stack "build ${build_id}"; then
  echo "##[section]Immutable MR release completed: static=$static_image bff=${bff_image:-disabled}"
  exit 0
fi

echo "##[warning]Deployment failed; restoring the previously active Compose and .env state"
if [ "$compose_changed" = true ]; then
  cp "$compose_backup" "$compose_file"
  if [ "$env_existed" = true ]; then
    cp "$env_backup" "$env_file"
  else
    rm -f "$env_file"
  fi
  git -C "$repo_dir" add -A -- "${compose_path#/}" "${env_path#/}"
  git -C "$repo_dir" commit --quiet -m "rollback(monorepo): restore before build ${build_id}"
  if git_authenticated -C "$repo_dir" push --quiet origin HEAD:main; then
    if ! deploy_stack "rollback before build ${build_id}"; then
      echo "##[error]Rollback tags were pushed, but Komodo failed to redeploy them" >&2
    fi
  else
    echo "##[error]Could not push the rollback Compose commit" >&2
  fi
else
  echo "##[warning]No Compose change was published, so no Git rollback was required"
fi
echo "##[error]MR release failed and rollback was attempted" >&2
exit 1
