#!/usr/bin/env bash
set -Eeuo pipefail
trap 'code=$?; echo "##[error] MR image packaging failed at line $LINENO (exit=$code)"; exit "$code"' ERR

require() { [ -n "${!1:-}" ] || { echo "##[error] $1 is required"; exit 2; }; }
for variable in \
  MR_ARTIFACT_DIR \
  MR_TEMPLATES_DIR \
  MR_STATIC_IMAGE_REPOSITORY \
  MR_BFF_IMAGE_REPOSITORY \
  MR_IMAGE_TAG \
  MR_STATIC_RUNTIME_IMAGE \
  MR_BFF_RUNTIME_IMAGE
do
  require "$variable"
done
for command_name in docker jq awk; do
  command -v "$command_name" >/dev/null 2>&1 || {
    echo "##[error] $command_name is required on the Build agent"; exit 2;
  }
done

artifact_dir="$(cd "$MR_ARTIFACT_DIR" && pwd)"
templates_dir="$(cd "$MR_TEMPLATES_DIR" && pwd)"
manifest="$artifact_dir/manifest.json"
inventory="$artifact_dir/inventory.tsv"
modules="$artifact_dir/modules.tsv"
nginx_config="$artifact_dir/runtime/nginx/default.conf"
test -s "$manifest"
test -f "$inventory"
test -f "$modules"
test -s "$nginx_config"
test -s "$templates_dir/monorepo/docker/static.Dockerfile"
test -s "$templates_dir/monorepo/docker/bff.Dockerfile"

case "$MR_STATIC_IMAGE_REPOSITORY" in (*[!a-z0-9./:_-]*|'') echo "##[error] Unsafe static image repository"; exit 3;; esac
case "$MR_BFF_IMAGE_REPOSITORY" in (*[!a-z0-9./:_-]*|'') echo "##[error] Unsafe BFF image repository"; exit 3;; esac
case "$MR_IMAGE_TAG" in (*[!A-Za-z0-9_.-]*|'') echo "##[error] Unsafe image tag"; exit 3;; esac

static_image="${MR_STATIC_IMAGE_REPOSITORY}:${MR_IMAGE_TAG}"
bff_image=""
previous_static_image="${MR_PREVIOUS_STATIC_IMAGE:-}"
previous_bff_image="${MR_PREVIOUS_BFF_IMAGE:-}"
workdir="$(mktemp -d "${AGENT_TEMPDIRECTORY:-/tmp}/mr-images.XXXXXX")"
container_id=""
cleanup() {
  code=$?
  trap - EXIT
  if [ -n "$container_id" ]; then docker rm -f "$container_id" >/dev/null 2>&1 || true; fi
  rm -rf "$workdir" >/dev/null 2>&1 || true
  exit "$code"
}
trap cleanup EXIT

static_context="$workdir/static"
static_current="$static_context/current"
mkdir -p "$static_current/root" "$static_current/modules" "$static_context/nginx"

if [ -n "$previous_static_image" ]; then
  echo "##[section]Hydrating the previous static runtime: $previous_static_image"
  docker pull "$previous_static_image"
  container_id="$(docker create "$previous_static_image")"
  if docker cp "$container_id:/srv/monorepo/current/." "$static_current/" >/dev/null 2>&1; then
    :
  else
    echo "##[error] Previous static image has no managed /srv/monorepo/current tree: $previous_static_image"
    exit 4
  fi
  docker rm "$container_id" >/dev/null
  container_id=""
fi

while IFS=$'\t' read -r name kind route; do
  [ -n "$name" ] || continue
  case "$name" in (*[!A-Za-z0-9._-]*|'') echo "##[error] Unsafe module name: $name"; exit 5;; esac
  source="$artifact_dir/modules/$name"
  test -d "$source" || { echo "##[error] Packaged module is missing: $name"; exit 5; }
  case "$kind" in
    shell)
      rm -rf "$static_current/root"
      mkdir -p "$static_current/root"
      cp -a "$source/." "$static_current/root/"
      ;;
    static)
      rm -rf "$static_current/modules/$name"
      mkdir -p "$static_current/modules/$name"
      cp -a "$source/." "$static_current/modules/$name/"
      ;;
    bff) ;;
    *) echo "##[error] Unsupported module kind '$kind' for $name"; exit 5;;
  esac
done < "$modules"

shell_count=0
while IFS=$'\t' read -r name kind route; do
  [ -n "$name" ] || continue
  case "$kind" in
    shell)
      shell_count=$((shell_count + 1))
      test -n "$(find "$static_current/root" -mindepth 1 -print -quit 2>/dev/null)" || {
        echo "##[error] No successful or previous shell output exists for $name"; exit 6;
      }
      ;;
    static)
      test -n "$(find "$static_current/modules/$name" -mindepth 1 -print -quit 2>/dev/null)" || {
        echo "##[error] No successful or previous static output exists for $name"; exit 6;
      }
      rm -rf "$static_current/root/$name"
      ln -s "../modules/$name" "$static_current/root/$name"
      ;;
  esac
done < "$inventory"
[ "$shell_count" -le 1 ] || { echo "##[error] Multiple shell projects are not supported"; exit 6; }

cp "$inventory" "$static_current/inventory.tsv"
cp "$manifest" "$static_current/manifest.json"
cp "$nginx_config" "$static_context/nginx/default.conf"

echo "##[section]Building immutable static runtime: $static_image"
docker build \
  --build-arg "RUNTIME_IMAGE=$MR_STATIC_RUNTIME_IMAGE" \
  --file "$templates_dir/monorepo/docker/static.Dockerfile" \
  --tag "$static_image" \
  "$static_context"
docker push "$static_image"

bff_project="$(jq -r '.bffProject // empty' "$manifest")"
bff_entry="$(jq -r '.bffEntry // "main.js"' "$manifest")"
if [ -n "$bff_project" ]; then
  case "$bff_project" in (*[!A-Za-z0-9._-]*|'') echo "##[error] Unsafe BFF project: $bff_project"; exit 7;; esac
  case "$bff_entry" in (*[!A-Za-z0-9._/-]*|'') echo "##[error] Unsafe BFF entry: $bff_entry"; exit 7;; esac
  if awk -F '\t' -v name="$bff_project" '$1 == name && $2 == "bff" { found=1 } END { exit !found }' "$modules"; then
    bff_context="$workdir/bff"
    mkdir -p "$bff_context/app"
    cp -a "$artifact_dir/modules/$bff_project/." "$bff_context/app/"
    bff_image="${MR_BFF_IMAGE_REPOSITORY}:${MR_IMAGE_TAG}"
    echo "##[section]Building immutable BFF runtime: $bff_image"
    docker build \
      --build-arg "RUNTIME_IMAGE=$MR_BFF_RUNTIME_IMAGE" \
      --build-arg "BFF_ENTRY=$bff_entry" \
      --file "$templates_dir/monorepo/docker/bff.Dockerfile" \
      --tag "$bff_image" \
      "$bff_context"
    docker push "$bff_image"
  elif [ -n "$previous_bff_image" ]; then
    echo "##[section]BFF was unaffected or failed; retaining $previous_bff_image"
    docker pull "$previous_bff_image"
    bff_image="$previous_bff_image"
  else
    echo "##[error] No successful or previous BFF image exists for $bff_project"
    exit 7
  fi
fi

updated_manifest="$workdir/manifest.json"
jq \
  --arg staticImage "$static_image" \
  --arg previousStaticImage "$previous_static_image" \
  --arg bffImage "$bff_image" \
  --arg previousBffImage "$previous_bff_image" \
  --arg staticImageRepository "$MR_STATIC_IMAGE_REPOSITORY" \
  --arg bffImageRepository "$MR_BFF_IMAGE_REPOSITORY" \
  --arg imageTag "$MR_IMAGE_TAG" '
    del(.deploymentRoot) |
    .schemaVersion = 2 |
    .deploymentMode = "immutable-images" |
    .staticImage = $staticImage |
    .previousStaticImage = $previousStaticImage |
    .bffImage = $bffImage |
    .previousBffImage = $previousBffImage |
    .staticImageRepository = $staticImageRepository |
    .bffImageRepository = $bffImageRepository |
    .imageTag = $imageTag
  ' "$manifest" > "$updated_manifest"
mv "$updated_manifest" "$manifest"
echo "##[section]Immutable MR images are ready: static=$static_image bff=${bff_image:-disabled}"
