#!/usr/bin/env bash
set -Eeuo pipefail

root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
release_script="$root/dist/monorepo-release-inline-task.sh"
test_root="$(mktemp -d /tmp/mr-release-env-test.XXXXXX)"
cleanup() { rm -rf "$test_root"; }
trap cleanup EXIT

fail() {
  echo "[monorepo release env validation] $*" >&2
  exit 1
}

remote="$test_root/collection/project/_git/Locanit_Docker_DevOps"
seed="$test_root/seed"
mkdir -p "$(dirname "$remote")" "$seed/demo_locanit" "$test_root/artifacts/mr-drop" "$test_root/mock-bin"
git init --quiet --bare "$remote"
git -C "$seed" init --quiet
git -C "$seed" config user.name test
git -C "$seed" config user.email test@local

printf '%s\n' \
  'services:' \
  '  locanit_front_monorepo_demo:' \
  '    image: registry.buluttakin.com/locanit/front-monorepo-demo:1.0.15053' \
  '    restart: unless-stopped' \
  '  locanit_front_monorepo_bff_demo:' \
  '    image: registry.buluttakin.com/locanit/front-monorepo-bff-demo:1.0.15040' \
  '    restart: unless-stopped' \
  '  unrelated:' \
  '    image: registry.example/unrelated:7' \
  > "$seed/demo_locanit/compose.yml"
printf '%s\n' \
  'locanit_back:1.0.14925' \
  'COMPOSE_PROFILES=operations' \
  > "$seed/demo_locanit/.env"
git -C "$seed" add -- demo_locanit/compose.yml demo_locanit/.env
git -C "$seed" commit --quiet -m initial
git -C "$seed" branch -M main
git -C "$seed" remote add origin "$remote"
git -C "$seed" push --quiet -u origin main

write_manifest() {
  local build_id="$1"
  local static_tag="$2"
  local bff_tag="$3"
  jq -n \
    --arg buildId "$build_id" \
    --arg collectionUri "file://$test_root/collection" \
    --arg staticTag "$static_tag" \
    --arg bffTag "$bff_tag" '
      {
        schemaVersion: 2,
        deploymentMode: "immutable-images",
        buildId: $buildId,
        collectionUri: $collectionUri,
        azureProjectId: "project",
        komodoStack: "locanit_docker_devops-demo",
        composeRepository: "Locanit_Docker_DevOps",
        composePath: "/demo_locanit/compose.yml",
        service: "front-monorepo",
        staticContainer: "locanit_front_monorepo_demo",
        bffContainer: "locanit_front_monorepo_bff_demo",
        staticImageRepository: "registry.buluttakin.com/locanit/front-monorepo-demo",
        bffImageRepository: "registry.buluttakin.com/locanit/front-monorepo-bff-demo",
        staticImage: ("registry.buluttakin.com/locanit/front-monorepo-demo:" + $staticTag),
        bffImage: ("registry.buluttakin.com/locanit/front-monorepo-bff-demo:" + $bffTag)
      }
    ' > "$test_root/artifacts/mr-drop/manifest.json"
}

printf '%s\n' \
  '#!/usr/bin/env bash' \
  'set -Eeuo pipefail' \
  'output=""' \
  'while [ "$#" -gt 0 ]; do' \
  '  case "$1" in' \
  '    -o) output="$2"; shift 2 ;;' \
  '    *) shift ;;' \
  '  esac' \
  'done' \
  '[ -n "$output" ]' \
  'count=0' \
  'if [ -n "${MOCK_CURL_COUNT_FILE:-}" ] && [ -f "$MOCK_CURL_COUNT_FILE" ]; then count="$(cat "$MOCK_CURL_COUNT_FILE")"; fi' \
  'count=$((count + 1))' \
  'if [ -n "${MOCK_CURL_COUNT_FILE:-}" ]; then printf "%s" "$count" > "$MOCK_CURL_COUNT_FILE"; fi' \
  'success=true' \
  'if [ "${MOCK_KOMODO_FAIL_FIRST:-0}" = 1 ] && [ "$count" -eq 1 ]; then success=false; fi' \
  'printf "{\"update\":{\"id\":\"mock-update\",\"status\":\"Complete\",\"success\":%s,\"logs\":[]}}" "$success" > "$output"' \
  'printf 200' \
  > "$test_root/mock-bin/curl"
chmod +x "$test_root/mock-bin/curl"

run_release() {
  env \
    PATH="$test_root/mock-bin:$PATH" \
    SYSTEM_DEFAULTWORKINGDIRECTORY="$test_root/artifacts" \
    AZP_TOKEN=test-token \
    KOMODO_API_KEY=test-key \
    KOMODO_API_SECRET=test-secret \
    KOMODO_ADDRESS=https://komodo.test \
    MOCK_CURL_COUNT_FILE="$test_root/curl-count" \
    "$@" \
    bash "$release_script"
}

write_manifest 15060 1.0.15060 1.0.15040
printf 0 > "$test_root/curl-count"
run_release

git --git-dir="$remote" show main:demo_locanit/compose.yml > "$test_root/compose.after"
git --git-dir="$remote" show main:demo_locanit/.env > "$test_root/env.after"
grep -Fq 'image: registry.buluttakin.com/locanit/front-monorepo-demo:${front_monorepo}' "$test_root/compose.after"
grep -Fq 'image: registry.buluttakin.com/locanit/front-monorepo-bff-demo:${front_monorepo_bff}' "$test_root/compose.after"
grep -Fq 'image: registry.example/unrelated:7' "$test_root/compose.after"
grep -Fxq 'locanit_back:1.0.14925' "$test_root/env.after"
grep -Fxq 'front_monorepo:1.0.15060' "$test_root/env.after"
grep -Fxq 'front_monorepo_bff:1.0.15040' "$test_root/env.after"
grep -Fxq 'COMPOSE_PROFILES=operations,mr-front-monorepo-bff' "$test_root/env.after"

commit_count="$(git --git-dir="$remote" rev-list --count main)"
printf 0 > "$test_root/curl-count"
run_release
[ "$(git --git-dir="$remote" rev-list --count main)" = "$commit_count" ] || fail 'idempotent release created a Git commit'

cp "$test_root/compose.after" "$test_root/compose.before-rollback"
cp "$test_root/env.after" "$test_root/env.before-rollback"
write_manifest 15061 1.0.15061 1.0.15061
printf 0 > "$test_root/curl-count"
if run_release MOCK_KOMODO_FAIL_FIRST=1; then
  fail 'controlled failed deployment unexpectedly succeeded'
fi
git --git-dir="$remote" show main:demo_locanit/compose.yml > "$test_root/compose.rolled-back"
git --git-dir="$remote" show main:demo_locanit/.env > "$test_root/env.rolled-back"
cmp -s "$test_root/compose.before-rollback" "$test_root/compose.rolled-back" || fail 'Compose was not restored exactly'
cmp -s "$test_root/env.before-rollback" "$test_root/env.rolled-back" || fail '.env was not restored exactly'
[ "$(cat "$test_root/curl-count")" = 2 ] || fail 'rollback did not redeploy the previous state'

echo 'Monorepo Release .env validation passed: migration, idempotency, and exact rollback are correct.'
