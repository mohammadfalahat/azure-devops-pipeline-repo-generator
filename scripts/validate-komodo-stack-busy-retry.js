#!/usr/bin/env node

const fs = require('fs');
const os = require('os');
const path = require('path');
const { spawnSync } = require('child_process');
const yaml = require('js-yaml');

const root = path.resolve(__dirname, '..');
const template = yaml.load(
  fs.readFileSync(path.join(root, 'examples/shared-templates/monorepo/pipeline.yml'), 'utf8')
);
const ensureStep = template?.stages?.[0]?.jobs?.[0]?.steps?.find(
  (step) => step.displayName === 'Ensure Komodo GitOps repository and shared Docker Stack'
);
if (!ensureStep?.inputs?.script) {
  throw new Error('[Komodo busy retry validation] Ensure Stack script was not found');
}

const testRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'komodo-busy-retry-test-'));
try {
  const mockBin = path.join(testRoot, 'bin');
  const stepPath = path.join(testRoot, 'ensure-stack.sh');
  const manifestPath = path.join(testRoot, 'manifest.json');
  const countPath = path.join(testRoot, 'update-stack-count');
  fs.mkdirSync(mockBin);
  fs.writeFileSync(stepPath, ensureStep.inputs.script, { mode: 0o700 });
  fs.writeFileSync(manifestPath, '{}\n');

  const mockCurl = `#!/usr/bin/env bash
set -Eeuo pipefail
output=""
request_file=""
while [ "$#" -gt 0 ]; do
  case "$1" in
    -o) output="$2"; shift 2 ;;
    --data-binary) request_file="\${2#@}"; shift 2 ;;
    --config|-w) shift 2 ;;
    -sS) shift ;;
    *) shift ;;
  esac
done
[ -n "$output" ] && [ -n "$request_file" ]
request_type="$(jq -r '.type // ""' "$request_file")"
case "$request_type" in
  GetServer) printf '%s' '{"_id":{"$oid":"server-id"}}' > "$output" ;;
  GetRepo) printf '%s' '{"_id":{"$oid":"repo-id"}}' > "$output" ;;
  UpdateRepo) printf '%s' '{}' > "$output" ;;
  GetStack) printf '%s' '{"_id":{"$oid":"stack-id"},"config":{"environment":"","extra_args":[]}}' > "$output" ;;
  UpdateStack)
    count=0
    [ ! -f "$MOCK_UPDATE_STACK_COUNT" ] || count="$(cat "$MOCK_UPDATE_STACK_COUNT")"
    count=$((count + 1))
    printf '%s' "$count" > "$MOCK_UPDATE_STACK_COUNT"
    if [ "\${MOCK_NON_BUSY_ERROR:-0}" = 1 ]; then
      printf '%s' '{"error":"permission denied"}' > "$output"
      printf 403
      exit 0
    fi
    if [ "$count" -le "\${MOCK_BUSY_CALLS:-0}" ]; then
      printf '%s' '{"error":"Stack busy"}' > "$output"
      printf 500
      exit 0
    fi
    printf '%s' '{}' > "$output"
    ;;
  *) printf '%s' '{}' > "$output" ;;
esac
printf 200
`;
  const curlPath = path.join(mockBin, 'curl');
  fs.writeFileSync(curlPath, mockCurl, { mode: 0o700 });

  const baseEnv = {
    ...process.env,
    PATH: `${mockBin}:${process.env.PATH}`,
    AGENT_TEMPDIRECTORY: testRoot,
    KOMODO_ADDRESS: 'https://komodo.test',
    KOMODO_API_KEY: 'test-key',
    KOMODO_API_SECRET: 'test-secret',
    KOMODO_STACK_BUSY_MAX_ATTEMPTS: '5',
    KOMODO_STACK_BUSY_RETRY_SECONDS: '0',
    MR_MANIFEST: manifestPath,
    MR_KOMODO_SERVER: 'demo-server',
    MR_KOMODO_REPOSITORY: 'locanit_docker_devops-demo',
    MR_KOMODO_STACK: 'locanit_docker_devops-demo',
    MR_COMPOSE_REPOSITORY: 'Locanit_Docker_DevOps',
    MR_COMPOSE_PATH: '/demo_locanit/compose.yml',
    MR_PROJECT_KEY: 'locanit',
    MR_SERVICE_KEY: 'front-monorepo',
    MR_ENVIRONMENT: 'demo',
    MR_RUNTIME_VARIABLE_PREFIX: 'MR_LOCANIT_FRONT_MONOREPO_DEMO',
    MR_BFF_PROFILE: 'mr-front-monorepo-bff',
    MR_GIT_ACCOUNT: '23',
    SYSTEM_COLLECTIONURI: 'https://azure.example/ShonizCollection/',
    SYSTEM_TEAMPROJECT: 'Locanit',
    MOCK_UPDATE_STACK_COUNT: countPath
  };

  const busyResult = spawnSync('bash', [stepPath], {
    env: { ...baseEnv, MOCK_BUSY_CALLS: '2' },
    encoding: 'utf8'
  });
  if (busyResult.status !== 0) {
    throw new Error(`[Komodo busy retry validation] retry run failed:\n${busyResult.stderr}`);
  }
  if (fs.readFileSync(countPath, 'utf8') !== '3') {
    throw new Error('[Komodo busy retry validation] UpdateStack did not succeed after two retries');
  }
  const warnings = `${busyResult.stdout}\n${busyResult.stderr}`.match(/retrying in 0s/g) || [];
  if (warnings.length !== 2) {
    throw new Error('[Komodo busy retry validation] expected one warning for each busy retry');
  }

  fs.writeFileSync(countPath, '0');
  const permanentResult = spawnSync('bash', [stepPath], {
    env: { ...baseEnv, MOCK_NON_BUSY_ERROR: '1' },
    encoding: 'utf8'
  });
  if (permanentResult.status === 0 || fs.readFileSync(countPath, 'utf8') !== '1') {
    throw new Error('[Komodo busy retry validation] a non-busy failure was retried or accepted');
  }

  fs.writeFileSync(countPath, '0');
  const exhaustedResult = spawnSync('bash', [stepPath], {
    env: {
      ...baseEnv,
      KOMODO_STACK_BUSY_MAX_ATTEMPTS: '3',
      MOCK_BUSY_CALLS: '99'
    },
    encoding: 'utf8'
  });
  if (exhaustedResult.status === 0 || fs.readFileSync(countPath, 'utf8') !== '3') {
    throw new Error('[Komodo busy retry validation] busy retries did not stop at the configured bound');
  }

  console.log('Komodo Stack busy retry validation passed: transient busy responses retry, bounds are enforced, and permanent errors fail immediately.');
} finally {
  fs.rmSync(testRoot, { recursive: true, force: true });
}
