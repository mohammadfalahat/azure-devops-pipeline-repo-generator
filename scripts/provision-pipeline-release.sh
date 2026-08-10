#!/usr/bin/env bash

set -euo pipefail

# Provision an Azure DevOps YAML pipeline after its YAML file has already been
# pushed, then provision a classic Release definition that uses that pipeline as
# its Build artifact and runs one inline Bash task.

CURRENT_STEP="startup"
HANDLED_ERROR=0

on_unhandled_error() {
  local exit_code="$1"
  local line_no="$2"
  local command_text="$3"
  if [[ "$HANDLED_ERROR" == "1" ]]; then
    exit "$exit_code"
  fi

  echo "[ERROR] Step failed unexpectedly: $CURRENT_STEP" >&2
  echo "[ERROR] Exit code: $exit_code" >&2
  echo "[ERROR] Line: $line_no" >&2
  echo "[ERROR] Command: $command_text" >&2
  exit "$exit_code"
}

trap 'on_unhandled_error $? $LINENO "$BASH_COMMAND"' ERR

step() {
  CURRENT_STEP="$1"
  echo "[STEP] $CURRENT_STEP"
}

fail() {
  local message="$1"
  local details="${2:-}"
  echo "[ERROR] Step: $CURRENT_STEP" >&2
  echo "[ERROR] $message" >&2
  if [[ -n "$details" ]]; then
    echo "[ERROR] Details:" >&2
    echo "$details" >&2
  fi
  HANDLED_ERROR=1
  exit 1
}

usage() {
  cat <<'USAGE'
Usage:
  AZP_TOKEN=... ADO_URL=... COLLECTION=... PROJECT=... \
  PIPELINE_NAME=... REPO_NAME=... YAML_PATH=... \
  RELEASE_AGENT_QUEUE_NAME=... RELEASE_BASH_SCRIPT_FILE=./release-task.sh \
  scripts/provision-pipeline-release.sh

Required environment variables:
  AZP_TOKEN, ADO_URL, COLLECTION, PROJECT, PIPELINE_NAME, REPO_NAME, YAML_PATH
  RELEASE_AGENT_QUEUE_ID or RELEASE_AGENT_QUEUE_NAME
  RELEASE_BASH_SCRIPT or RELEASE_BASH_SCRIPT_FILE or
  RELEASE_BASH_SCRIPT_GIT_URL + RELEASE_BASH_SCRIPT_GIT_PATH

Optional environment variables:
  DEFAULT_BRANCH=refs/heads/main
  PIPELINE_FOLDER=komodo
  RELEASE_NAME=${PIPELINE_NAME}_Release
  RELEASE_FOLDER=komodo
  RELEASE_ENVIRONMENT_NAME=komodo
  RELEASE_VARIABLE_GROUP_NAME=KomodoAPI
  RELEASE_ARTIFACT_ALIAS=_${PIPELINE_NAME}
  RELEASE_BASH_TASK_NAME='Run inline Bash'
  RELEASE_BASH_SCRIPT_GIT_REF=main
  CREATE_RELEASE_INSTANCE=false
  API_VERSION=7.1
  RELEASE_API_VERSION=7.1-preview.4

Use --self-test to validate JSON generation without calling Azure DevOps.
USAGE
}

if [[ "${1:-}" == "--help" || "${1:-}" == "-h" ]]; then
  usage
  exit 0
fi

require_command() {
  if ! command -v "$1" >/dev/null 2>&1; then
    fail "Required command not found: $1" "Install '$1' and retry."
  fi
}

require_env() {
  if [[ -z "${!1:-}" ]]; then
    fail "Required environment variable is missing: $1" "Run with '$1=...' exported or inline before the command."
  fi
}

curl_with_basic_auth() {
  # Feed curl's Basic-auth setting over stdin instead of expanding the PAT in
  # curl's command-line arguments. This keeps the secret out of shell history
  # and process listings while avoiding a credential file on disk.
  local escaped_token="$AZP_TOKEN"
  if [[ "$escaped_token" == *$'\n'* || "$escaped_token" == *$'\r'* ]]; then
    fail "AZP_TOKEN contains a newline." "Use the PAT exactly as issued by Azure DevOps."
  fi
  escaped_token="${escaped_token//\\/\\\\}"
  escaped_token="${escaped_token//\"/\\\"}"
  printf 'user = ":%s"\n' "$escaped_token" | curl --config - "$@"
}

collection_api_url() {
  local area_path="$1"
  local api_version="$2"
  printf '%s/%s/%s?api-version=%s' "${ADO_URL%/}" "$COLLECTION" "$area_path" "$api_version"
}

project_api_url() {
  local area_path="$1"
  local api_version="$2"
  printf '%s/%s/%s/%s?api-version=%s' "${ADO_URL%/}" "$COLLECTION" "$PROJECT" "$area_path" "$api_version"
}

pipeline_collection_url() {
  local api_version="$1"
  printf '%s&repositoryId=%s' "$(project_api_url '_apis/pipelines' "$api_version")" "$REPO_ID"
}

api_request() {
  local method="$1"
  local url="$2"
  local payload_file="${3:-}"
  local response_file http_code curl_exit response_body
  response_file="$(mktemp)"
  curl_exit=0

  if [[ "$method" == "GET" ]]; then
    http_code="$(curl_with_basic_auth -sS \
      -H 'Accept: application/json' \
      -o "$response_file" \
      -w '%{http_code}' \
      "$url")" || curl_exit=$?
  else
    [[ -n "$payload_file" && -f "$payload_file" ]] || fail "Payload file is missing for $method request." "URL: $url\nPayload file: ${payload_file:-not-set}"
    http_code="$(curl_with_basic_auth -sS \
      -X "$method" \
      -H 'Content-Type: application/json' \
      -H 'Accept: application/json' \
      --data-binary "@$payload_file" \
      -o "$response_file" \
      -w '%{http_code}' \
      "$url")" || curl_exit=$?
  fi

  response_body="$(cat "$response_file")"
  rm -f "$response_file"

  if [[ "$curl_exit" -ne 0 ]]; then
    fail "HTTP request could not be completed." "Method: $method\nURL: $url\nCurl exit code: $curl_exit\nResponse body:\n$response_body"
  fi

  if ! [[ "$http_code" =~ ^[0-9]{3}$ ]]; then
    fail "Azure DevOps returned an invalid HTTP status code." "Method: $method\nURL: $url\nHTTP status: ${http_code:-empty}\nResponse body:\n$response_body"
  fi

  if (( http_code < 200 || http_code >= 300 )); then
    fail "Azure DevOps REST API request failed." "Method: $method\nURL: $url\nHTTP status: $http_code\nResponse body:\n$response_body"
  fi

  printf '%s' "$response_body"
}

api_get() {
  api_request GET "$1"
}

api_post() {
  api_request POST "$1" "$2"
}

normalize_folder() {
  local folder="$1"
  if [[ -z "$folder" || "$folder" == "\\" ]]; then
    printf '\\'
    return
  fi
  folder="${folder//\//\\}"
  folder="${folder#\\}"
  printf '\\%s' "$folder"
}

read_inline_script() {
  if [[ -n "${RELEASE_BASH_SCRIPT:-}" ]]; then
    printf '%s' "$RELEASE_BASH_SCRIPT"
    return
  fi

  if [[ -n "${RELEASE_BASH_SCRIPT_FILE:-}" ]]; then
    [[ -f "$RELEASE_BASH_SCRIPT_FILE" ]] || {
      fail "Release Bash script file was not found." "RELEASE_BASH_SCRIPT_FILE=$RELEASE_BASH_SCRIPT_FILE"
    }
    cat "$RELEASE_BASH_SCRIPT_FILE"
    return
  fi

  if [[ -n "${RELEASE_BASH_SCRIPT_GIT_URL:-}" || -n "${RELEASE_BASH_SCRIPT_GIT_PATH:-}" ]]; then
    require_env RELEASE_BASH_SCRIPT_GIT_URL
    require_env RELEASE_BASH_SCRIPT_GIT_PATH
    require_command git

    local temp_dir script_path
    temp_dir="$(mktemp -d)"
    trap 'rm -rf "$temp_dir"' RETURN

    echo "[INFO] Cloning release Bash script repository..." >&2
    git clone --depth 1 --branch "${RELEASE_BASH_SCRIPT_GIT_REF:-main}" \
      "$RELEASE_BASH_SCRIPT_GIT_URL" "$temp_dir/script-repo" >&2 || \
      fail "Could not clone release Bash script repository." "Repository: $RELEASE_BASH_SCRIPT_GIT_URL\nRef: ${RELEASE_BASH_SCRIPT_GIT_REF:-main}"

    script_path="$temp_dir/script-repo/$RELEASE_BASH_SCRIPT_GIT_PATH"
    [[ -f "$script_path" ]] || {
      fail "Script path not found in cloned repository." "Expected path: $RELEASE_BASH_SCRIPT_GIT_PATH\nRepository: $RELEASE_BASH_SCRIPT_GIT_URL"
    }
    cat "$script_path"
    return
  fi

  fail "No release Bash inline script source was provided." "Provide one of:\n- RELEASE_BASH_SCRIPT\n- RELEASE_BASH_SCRIPT_FILE\n- RELEASE_BASH_SCRIPT_GIT_URL + RELEASE_BASH_SCRIPT_GIT_PATH"
}

build_pipeline_payload() {
  jq -n \
    --arg name "$PIPELINE_NAME" \
    --arg folder "$PIPELINE_FOLDER" \
    --arg yamlPath "$YAML_PATH" \
    --arg repoId "$REPO_ID" \
    --arg repoName "$REPO_NAME" \
    --arg defaultBranch "$DEFAULT_BRANCH" \
    '{
      name: $name,
      folder: $folder,
      configuration: {
        type: "yaml",
        path: $yamlPath,
        repository: {
          id: $repoId,
          name: $repoName,
          type: "azureReposGit",
          defaultBranch: $defaultBranch
        }
      }
    }'
}

build_release_definition_payload() {
  local inline_script="$1"
  jq -n \
    --arg name "$RELEASE_NAME" \
    --arg path "$RELEASE_FOLDER" \
    --arg environmentName "$RELEASE_ENVIRONMENT_NAME" \
    --arg artifactAlias "$RELEASE_ARTIFACT_ALIAS" \
    --arg projectId "$PROJECT_ID" \
    --arg projectName "$PROJECT" \
    --arg definitionId "$PIPELINE_ID" \
    --arg definitionName "$PIPELINE_NAME" \
    --arg repoId "$REPO_ID" \
    --arg repoName "$REPO_NAME" \
    --arg defaultBranch "$DEFAULT_BRANCH" \
    --arg queueId "$RELEASE_AGENT_QUEUE_ID" \
    --arg queueName "$RELEASE_AGENT_QUEUE_NAME" \
    --arg variableGroupId "$RELEASE_VARIABLE_GROUP_ID" \
    --arg taskName "$RELEASE_BASH_TASK_NAME" \
    --arg inlineScript "$inline_script" \
    '{
      name: $name,
      path: $path,
      description: "Generated by azure-devops-pipeline-repo-generator",
      releaseNameFormat: "Release-$(rev:r)",
      artifacts: [{
        sourceId: ($projectId + ":" + $definitionId),
        type: "Build",
        alias: $artifactAlias,
        definitionReference: {
          definition: { id: $definitionId, name: $definitionName },
          project: { id: $projectId, name: $projectName },
          repository: { id: $repoId, name: $repoName },
          defaultVersionBranch: { id: $defaultBranch, name: $defaultBranch },
          defaultVersionType: { id: "latestType", name: "Latest" },
          defaultVersionSpecific: { id: "", name: "" },
          defaultVersionTags: { id: "", name: "" },
          artifactSourceDefinitionUrl: { id: "", name: "" }
        },
        isPrimary: true,
        isRetained: false
      }],
      environments: [{
        name: $environmentName,
        rank: 1,
        variables: {},
        variableGroups: [],
        demands: [],
        conditions: [{ name: "ReleaseStarted", conditionType: "event", value: "", result: null }],
        executionPolicy: { concurrencyCount: 0, queueDepthCount: 0 },
        schedules: [],
        retentionPolicy: { daysToKeep: 30, releasesToKeep: 3, retainBuild: true },
        processParameters: {},
        preDeployApprovals: {
          approvals: [{ rank: 1, isAutomated: true, isNotificationOn: false }],
          approvalOptions: {
            requiredApproverCount: null,
            releaseCreatorCanBeApprover: false,
            autoTriggeredAndPreviousEnvironmentApprovedCanBeSkipped: false,
            enforceIdentityRevalidation: false,
            timeoutInMinutes: 0,
            executionOrder: "beforeGates"
          }
        },
        postDeployApprovals: {
          approvals: [{ rank: 1, isAutomated: true, isNotificationOn: false }],
          approvalOptions: {
            requiredApproverCount: null,
            releaseCreatorCanBeApprover: false,
            autoTriggeredAndPreviousEnvironmentApprovedCanBeSkipped: false,
            enforceIdentityRevalidation: false,
            timeoutInMinutes: 0,
            executionOrder: "afterSuccessfulGates"
          }
        },
        deployPhases: [{
          name: "Agent job",
          phaseType: "agentBasedDeployment",
          rank: 1,
          workflowTasks: [{
            taskId: "6c731c3c-3c68-459a-a5c9-bde6e6595b5b",
            version: "3.*",
            name: $taskName,
            refName: "",
            enabled: true,
            alwaysRun: false,
            continueOnError: false,
            timeoutInMinutes: 0,
            definitionType: "task",
            condition: "succeeded()",
            inputs: {
              targetType: "inline",
              script: $inlineScript,
              workingDirectory: "",
              failOnStderr: "false",
              noProfile: "true",
              noRc: "true"
            }
          }],
          deploymentInput: {
            queueId: ($queueId | tonumber),
            queueName: $queueName,
            demands: [],
            enableAccessToken: false,
            skipArtifactsDownload: false,
            timeoutInMinutes: 0,
            jobCancelTimeoutInMinutes: 1,
            condition: "succeeded()",
            overrideInputs: {},
            parallelExecution: { parallelExecutionType: "none" },
            artifactsDownloadInput: { downloadInputs: [] }
          }
        }]
      }],
      variables: {},
      variableGroups: [($variableGroupId | tonumber)],
      triggers: [],
      properties: {}
    }'
}

build_release_instance_payload() {
  jq -n \
    --arg definitionId "$RELEASE_DEFINITION_ID" \
    --arg artifactAlias "$RELEASE_ARTIFACT_ALIAS" \
    '{
      definitionId: ($definitionId | tonumber),
      description: "Generated by azure-devops-pipeline-repo-generator",
      artifacts: [{ alias: $artifactAlias, instanceReference: { id: "latest", name: "Latest" } }],
      isDraft: false
    }'
}

resolve_agent_queue() {
  if [[ -n "${RELEASE_AGENT_QUEUE_ID:-}" ]]; then
    RELEASE_AGENT_QUEUE_NAME="${RELEASE_AGENT_QUEUE_NAME:-}"
    return
  fi

  require_env RELEASE_AGENT_QUEUE_NAME
  echo "[INFO] Resolving release agent queue: $RELEASE_AGENT_QUEUE_NAME"
  local queue_response
  queue_response="$(api_get "$(project_api_url '_apis/distributedtask/queues' "$API_VERSION")")"
  RELEASE_AGENT_QUEUE_ID="$(echo "$queue_response" | jq -r --arg name "$RELEASE_AGENT_QUEUE_NAME" '.value[] | select(.name == $name) | .id' | head -n 1)"
  if [[ -z "$RELEASE_AGENT_QUEUE_ID" || "$RELEASE_AGENT_QUEUE_ID" == "null" ]]; then
    fail "Agent queue was not found." "RELEASE_AGENT_QUEUE_NAME=$RELEASE_AGENT_QUEUE_NAME\nCheck Project settings > Agent pools/queues or set RELEASE_AGENT_QUEUE_ID explicitly."
  fi
}

resolve_release_variable_group() {
  require_env RELEASE_VARIABLE_GROUP_NAME
  echo "[INFO] Resolving release variable group: $RELEASE_VARIABLE_GROUP_NAME"

  local encoded_group_name group_url group_response variable_name
  encoded_group_name="$(jq -rn --arg value "$RELEASE_VARIABLE_GROUP_NAME" '$value | @uri')"
  group_url="$(project_api_url '_apis/distributedtask/variablegroups' "$API_VERSION")&groupName=${encoded_group_name}&actionFilter=Use"
  group_response="$(api_get "$group_url")"
  RELEASE_VARIABLE_GROUP_ID="$(
    echo "$group_response" |
      jq -r --arg name "$RELEASE_VARIABLE_GROUP_NAME" '.value[] | select(.name == $name) | .id' |
      sort -n |
      head -n 1
  )"
  if [[ -z "$RELEASE_VARIABLE_GROUP_ID" || "$RELEASE_VARIABLE_GROUP_ID" == "null" ]]; then
    fail "Release variable group was not found or cannot be used." "RELEASE_VARIABLE_GROUP_NAME=$RELEASE_VARIABLE_GROUP_NAME\nCheck Pipelines > Library and the caller's variable-group permissions."
  fi

  for variable_name in AZP_TOKEN KOMODO_API_KEY KOMODO_API_SECRET; do
    if ! echo "$group_response" | jq -e \
      --arg groupName "$RELEASE_VARIABLE_GROUP_NAME" \
      --arg variableName "$variable_name" \
      '.value[] | select(.name == $groupName) | .variables | has($variableName)' >/dev/null; then
      fail "Release variable group is missing a required variable." "RELEASE_VARIABLE_GROUP_NAME=$RELEASE_VARIABLE_GROUP_NAME\nMissing variable: $variable_name"
    fi
  done
  echo "[INFO] Release variable group ID: $RELEASE_VARIABLE_GROUP_ID"
}

self_test() {
  require_command jq
  ADO_URL="${ADO_URL:-https://azure.example.local}"
  COLLECTION="${COLLECTION:-DefaultCollection}"
  PROJECT="${PROJECT:-SampleProject}"
  PIPELINE_NAME="${PIPELINE_NAME:-SampleRepo_demo}"
  REPO_NAME="${REPO_NAME:-SampleRepo}"
  YAML_PATH="${YAML_PATH:-/demo/pipeline.yml}"
  DEFAULT_BRANCH="${DEFAULT_BRANCH:-refs/heads/main}"
  PIPELINE_FOLDER="$(normalize_folder "${PIPELINE_FOLDER:-komodo}")"
  RELEASE_NAME="${RELEASE_NAME:-${PIPELINE_NAME}_Release}"
  RELEASE_FOLDER="$(normalize_folder "${RELEASE_FOLDER:-komodo}")"
  RELEASE_ENVIRONMENT_NAME="${RELEASE_ENVIRONMENT_NAME:-komodo}"
  RELEASE_AGENT_QUEUE_ID="${RELEASE_AGENT_QUEUE_ID:-1}"
  RELEASE_AGENT_QUEUE_NAME="${RELEASE_AGENT_QUEUE_NAME:-Default}"
  RELEASE_VARIABLE_GROUP_NAME="${RELEASE_VARIABLE_GROUP_NAME:-KomodoAPI}"
  RELEASE_VARIABLE_GROUP_ID="${RELEASE_VARIABLE_GROUP_ID:-7}"
  RELEASE_ARTIFACT_ALIAS="${RELEASE_ARTIFACT_ALIAS:-_${PIPELINE_NAME}}"
  RELEASE_BASH_TASK_NAME="${RELEASE_BASH_TASK_NAME:-Run inline Bash}"
  REPO_ID="00000000-0000-0000-0000-000000000001"
  PROJECT_ID="00000000-0000-0000-0000-000000000002"
  PIPELINE_ID="123"

  local release_payload
  build_pipeline_payload | jq -e '.folder == "\\komodo" and .configuration.path == "/demo/pipeline.yml"' >/dev/null
  release_payload="$(build_release_definition_payload $'#!/usr/bin/env bash\necho self-test')"
  echo "$release_payload" | jq -e '
    .artifacts[0].definitionReference.definition.id == "123" and
    .environments[0].deployPhases[0].workflowTasks[0].inputs.targetType == "inline" and
    .environments[0].conditions[0].name == "ReleaseStarted" and
    .environments[0].preDeployApprovals.approvals[0].isAutomated == true and
    .environments[0].postDeployApprovals.approvals[0].isAutomated == true and
    .variableGroups == [7]
  ' >/dev/null
  echo "Self-test passed. Pipeline and release payloads are valid JSON."
}

if [[ "${1:-}" == "--self-test" ]]; then
  step "Run local self-test"
  self_test
  exit 0
fi

step "Validate required local commands and environment variables"
require_command curl
require_command jq
require_env AZP_TOKEN
require_env ADO_URL
require_env COLLECTION
require_env PROJECT
require_env PIPELINE_NAME
require_env REPO_NAME
require_env YAML_PATH

API_VERSION="${API_VERSION:-7.1}"
RELEASE_API_VERSION="${RELEASE_API_VERSION:-7.1-preview.4}"
DEFAULT_BRANCH="${DEFAULT_BRANCH:-refs/heads/main}"
PIPELINE_FOLDER="$(normalize_folder "${PIPELINE_FOLDER:-komodo}")"
RELEASE_NAME="${RELEASE_NAME:-${PIPELINE_NAME}_Release}"
RELEASE_FOLDER="$(normalize_folder "${RELEASE_FOLDER:-komodo}")"
RELEASE_ENVIRONMENT_NAME="${RELEASE_ENVIRONMENT_NAME:-komodo}"
RELEASE_VARIABLE_GROUP_NAME="${RELEASE_VARIABLE_GROUP_NAME:-KomodoAPI}"
RELEASE_ARTIFACT_ALIAS="${RELEASE_ARTIFACT_ALIAS:-_${PIPELINE_NAME}}"
RELEASE_BASH_TASK_NAME="${RELEASE_BASH_TASK_NAME:-Run inline Bash}"
CREATE_RELEASE_INSTANCE="${CREATE_RELEASE_INSTANCE:-false}"

step "Load inline Bash task source"
INLINE_SCRIPT="$(read_inline_script)"

step "Resolve release agent queue"
resolve_agent_queue

step "Get Azure DevOps project ID"
PROJECT_RESPONSE="$(api_get "$(collection_api_url '_apis/projects' "$API_VERSION")")"
PROJECT_ID="$(echo "$PROJECT_RESPONSE" | jq -r --arg name "$PROJECT" '.value[] | select(.name == $name) | .id' | head -n 1)"
if [[ -z "$PROJECT_ID" || "$PROJECT_ID" == "null" ]]; then
  fail "Project was not found in the collection." "PROJECT=$PROJECT\nCOLLECTION=$COLLECTION\nADO_URL=$ADO_URL"
fi
echo "[INFO] Project ID: $PROJECT_ID"

step "Resolve Release variable group"
resolve_release_variable_group

step "Get target repository ID"
REPO_RESPONSE="$(api_get "$(project_api_url '_apis/git/repositories' "$API_VERSION")")"
REPO_ID="$(echo "$REPO_RESPONSE" | jq -r --arg name "$REPO_NAME" '.value[] | select(.name == $name) | .id' | head -n 1)"
if [[ -z "$REPO_ID" || "$REPO_ID" == "null" ]]; then
  fail "Repository was not found in the project." "REPO_NAME=$REPO_NAME\nPROJECT=$PROJECT"
fi
echo "[INFO] Repository ID: $REPO_ID"

step "Check existing YAML pipeline"
echo "[INFO] Pipeline name: $PIPELINE_NAME"
PIPELINES_RESPONSE="$(api_get "$(project_api_url '_apis/pipelines' "$API_VERSION")")"
PIPELINE_ID="$(echo "$PIPELINES_RESPONSE" | jq -r --arg name "$PIPELINE_NAME" '.value[] | select(.name == $name) | .id' | head -n 1)"
if [[ -n "$PIPELINE_ID" && "$PIPELINE_ID" != "null" ]]; then
  echo "[INFO] Pipeline already exists. ID=$PIPELINE_ID"
else
  step "Create YAML pipeline in komodo folder"
  echo "[INFO] Creating pipeline in folder: $PIPELINE_FOLDER"
  PIPELINE_PAYLOAD_FILE="$(mktemp)"
  build_pipeline_payload > "$PIPELINE_PAYLOAD_FILE"
  # Azure DevOps Server requires repositoryId on pipeline creation even though
  # the repository GUID is also present in the JSON body.
  PIPELINE_RESPONSE="$(api_post "$(pipeline_collection_url "$API_VERSION")" "$PIPELINE_PAYLOAD_FILE")"
  rm -f "$PIPELINE_PAYLOAD_FILE"
  echo "$PIPELINE_RESPONSE" | jq .
  PIPELINE_ID="$(echo "$PIPELINE_RESPONSE" | jq -r '.id')"
fi
if [[ -z "$PIPELINE_ID" || "$PIPELINE_ID" == "null" ]]; then
  fail "Pipeline creation response did not contain an ID." "Pipeline response:\n${PIPELINE_RESPONSE:-not-created}"
fi
echo "[SUCCESS] Pipeline ready. ID=$PIPELINE_ID"

step "Check existing classic release definition"
echo "[INFO] Release definition name: $RELEASE_NAME"
RELEASE_DEFINITIONS_RESPONSE="$(api_get "$(project_api_url '_apis/release/definitions' "$RELEASE_API_VERSION")")"
RELEASE_DEFINITION_ID="$(echo "$RELEASE_DEFINITIONS_RESPONSE" | jq -r --arg name "$RELEASE_NAME" '.value[] | select(.name == $name) | .id' | head -n 1)"
if [[ -n "$RELEASE_DEFINITION_ID" && "$RELEASE_DEFINITION_ID" != "null" ]]; then
  echo "[INFO] Release definition already exists. ID=$RELEASE_DEFINITION_ID"
else
  step "Create classic release definition with pipeline artifact and inline Bash task"
  echo "[INFO] Creating release definition in folder: $RELEASE_FOLDER"
  RELEASE_PAYLOAD_FILE="$(mktemp)"
  build_release_definition_payload "$INLINE_SCRIPT" > "$RELEASE_PAYLOAD_FILE"
  RELEASE_RESPONSE="$(api_post "$(project_api_url '_apis/release/definitions' "$RELEASE_API_VERSION")" "$RELEASE_PAYLOAD_FILE")"
  rm -f "$RELEASE_PAYLOAD_FILE"
  echo "$RELEASE_RESPONSE" | jq .
  RELEASE_DEFINITION_ID="$(echo "$RELEASE_RESPONSE" | jq -r '.id')"
fi
if [[ -z "$RELEASE_DEFINITION_ID" || "$RELEASE_DEFINITION_ID" == "null" ]]; then
  fail "Release definition response did not contain an ID." "Release response:\n${RELEASE_RESPONSE:-not-created}"
fi
echo "[SUCCESS] Release definition ready. ID=$RELEASE_DEFINITION_ID"

if [[ "$CREATE_RELEASE_INSTANCE" == "true" ]]; then
  step "Create release instance from definition"
  RELEASE_INSTANCE_PAYLOAD_FILE="$(mktemp)"
  build_release_instance_payload > "$RELEASE_INSTANCE_PAYLOAD_FILE"
  RELEASE_INSTANCE_RESPONSE="$(api_post "$(project_api_url '_apis/release/releases' "$RELEASE_API_VERSION")" "$RELEASE_INSTANCE_PAYLOAD_FILE")"
  rm -f "$RELEASE_INSTANCE_PAYLOAD_FILE"
  echo "$RELEASE_INSTANCE_RESPONSE" | jq .
  RELEASE_INSTANCE_ID="$(echo "$RELEASE_INSTANCE_RESPONSE" | jq -r '.id')"
  if [[ -z "$RELEASE_INSTANCE_ID" || "$RELEASE_INSTANCE_ID" == "null" ]]; then
    fail "Release instance response did not contain an ID." "Release instance response:\n$RELEASE_INSTANCE_RESPONSE"
  fi
  echo "[SUCCESS] Release instance created. ID=$RELEASE_INSTANCE_ID"
fi

echo "[DONE] Pipeline and release provisioning completed."
