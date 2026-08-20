#!/usr/bin/env node
'use strict';

const fs = require('fs');
const path = require('path');
const { spawnSync } = require('child_process');

const fail = (message) => {
  console.error(`##[error] ${message}`);
  process.exit(1);
};

const requiredEnv = (name) => {
  const value = String(process.env[name] || '').trim();
  if (!value) fail(`${name} is required.`);
  return value;
};

const sourceDir = path.resolve(requiredEnv('MR_SOURCE_DIR'));
const configPath = path.resolve(requiredEnv('MR_CONFIG_PATH'));
const artifactDir = path.resolve(requiredEnv('MR_ARTIFACT_DIR'));

const readScalar = (text, key, fallback = '') => {
  const match = text.match(new RegExp(`^\\s*${key}\\s*:\\s*(.+?)\\s*$`, 'm'));
  if (!match) return fallback;
  const raw = match[1].replace(/\s+#.*$/, '').trim();
  if (!raw) return fallback;
  if (raw.startsWith('"')) {
    try {
      return JSON.parse(raw);
    } catch {
      fail(`Invalid JSON-quoted ${key} in ${configPath}.`);
    }
  }
  if (raw.startsWith("'") && raw.endsWith("'")) {
    return raw.slice(1, -1).replace(/''/g, "'");
  }
  return raw;
};

const readList = (text, key, fallback = []) => {
  const lines = text.split(/\r?\n/);
  const index = lines.findIndex((line) => new RegExp(`^\\s*${key}\\s*:\\s*$`).test(line));
  if (index < 0) return fallback;
  const indent = (lines[index].match(/^\s*/) || [''])[0].length;
  const values = [];
  for (let cursor = index + 1; cursor < lines.length; cursor += 1) {
    const line = lines[cursor];
    if (!line.trim() || line.trim().startsWith('#')) continue;
    const currentIndent = (line.match(/^\s*/) || [''])[0].length;
    if (currentIndent <= indent) break;
    const match = line.match(/^\s*-\s*(.+?)\s*$/);
    if (!match) continue;
    let value = match[1].replace(/\s+#.*$/, '').trim();
    if (value.startsWith('"')) {
      try {
        value = JSON.parse(value);
      } catch {
        fail(`Invalid JSON-quoted list item under ${key}.`);
      }
    } else if (value.startsWith("'") && value.endsWith("'")) {
      value = value.slice(1, -1).replace(/''/g, "'");
    }
    if (value) values.push(value);
  }
  return values.length ? values : fallback;
};

if (!fs.existsSync(configPath)) fail(`Deployment contract not found: ${configPath}`);
const configText = fs.readFileSync(configPath, 'utf8');
const installCommand = readScalar(configText, 'install_command', 'pnpm install --frozen-lockfile');
const buildCommand = readScalar(
  configText,
  'build_command',
  'node tools/scripts/with-env.cjs production pnpm exec nx run-many -t build --projects={projects} --parallel=3'
);
const artifactName = readScalar(configText, 'artifact_name', 'mr-drop');
const bffEntry = readScalar(configText, 'bff_entry', 'main.js');
const shellNames = readList(configText, 'shell_projects', ['shell', 'host']);
const bffNames = readList(configText, 'bff_projects', ['bff']);
const continueOnModuleError = /^(1|true|yes)$/i.test(
  readScalar(configText, 'continue_on_module_error', 'true')
);

const validateProjectName = (name) => {
  if (!/^[A-Za-z0-9][A-Za-z0-9._-]{0,127}$/.test(name)) {
    fail(`Nx returned an unsafe project name: ${JSON.stringify(name)}`);
  }
  return name;
};

const run = (command, { capture = false, allowFailure = false } = {}) => {
  console.log(`##[command]${command}`);
  const result = spawnSync(command, {
    cwd: sourceDir,
    env: process.env,
    shell: true,
    encoding: 'utf8',
    stdio: capture ? ['ignore', 'pipe', 'pipe'] : 'inherit',
    maxBuffer: 32 * 1024 * 1024
  });
  if (result.error) fail(`Could not run command: ${result.error.message}`);
  if (result.status !== 0 && !allowFailure) {
    if (capture && result.stderr) process.stderr.write(result.stderr);
    fail(`Command exited with code ${result.status}: ${command}`);
  }
  return capture ? String(result.stdout || '').trim() : '';
};

const parseProjectList = (output, label) => {
  if (!output) return [];
  try {
    const parsed = JSON.parse(output);
    const values = Array.isArray(parsed) ? parsed : parsed.projects;
    if (!Array.isArray(values)) throw new Error('not an array');
    return [...new Set(values.map((value) => validateProjectName(String(value))))].sort();
  } catch {
    const values = output
      .split(/\r?\n|,/) 
      .map((value) => value.trim())
      .filter((value) => value && !value.startsWith('[') && !value.startsWith('>'));
    if (!values.length) fail(`Could not parse ${label} returned by Nx.`);
    return [...new Set(values.map(validateProjectName))].sort();
  }
};

const nxJson = (args, label, { allowEmpty = false } = {}) => {
  const output = run(`pnpm exec nx ${args}`, { capture: true });
  const parsed = parseProjectList(output, label);
  if (!parsed.length && !allowEmpty) fail(`Nx returned no ${label}.`);
  return parsed;
};

const git = (args, allowFailure = false) =>
  run(`git ${args}`, { capture: true, allowFailure });

const readProject = (name) => {
  const output = run(`pnpm exec nx show project ${JSON.stringify(name)} --json`, { capture: true });
  try {
    return JSON.parse(output);
  } catch {
    fail(`Nx returned invalid project metadata for ${name}.`);
  }
};

const classify = (name, metadata) => {
  const lower = name.toLowerCase();
  const tags = Array.isArray(metadata.tags) ? metadata.tags.map((tag) => String(tag).toLowerCase()) : [];
  if (shellNames.some((candidate) => candidate.toLowerCase() === lower) || tags.includes('deploy:shell')) {
    return 'shell';
  }
  if (bffNames.some((candidate) => candidate.toLowerCase() === lower) || tags.includes('deploy:bff')) {
    return 'bff';
  }
  return 'static';
};

process.chdir(sourceDir);
run(installCommand);

const applicationDiscovery = run('pnpm exec nx show projects --type app --with-target build --json', {
  capture: true,
  allowFailure: true
});
let allProjects;
try {
  allProjects = parseProjectList(applicationDiscovery, 'application projects');
} catch {
  allProjects = [];
}
if (!allProjects.length) {
  console.warn('##[warning] Nx application-type discovery was empty; falling back to buildable projects.');
  allProjects = nxJson('show projects --with-target build --json', 'buildable projects');
}

const projectMetadata = new Map(allProjects.map((name) => [name, readProject(name)]));
const projectKinds = new Map(
  allProjects.map((name) => [name, classify(name, projectMetadata.get(name))])
);
const bffProjects = allProjects.filter((name) => projectKinds.get(name) === 'bff');
if (bffProjects.length > 1) {
  fail(`The generic MR runtime supports one BFF, but Nx discovery classified: ${bffProjects.join(', ')}.`);
}

const head = git('rev-parse HEAD');
const explicitBase = String(process.env.MR_BASE_COMMIT || '').trim();
const parent = explicitBase || git(`rev-parse ${head}^`, true);
const buildAll = /^(1|true|yes)$/i.test(String(process.env.MR_BUILD_ALL || '')) ||
  String(process.env.BUILD_REASON || '').toLowerCase() === 'manual' ||
  !/^[0-9a-f]{40}$/i.test(parent);

let selectedProjects;
let shellChanged = false;
if (buildAll) {
  selectedProjects = [...allProjects];
  console.log('##[section]Full MR build selected (manual or first build).');
} else {
  selectedProjects = nxJson(
    `show projects --affected --base=${parent} --head=${head} --type app --with-target build --json`,
    'affected application projects',
    { allowEmpty: true }
  ).filter((name) => allProjects.includes(name));
  shellChanged = selectedProjects.some((name) => projectKinds.get(name) === 'shell');
  if (shellChanged) {
    selectedProjects = [...allProjects];
    console.log('##[section]A shell project changed; all MR applications will be rebuilt.');
  }
}

const successfulProjects = [];
const failedProjects = [];
for (const name of selectedProjects) {
  const command = buildCommand.replaceAll('{projects}', name);
  console.log(`##[section]Building MR module ${name}`);
  console.log(`##[command]${command}`);
  const result = spawnSync(command, {
    cwd: sourceDir,
    env: process.env,
    shell: true,
    encoding: 'utf8',
    stdio: 'inherit'
  });
  if (!result.error && result.status === 0) {
    successfulProjects.push(name);
    continue;
  }
  const reason = result.error?.message || `exit code ${result.status}`;
  failedProjects.push(name);
  console.error(`##vso[task.logissue type=warning]MR module ${name} failed to build (${reason}); its previous deployed version will be retained.`);
}

const failedShell = failedProjects.find((name) => projectKinds.get(name) === 'shell');
if (failedShell) {
  fail(`Shell module ${failedShell} failed. No MR artifact will be deployed because a shell change affects the whole application.`);
}
if (failedProjects.length && !continueOnModuleError) {
  fail(`MR module builds failed: ${failedProjects.join(', ')}.`);
}
if (selectedProjects.length && !successfulProjects.length) {
  fail('Every affected MR module failed to build; there is nothing safe to deploy.');
}
if (!selectedProjects.length) {
  console.log('##[section]No affected deployable Nx applications were found; publishing an inventory-only artifact.');
}

const normalizePath = (candidate) => {
  if (!candidate || typeof candidate !== 'string') return '';
  return candidate
    .replaceAll('{workspaceRoot}', sourceDir)
    .replace(/^\.\//, '')
    .trim();
};

const safeInsideSource = (candidate) => {
  const absolute = path.resolve(sourceDir, candidate);
  const relative = path.relative(sourceDir, absolute);
  if (!relative || relative.startsWith('..') || path.isAbsolute(relative)) return '';
  return absolute;
};

const findOutput = (name, metadata) => {
  const build = metadata.targets?.build || {};
  const optionPath = typeof build.options?.outputPath === 'string'
    ? build.options.outputPath
    : build.options?.outputPath?.base;
  const projectRoot = String(metadata.root || '').replace(/^\/+/, '');
  const outputTemplates = Array.isArray(build.outputs) ? build.outputs : [];
  const candidates = [
    optionPath,
    ...outputTemplates.map((value) => String(value)
      .replaceAll('{projectRoot}', projectRoot)
      .replaceAll('{options.outputPath}', optionPath || '')
    ),
    projectRoot ? `dist/${projectRoot}` : '',
    projectRoot ? `dist/apps/${path.basename(projectRoot)}` : '',
    `dist/${name}`
  ];
  for (const candidate of candidates) {
    const normalized = normalizePath(candidate);
    const absolute = safeInsideSource(normalized);
    if (absolute && fs.existsSync(absolute) && fs.statSync(absolute).isDirectory()) return absolute;
  }
  fail(
    `Could not locate the build output for ${name}. Configure targets.build.options.outputPath in Nx project metadata.`
  );
};

fs.rmSync(artifactDir, { recursive: true, force: true });
fs.mkdirSync(path.join(artifactDir, 'modules'), { recursive: true });
const selected = new Set(successfulProjects);
const inventory = [];
const packaged = [];

for (const name of allProjects) {
  const metadata = projectMetadata.get(name);
  const kind = projectKinds.get(name);
  const route = kind === 'shell' ? '/' : kind === 'bff' ? '/api/' : `/${name}/`;
  const record = { name, kind, route, root: String(metadata.root || '') };
  inventory.push(record);
  if (!selected.has(name)) continue;
  const output = findOutput(name, metadata);
  const destination = path.join(artifactDir, 'modules', name);
  fs.cpSync(output, destination, { recursive: true, force: true });
  packaged.push({ ...record, artifactPath: `modules/${name}` });
}

const projectKey = requiredEnv('MR_PROJECT_KEY');
const environment = requiredEnv('MR_ENVIRONMENT');
const buildId = requiredEnv('MR_BUILD_ID');
const deploymentRoot = requiredEnv('MR_DEPLOYMENT_ROOT');
const manifest = {
  schemaVersion: 1,
  kind: 'nx-monorepo',
  artifactName,
  buildId,
  project: projectKey,
  service: requiredEnv('MR_SERVICE_KEY'),
  environment,
  sourceCommit: head,
  baseCommit: /^[0-9a-f]{40}$/i.test(parent) ? parent : null,
  collectionUri: requiredEnv('SYSTEM_COLLECTIONURI'),
  azureProjectId: requiredEnv('SYSTEM_TEAMPROJECTID'),
  komodoServer: requiredEnv('MR_KOMODO_SERVER'),
  komodoStack: requiredEnv('MR_KOMODO_STACK'),
  deploymentRoot,
  staticContainer: requiredEnv('MR_STATIC_CONTAINER'),
  bffContainer: String(process.env.MR_BFF_CONTAINER || ''),
  bffProject: bffProjects[0] || '',
  bffEntry,
  composeRepository: requiredEnv('MR_COMPOSE_REPOSITORY'),
  composePath: requiredEnv('MR_COMPOSE_PATH'),
  modules: inventory,
  packagedModules: packaged.map(({ name }) => name),
  failedModules: failedProjects
};

const tsv = (records) => records
  .map(({ name, kind, route }) => [name, kind, route].join('\t'))
  .join('\n') + (records.length ? '\n' : '');
fs.writeFileSync(path.join(artifactDir, 'manifest.json'), `${JSON.stringify(manifest, null, 2)}\n`);
fs.writeFileSync(path.join(artifactDir, 'inventory.tsv'), tsv(inventory));
fs.writeFileSync(path.join(artifactDir, 'modules.tsv'), tsv(packaged));
console.log(`##[section]MR artifact ready: ${packaged.length} packaged / ${inventory.length} discovered applications.`);
if (failedProjects.length) {
  console.log(`##vso[task.complete result=SucceededWithIssues;]Deployed successful modules; retained previous versions for: ${failedProjects.join(', ')}`);
}
