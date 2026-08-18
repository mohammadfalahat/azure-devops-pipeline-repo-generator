(() => {
  const getHostBase = () => {
    if (!document.referrer) {
      return window.location.origin;
    }

  const referrer = new URL(document.referrer);
  const segments = referrer.pathname.split('/').filter(Boolean);
  const hasTfsVirtualDir = segments[0]?.toLowerCase() === 'tfs';
  const collectionSegment = hasTfsVirtualDir ? segments[1] : segments[0];

  const pathSegments = [referrer.origin];
  if (hasTfsVirtualDir) {
    pathSegments.push('tfs');
  }
  if (collectionSegment) {
    pathSegments.push(collectionSegment);
  }

  return pathSegments.join('/');
};

  const normalizeHostUri = (hostUri) => {
    if (!hostUri) return '';

    const trimmed = hostUri.replace(/\/+$/, '');
    const withoutApis = trimmed.replace(/\/_apis\/?$/i, '');

    return `${withoutApis.replace(/\/+$/, '')}/`;
  };

  const buildCollectionUri = (hostUri, collectionName) => {
    const normalizedCollection = String(collectionName || '').trim();
    if (!normalizedCollection || /[\\/\0\r\n]/.test(normalizedCollection)) {
      throw new Error('Central Azure DevOps collection name is invalid.');
    }

    let parsed;
    try {
      parsed = new URL(normalizeHostUri(hostUri));
    } catch {
      throw new Error('Azure DevOps collection URI is invalid.');
    }
    if (!/^https?:$/.test(parsed.protocol)) {
      throw new Error('Azure DevOps collection URI must use HTTP or HTTPS.');
    }

    const currentSegments = parsed.pathname.split('/').filter(Boolean);
    const serverPathSegments = currentSegments.length > 0 ? currentSegments.slice(0, -1) : [];
    const centralPath = [...serverPathSegments, encodeURIComponent(normalizedCollection)].join('/');
    return `${parsed.origin}/${centralPath}/`;
  };

  const buildCentralGitItemUrl = ({ hostUri, source }) => {
    const collectionUri = buildCollectionUri(hostUri, source.collection);
    return `${collectionUri}${encodeURIComponent(source.project)}/_apis/git/repositories/${encodeURIComponent(
      source.repository
    )}/items?path=${encodeURIComponent(source.path)}&versionDescriptor.version=${encodeURIComponent(
      source.branch
    )}&versionDescriptor.versionType=branch&%24format=text&api-version=6.0`;
  };

  const loadScript = (src) =>
    new Promise((resolve, reject) => {
      const script = document.createElement('script');
      script.src = src;
      script.async = false;
      script.onload = resolve;
      script.onerror = () => reject(new Error(`Failed to load Azure DevOps SDK from ${src}`));
      document.head.appendChild(script);
    });

  const hasCoreSdkApis = (sdk) =>
    Boolean(
      sdk &&
        sdk.init &&
        sdk.ready &&
        sdk.getAccessToken &&
        sdk.getService &&
        (sdk.getWebContext || sdk.getHostContext)
    );

  const normalizeSdk = (sdk) => {
    if (!sdk) return sdk;

    const getHostContext = () => {
      const webContext = sdk.getWebContext?.();
      const hostFromWeb = webContext?.host || webContext?.collection;
      const host = sdk.getHostContext?.()?.host || hostFromWeb || {};
      return {
        host: {
          name: host.name || webContext?.collection?.name,
          uri: host.uri || hostFromWeb?.uri || getHostBase(),
          relativeUri: host.relativeUri || '/',
          hostType: host.hostType || webContext?.host?.hostType,
          id: host.id || webContext?.host?.id
        }
      };
    };

    if (!sdk.getHostContext) {
      sdk.getHostContext = getHostContext;
    }
    if (!sdk.getWebContext) {
      sdk.getWebContext = () => ({ host: getHostContext().host });
    }
    if (!sdk.notifyLoadSucceeded) {
      sdk.notifyLoadSucceeded = () => {};
    }
    if (!sdk.notifyLoadFailed) {
      sdk.notifyLoadFailed = () => {};
    }

    return sdk;
  };

  const loadVssSdk = async () => {
    const ambientSdk = normalizeSdk(window.VSS || window.parent?.VSS);
    if (hasCoreSdkApis(ambientSdk)) {
      return ambientSdk;
    }

    const localSdk = new URL('./lib/VSS.SDK.min.js', window.location.href).toString();
    const localSdkFallback = new URL('./lib/VSS.SDK.js', window.location.href).toString();
    // Only load bundled SDK assets. Some on-prem Azure DevOps hosts challenge
    // requests to the platform SDK endpoint with browser-level Basic auth, which
    // causes repeated username/password popups even when the extension already
    // has a valid access token.
    const candidates = [localSdk, localSdkFallback];

    let lastError;
    for (const src of candidates) {
      try {
        await loadScript(src);
        if (hasCoreSdkApis(window.VSS)) {
          return normalizeSdk(window.VSS);
        }
        lastError = new Error('Azure DevOps SDK was loaded but did not initialize.');
      } catch (error) {
        lastError = error;
      }
    }

    throw lastError || new Error('Failed to load Azure DevOps SDK.');
  };

  const waitForSdkReady = async (sdk, timeoutMs = 15000) => {
    if (!sdk?.ready) {
      return;
    }

    await new Promise((resolve, reject) => {
      let settled = false;
      const timer = setTimeout(() => {
        if (settled) return;
        settled = true;
        reject(new Error('Timed out waiting for Azure DevOps host to respond.'));
      }, timeoutMs);

      try {
        sdk.ready(() => {
          if (settled) return;
          settled = true;
          clearTimeout(timer);
          resolve();
        });
      } catch (error) {
        if (settled) return;
        settled = true;
        clearTimeout(timer);
        reject(error);
      }
    });
  };

  const defaultValues = {
    pool: 'PublishDockerAgent',
    environment: 'demo',
    repositoryAddress: 'registry.buluttakin.com',
    containerRegistryService: 'BulutReg',
    dockerfileDir: '**'
  };
  const defaultPoolOptions = ['PublishDockerAgent', 'Default'];
  const defaultRegistryOptions = ['BulutReg', 'DockerReg'];
  const CENTRAL_COLLECTION_NAME = 'ShonizCollection';
  const DEPLOYMENT_TARGETS_CONFIG = Object.freeze({
    collection: CENTRAL_COLLECTION_NAME,
    project: 'SharedTemplates',
    repository: 'SharedTemplates',
    path: '/pipeline-generator.yml',
    branch: 'main'
  });
  const KOMODO_CREDENTIAL_CONFIG = Object.freeze({
    collection: CENTRAL_COLLECTION_NAME,
    project: 'SharedTemplates',
    repository: 'SharedTemplates',
    path: '/komodo-servers-creds.env',
    branch: 'main'
  });

  const mergeWithDefaults = (defaults, values) => {
    const seen = new Set();
    const combined = [];
    [...defaults, ...values].forEach((value) => {
      if (!value || seen.has(value)) return;
      seen.add(value);
      combined.push(value);
    });
    return combined;
  };

  const getQueryValue = (value) => (value && value !== 'undefined' && value !== 'null' ? value : undefined);

  const getDialogConfiguration = (sdk) => {
    const configuration = sdk?.getConfiguration?.();
    if (!configuration || typeof configuration !== 'object') return {};

    // Current hosts return the object supplied to openCustomDialog directly.
    // Accept the nested shapes as well for compatibility with older wrappers.
    return configuration.pipelineBootstrap || configuration.configuration || configuration;
  };

  const getHostNavigationState = async (sdk) => {
    const serviceIds = [
      sdk?.ServiceIds?.Navigation,
      'ms.vss-features.host-navigation-service'
    ].filter((value, index, values) => value && values.indexOf(value) === index);

    for (const serviceId of serviceIds) {
      try {
        const navigationService = await sdk.getService(serviceId);
        if (navigationService?.getCurrentState) {
          return navigationService.getCurrentState() || {};
        }
        if (navigationService?.getQueryParams) {
          return (await navigationService.getQueryParams()) || {};
        }
      } catch (error) {
        console.warn('Could not read Pipeline Generator context from host navigation state', {
          serviceId,
          message: error?.message
        });
      }
    }

    return {};
  };

  const branchLabel = document.getElementById('branch-label');
  const branchInput = document.getElementById('branch');
  const environmentSelect = document.getElementById('environment');
  const poolSelect = document.getElementById('pool');
  const serviceInput = document.getElementById('service');
  const registrySelect = document.getElementById('containerRegistryService');
  const dockerfileInput = document.getElementById('dockerfileDir');
  const form = document.getElementById('pipeline-form');
  const status = document.getElementById('status');
  const targetRepoInput = document.getElementById('targetRepo');
  const komodoSelect = document.getElementById('komodoServer');
  const reauthPanel = document.getElementById('reauth-panel');
  const reauthMessage = document.getElementById('reauth-message');
  const authorizeExtensionButton = document.getElementById('authorize-extension');
  const reauthenticateButton = document.getElementById('reauthenticate');
  const submitButton = form?.querySelector('button[type="submit"]');
  const completionPanel = document.getElementById('completion-panel');
  const nginxResultLink = document.getElementById('nginx-result-link');
  const composeResultLink = document.getElementById('compose-result-link');
  const pipelineResultLink = document.getElementById('pipeline-result-link');

  if (targetRepoInput) {
    targetRepoInput.disabled = true;
  }

  if (serviceInput) {
    serviceInput.addEventListener('input', () => {
      serviceInput.dataset.autofilled = 'false';
    });
  }

  const SCAFFOLD_BRANCH = 'main';
  const ZERO_OBJECT_ID = '0000000000000000000000000000000000000000';
  const PIPELINE_FOLDER = '\\komodo';
  // The Pipelines API remains a preview contract on supported Azure DevOps
  // Server versions. Keep the repositoryId query parameter on create: without
  // it, some on-prem servers accept the YAML commit but reject pipeline
  // registration because they cannot resolve the input repository.
  const PIPELINE_API_VERSION = '7.1-preview.1';
  const BUILD_API_VERSION = '7.1';
  const RELEASE_API_VERSION = '7.1-preview.4';
  const BASH_TASK_ID = '6c731c3c-3c68-459a-a5c9-bde6e6595b5b';
  const DEFAULT_RELEASE_CONFIG = Object.freeze({
    enabled: true,
    folder: '\\komodo',
    environmentName: 'komodo',
    bashTaskName: 'Run Komodo deployment',
    variableGroupName: 'KomodoAPI',
    requiredVariableNames: Object.freeze(['AZP_TOKEN', 'KOMODO_API_KEY', 'KOMODO_API_SECRET']),
    scriptSource: { type: 'inline', content: '' }
  });

  const normalizePipelineFolder = (folder, fallback) => {
    const candidate = String(folder || fallback || '').trim().replace(/\//g, '\\');
    if (!candidate || candidate === '\\') return '\\';
    return `\\${candidate.replace(/^\\+/, '')}`;
  };

  const getReleaseConfig = () => {
    const configured = window.PipelineGeneratorReleaseConfig || {};
    const source = configured.scriptSource || DEFAULT_RELEASE_CONFIG.scriptSource;
    return {
      enabled: configured.enabled !== false,
      folder: normalizePipelineFolder(configured.folder, DEFAULT_RELEASE_CONFIG.folder),
      environmentName: String(configured.environmentName || DEFAULT_RELEASE_CONFIG.environmentName).trim(),
      bashTaskName: String(configured.bashTaskName || DEFAULT_RELEASE_CONFIG.bashTaskName).trim(),
      variableGroupName: String(
        configured.variableGroupName || DEFAULT_RELEASE_CONFIG.variableGroupName
      ).trim(),
      requiredVariableNames: Array.isArray(configured.requiredVariableNames)
        ? configured.requiredVariableNames.map((name) => String(name).trim()).filter(Boolean)
        : [...DEFAULT_RELEASE_CONFIG.requiredVariableNames],
      scriptSource: source
    };
  };

  const state = {
    sdk: null,
    accessToken: null,
    accessTokenError: null,
    hostUri: null,
    projectId: null,
    rawProjectName: null,
    projectName: null,
    repoId: null,
    rawRepositoryName: null,
    repositoryName: null,
    generatedRepoId: null,
    generatedRepositoryName: null,
    deploymentTargets: null,
    deploymentTargetsReady: false,
    provisioningComplete: false,
    branch: SCAFFOLD_BRANCH,
    sourceBranch: null
  };
  let initializationPromise;

  const setStatus = (message, isError = false) => {
    status.textContent = message;
    status.className = isError ? 'status-error' : 'status-success';
  };

  const setSubmitting = (isSubmitting) => {
    if (submitButton) {
      submitButton.disabled = isSubmitting || state.provisioningComplete || !state.deploymentTargetsReady;
    }
  };

  const setReauthenticationVisibility = (isVisible, message) => {
    if (!reauthPanel) return;
    if (message && reauthMessage) {
      reauthMessage.textContent = message;
    }
    reauthPanel.classList?.toggle('hidden', !isVisible);
    if (!reauthPanel.classList) {
      reauthPanel.className = isVisible ? 'auth-fallback' : 'auth-fallback hidden';
    }
  };

  const setCompletionVisibility = (isVisible) => {
    completionPanel?.classList?.toggle('hidden', !isVisible);
    if (completionPanel && !completionPanel.classList) {
      completionPanel.className = isVisible ? 'completion-panel' : 'completion-panel hidden';
    }
  };

  const finishProvisioning = () => {
    state.provisioningComplete = true;
    if (form) {
      form.hidden = true;
    }
    setSubmitting(false);
    setCompletionVisibility(true);
    completionPanel?.focus?.({ preventScroll: true });
    completionPanel?.scrollIntoView?.({ behavior: 'smooth', block: 'start' });
  };

  const loadPools = async ({ hostUri, projectId, accessToken }) => {
    if (!poolSelect) return [];
    try {
      const dynamicPools = await fetchAgentQueues({ hostUri, projectId, accessToken });
      const options = mergeWithDefaults(defaultPoolOptions, dynamicPools).map((name) => ({ value: name, label: name }));
      populateSelectOptions(poolSelect, options);
      poolSelect.value = poolSelect.value || defaultValues.pool;
      return options;
    } catch (error) {
      console.warn('Falling back to default pools', error);
      const fallback = defaultPoolOptions.map((name) => ({ value: name, label: name }));
      populateSelectOptions(poolSelect, fallback);
      poolSelect.value = defaultValues.pool;
      return fallback;
    }
  };

  const loadContainerRegistries = async ({ hostUri, projectId, accessToken }) => {
    if (!registrySelect) return [];
    try {
      const registries = await fetchContainerRegistries({ hostUri, projectId, accessToken });
      const options = mergeWithDefaults(defaultRegistryOptions, registries).map((name) => ({ value: name, label: name }));
      populateSelectOptions(registrySelect, options);
      registrySelect.value = registrySelect.value || defaultValues.containerRegistryService;
      return options;
    } catch (error) {
      console.warn('Falling back to default container registries', error);
      const fallback = defaultRegistryOptions.map((name) => ({ value: name, label: name }));
      populateSelectOptions(registrySelect, fallback);
      registrySelect.value = defaultValues.containerRegistryService;
      return fallback;
    }
  };

  const parseDeploymentTargetScalar = (rawValue, lineNumber) => {
    const raw = String(rawValue || '').trim();
    let value;
    if (raw.startsWith('"')) {
      const match = raw.match(/^("(?:\\.|[^"\\])*")\s*(?:#.*)?$/);
      if (!match) {
        throw new Error(`Invalid double-quoted value in pipeline-generator.yml at line ${lineNumber}.`);
      }
      try {
        value = JSON.parse(match[1]);
      } catch (error) {
        throw new Error(`Invalid double-quoted value in pipeline-generator.yml at line ${lineNumber}.`);
      }
    } else if (raw.startsWith("'")) {
      const match = raw.match(/^('(?:''|[^'])*')\s*(?:#.*)?$/);
      if (!match) {
        throw new Error(`Invalid single-quoted value in pipeline-generator.yml at line ${lineNumber}.`);
      }
      value = match[1].slice(1, -1).replace(/''/g, "'");
    } else {
      value = raw.replace(/\s+#.*$/, '').trim();
    }

    if (typeof value !== 'string' || !value.trim()) {
      throw new Error(`Empty deployment target in pipeline-generator.yml at line ${lineNumber}.`);
    }
    value = value.trim();
    if (value.length > 200 || /[\0\r\n]/.test(value) || value.includes("'")) {
      throw new Error(`Unsafe deployment target in pipeline-generator.yml at line ${lineNumber}.`);
    }
    return value;
  };

  const uniqueCaseInsensitive = (values) => {
    const seen = new Set();
    return values.filter((value) => {
      const key = value.toLowerCase();
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });
  };

  const parseDeploymentTargetsYaml = (yamlText) => {
    const result = { servers: [], environments: [], environmentConfigs: [] };
    let section;
    let currentEnvironment;
    String(yamlText || '')
      .replace(/^\uFEFF/, '')
      .split(/\r?\n/)
      .forEach((line, index) => {
        const lineNumber = index + 1;
        const trimmed = line.trim();
        if (!trimmed || trimmed.startsWith('#')) return;

        const sectionMatch = line.match(/^([A-Za-z][A-Za-z0-9_-]*):\s*(?:#.*)?$/);
        if (sectionMatch) {
          section = ['servers', 'environments'].includes(sectionMatch[1]) ? sectionMatch[1] : undefined;
          currentEnvironment = undefined;
          return;
        }

        if (section === 'servers') {
          const listMatch = line.match(/^\s+-\s+(.+?)\s*$/);
          if (!listMatch) {
            throw new Error(`Expected a YAML list item in servers at line ${lineNumber}.`);
          }
          result.servers.push(parseDeploymentTargetScalar(listMatch[1], lineNumber));
          return;
        }

        if (section === 'environments') {
          const listMatch = line.match(/^\s+-\s+(.+?)\s*$/);
          if (listMatch) {
            const inlineName = listMatch[1].match(/^name\s*:\s*(.+)$/i);
            if (inlineName) {
              currentEnvironment = {
                name: parseDeploymentTargetScalar(inlineName[1], lineNumber),
                domain: ''
              };
            } else {
              const scalar = parseDeploymentTargetScalar(listMatch[1], lineNumber);
              const separator = scalar.indexOf(':');
              currentEnvironment = {
                name: separator > 0 ? scalar.slice(0, separator).trim() : scalar,
                domain: separator > 0 ? scalar.slice(separator + 1).trim() : ''
              };
            }
            result.environmentConfigs.push(currentEnvironment);
            return;
          }

          const propertyMatch = line.match(/^\s+(name|domain)\s*:\s*(.+?)\s*$/i);
          if (!propertyMatch || !currentEnvironment) {
            throw new Error(`Expected an environment name/domain entry at line ${lineNumber}.`);
          }
          const property = propertyMatch[1].toLowerCase();
          if (currentEnvironment[property]) {
            throw new Error(`Duplicate environment ${property} at line ${lineNumber}.`);
          }
          currentEnvironment[property] = parseDeploymentTargetScalar(propertyMatch[2], lineNumber);
        }
      });

    result.servers = uniqueCaseInsensitive(result.servers);
    if (!result.environmentConfigs.length) {
      throw new Error('pipeline-generator.yml must contain a non-empty environments list.');
    }
    const seenEnvironments = new Set();
    result.environmentConfigs.forEach((environment) => {
      const name = String(environment.name || '').trim();
      const domain = String(environment.domain || '').trim().toLowerCase();
      if (!/^[A-Za-z0-9][A-Za-z0-9._-]*$/.test(name)) {
        throw new Error(`Environment contains unsupported path characters: ${name || '(empty)'}.`);
      }
      if (
        !/^(?=.{1,253}$)(?:[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?\.)+[a-z]{2,63}$/i.test(domain)
      ) {
        throw new Error(`Environment ${name} must define a valid domain.`);
      }
      const key = name.toLowerCase();
      if (seenEnvironments.has(key)) {
        throw new Error(`Duplicate environment name in pipeline-generator.yml: ${name}.`);
      }
      seenEnvironments.add(key);
      environment.name = name;
      environment.domain = domain;
    });
    result.environments = result.environmentConfigs.map((environment) => environment.name);
    return result;
  };

  const fetchDeploymentTargets = async ({ hostUri, accessToken }) => {
    const source = DEPLOYMENT_TARGETS_CONFIG;
    const url = buildCentralGitItemUrl({ hostUri, source });
    const res = await fetch(url, { headers: authHeaders(accessToken), cache: 'no-store' });
    if (!res.ok) {
      const detail = await readErrorDetail(res);
      throw markRequiredExtensionScope(
        buildHttpError(`Failed to load deployment targets from ${source.path}`, res, detail),
        'vso.code'
      );
    }
    return parseDeploymentTargetsYaml(await res.text());
  };

  const parseKomodoCredentialFile = (text) => {
    const values = {};
    const supported = new Set(['KOMODO_ADDRESS', 'KOMODO_API_KEY', 'KOMODO_API_SECRET']);
    String(text || '')
      .replace(/^\uFEFF/, '')
      .split(/\r?\n/)
      .forEach((line, index) => {
        const trimmed = line.trim();
        if (!trimmed || trimmed.startsWith('#')) return;
        const match = trimmed.match(/^([A-Z][A-Z0-9_]*)\s*=\s*(.*)$/);
        if (!match || !supported.has(match[1])) {
          throw new Error(`Unsupported credential-file entry at line ${index + 1}.`);
        }
        if (Object.prototype.hasOwnProperty.call(values, match[1])) {
          throw new Error(`Duplicate ${match[1]} entry in the Komodo credential file.`);
        }
        let value = match[2].trim();
        if (value.startsWith('"')) {
          try {
            value = JSON.parse(value);
          } catch {
            throw new Error(`Invalid quoted credential-file value at line ${index + 1}.`);
          }
        } else if (value.startsWith("'")) {
          if (!/^'(?:''|[^'])*'$/.test(value)) {
            throw new Error(`Invalid quoted credential-file value at line ${index + 1}.`);
          }
          value = value.slice(1, -1).replace(/''/g, "'");
        }
        if (!value || value.length > 2048 || /[\0\r\n]/.test(value)) {
          throw new Error(`Invalid credential-file value at line ${index + 1}.`);
        }
        values[match[1]] = value;
      });

    const missing = [...supported].filter((name) => !values[name]);
    if (missing.length) {
      throw new Error(`Komodo credential file is missing: ${missing.join(', ')}.`);
    }
    if (!/^https:\/\/[^\s]+$/i.test(values.KOMODO_ADDRESS)) {
      throw new Error('KOMODO_ADDRESS in the credential file must be an HTTPS URL.');
    }
    return Object.freeze({
      address: values.KOMODO_ADDRESS.replace(/\/+$/, ''),
      apiKey: values.KOMODO_API_KEY,
      apiSecret: values.KOMODO_API_SECRET
    });
  };

  const fetchKomodoCredentials = async ({ hostUri, accessToken }) => {
    const source = KOMODO_CREDENTIAL_CONFIG;
    const url = buildCentralGitItemUrl({ hostUri, source });
    const res = await fetch(url, { headers: authHeaders(accessToken), cache: 'no-store' });
    if (!res.ok) {
      const detail = await readErrorDetail(res);
      const error = markRequiredExtensionScope(
        buildHttpError(`Failed to load Komodo credentials from ${source.path}`, res, detail),
        'vso.code'
      );
      error.domain = 'komodo';
      throw error;
    }
    try {
      return parseKomodoCredentialFile(await res.text());
    } catch (error) {
      error.domain = 'komodo';
      throw error;
    }
  };

  const extractEnabledKomodoServers = (payload) => {
    const records = Array.isArray(payload)
      ? payload
      : Array.isArray(payload?.data)
        ? payload.data
        : Array.isArray(payload?.value)
          ? payload.value
          : Array.isArray(payload?.servers)
            ? payload.servers
            : null;
    if (!records) {
      throw new Error('Komodo returned an unsupported ListFullServers response.');
    }
    const servers = uniqueCaseInsensitive(
      records
        .map((record) => record?.data || record)
        .filter((record) => record && record.template !== true && record.config?.enabled === true)
        .map((record) => String(record.name || '').trim())
        .filter(Boolean)
    ).sort((left, right) => left.localeCompare(right, 'en', { sensitivity: 'base' }));
    const unsafeServer = servers.find((server) => server.length > 200 || /[\0\r\n']/.test(server));
    if (unsafeServer || !servers.length) {
      throw new Error(
        unsafeServer ? 'Komodo returned an unsafe server name.' : 'Komodo returned no enabled servers.'
      );
    }
    return servers;
  };

  const fetchKomodoServers = async ({ hostUri, accessToken }) => {
    try {
      const credentials = await fetchKomodoCredentials({ hostUri, accessToken });
      const res = await fetch(`${credentials.address}/read`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'X-Api-Key': credentials.apiKey,
          'X-Api-Secret': credentials.apiSecret
        },
        body: JSON.stringify({ type: 'ListFullServers', params: { query: {} } }),
        cache: 'no-store',
        credentials: 'omit',
        referrerPolicy: 'no-referrer'
      });
      if (!res.ok) {
        const error = new Error(`Komodo ListFullServers returned HTTP ${res.status}.`);
        error.status = res.status;
        throw error;
      }
      return extractEnabledKomodoServers(await res.json());
    } catch (error) {
      error.domain = 'komodo';
      throw error;
    }
  };

  const loadDeploymentTargets = async ({ hostUri, accessToken, branch }) => {
    state.deploymentTargetsReady = false;
    if (environmentSelect) environmentSelect.disabled = true;
    if (komodoSelect) komodoSelect.disabled = true;
    try {
      const [targets, komodoServers] = await Promise.all([
        fetchDeploymentTargets({ hostUri, accessToken }),
        fetchKomodoServers({ hostUri, accessToken })
      ]);
      targets.servers = komodoServers;
      state.deploymentTargets = targets;
      populateSelectOptions(
        environmentSelect,
        targets.environmentConfigs.map(({ name, domain }) => ({
          value: name,
          label: `${name} — ${domain}`
        }))
      );
      populateSelectOptions(
        komodoSelect,
        targets.servers.map((value) => ({ value, label: value }))
      );
      if (environmentSelect) {
        const defaultEnvironment = targets.environments.find(
          (value) => value.toLowerCase() === defaultValues.environment.toLowerCase()
        );
        environmentSelect.value = defaultEnvironment || targets.environments[0];
        environmentSelect.disabled = false;
      }
      if (komodoSelect) {
        komodoSelect.disabled = false;
      }
      applyDetectedEnvironment(branch);
      setKomodoServerFromEnvironment(environmentSelect?.value);
      state.deploymentTargetsReady = true;
      return targets;
    } catch (error) {
      state.deploymentTargets = null;
      state.deploymentTargetsReady = false;
      populateSelectOptions(environmentSelect, [], 'Deployment environments unavailable');
      populateSelectOptions(komodoSelect, [], 'Active Komodo servers unavailable');
      if (environmentSelect) environmentSelect.disabled = true;
      if (komodoSelect) komodoSelect.disabled = true;
      throw error;
    }
  };

  const refreshDockerfiles = async ({ hostUri, projectId, repoId, branch, accessToken }) => {
    if (!dockerfileInput || !accessToken || !projectId || !repoId || !hostUri) return [];
    dockerfileInput.value = defaultValues.dockerfileDir || '';
    try {
      const dockerfiles = await fetchDockerfileDirectories({ hostUri, projectId, repoId, branch, accessToken });
      if (dockerfiles.length) {
        const defaultPath = dockerfiles[0];
        dockerfileInput.value = defaultPath;
      } else {
        dockerfileInput.value = defaultValues.dockerfileDir || '';
        setStatus('No Dockerfile was found in this branch. Please provide the directory manually.', true);
      }
      return dockerfiles;
    } catch (error) {
      console.error(error);
      dockerfileInput.value = defaultValues.dockerfileDir || '';
      setStatus('Could not auto-detect Dockerfile location. Please fill it manually.', true);
      return [];
    }
  };

  const delay = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

  const normalizeAccessTokenError = (error) => {
    const message = error?.message || 'Unknown Azure DevOps authentication error';
    if (/HostAuthorizationNotFound/i.test(message)) {
      return 'HostAuthorizationNotFound: A Collection Administrator must open Collection Settings → Extensions, select Pipeline Generator, and authorize its requested scopes. If no authorization action is available, reinstall the same published extension version.';
    }
    return message;
  };

  const isHostAuthorizationError = (value) =>
    /HostAuthorizationNotFound|Host authorization was not found/i.test(value?.message || value || '');

  const buildTokenRecoveryMessage = (errorMessage) =>
    isHostAuthorizationError(errorMessage)
      ? `${errorMessage} Use Open extension authorization below; signing out cannot create the missing extension authorization.`
      : `${errorMessage} Sign out and authenticate again below to rebuild the Azure DevOps host session.`;

  const getAccessTokenWithRetry = async (sdk, maxAttempts = 3, delayMs = 800) => {
    if (!sdk?.getAccessToken) {
      return undefined;
    }

    let lastError;
    for (let attempt = 1; attempt <= maxAttempts; attempt += 1) {
      try {
        const token = await sdk.getAccessToken();
        if (token) {
          return token;
        }
        lastError = new Error('Azure DevOps returned an empty access token.');
      } catch (error) {
        lastError = error;
        console.warn('[pipeline-generator] getAccessToken failed', {
          attempt,
          status: error?.status,
          message: error?.message
        });

        if (error?.status === 500) {
          break;
        }
      }

      if (attempt < maxAttempts) {
        await delay(delayMs * attempt);
      }
    }

    throw lastError;
  };

  const getAuthHeader = (token) => {
    const tokenValue = typeof token === 'string' ? token : token?.token;
    if (!tokenValue) {
      throw new Error('Extension access token was unavailable.');
    }

    // Azure DevOps SDK access tokens are OAuth/session tokens and must be sent
    // as Bearer credentials. Treating opaque (non-JWT) tokens as PATs and
    // forcing Basic auth causes repeated browser username/password prompts on
    // on-prem hosts whenever the server challenges unauthorized requests.
    return `Bearer ${tokenValue}`;
  };

  const authHeaders = (token) => ({
    Authorization: getAuthHeader(token),
    'X-TFS-FedAuthRedirect': 'Suppress'
  });

  const sanitizeErrorDetail = (detail = '') =>
    detail
      .replace(/<script[\s\S]*?<\/script>/gi, ' ')
      .replace(/<style[\s\S]*?<\/style>/gi, ' ')
      .replace(/<[^>]+>/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();

  const readErrorDetail = async (response) => {
    try {
      const text = await response.text();
      const sanitized = sanitizeErrorDetail(text || '');
      return sanitized.length > 500 ? `${sanitized.slice(0, 497)}...` : sanitized;
    } catch (error) {
      console.warn('Failed to read error response body', error);
      return '';
    }
  };

  const logAuthDiagnostics = (response, baseMessage) => {
    if (!response) return;

    const wwwAuthenticate = response.headers?.get('www-authenticate') || '';
    const shouldLog = response.status === 401 || response.status === 403 || Boolean(wwwAuthenticate);
    if (!shouldLog) return;

    console.warn('[pipeline-generator] Authorization challenge detected', {
      operation: baseMessage,
      status: response.status,
      url: response.url,
      wwwAuthenticate,
      fedAuthRedirect: response.headers?.get('x-tfs-fedauthredirect') || ''
    });
  };

  const buildHttpError = (baseMessage, response, detail) => {
    logAuthDiagnostics(response, baseMessage);
    const message = `${baseMessage} (${response.status})${detail ? `: ${detail}` : ''}`;
    const error = new Error(message);
    error.status = response.status;
    error.url = response.url;
    if (detail) {
      error.detail = detail;
    }
    return error;
  };

  const markErrorDomain = (error, domain) => {
    if (error && !error.domain) {
      error.domain = domain;
    }
    return error;
  };

  const markRequiredExtensionScope = (error, scope) => {
    if (error && !error.requiredExtensionScope) {
      error.requiredExtensionScope = scope;
    }
    return error;
  };

  const runProvisioningStep = async (label, work) => {
    setStatus(label);
    try {
      return await work();
    } catch (error) {
      if (error && !error.provisioningStep) {
        error.provisioningStep = label;
      }
      throw error;
    }
  };

  const sanitizePipelineNameSegment = (segment, fallback, { lowercase = true } = {}) => {
    const fallbackValue = fallback?.toString() || '';
    const value = segment?.toString().trim();
    const base = value || fallbackValue;
    const normalized = lowercase ? base.toLowerCase() : base;

    const cleaned = normalized
      .replace(/[\\/]+/g, '-')
      .replace(/[^\w.-]+/g, '-')
      .replace(/-+/g, '-')
      .replace(/^-+|-+$/g, '');

    return cleaned || (lowercase ? fallbackValue.toLowerCase() : fallbackValue) || 'segment';
  };

  const buildLegacyPipelineFilename = ({ projectName, repositoryName, branchName }) => {
    const projectSegment = sanitizePipelineNameSegment(projectName, 'project');
    const repoSegment = sanitizePipelineNameSegment(repositoryName || projectName, 'repo');
    const branchSegment = sanitizePipelineNameSegment(branchName?.replace(/^refs\/heads\//, ''), 'branch');
    return `${projectSegment}-${repoSegment}-${branchSegment}.yml`;
  };

  const buildLegacyEnvironmentFirstPipelineFilename = ({ projectName, repositoryName, environment, branchName }) => {
    const projectSegment = sanitizePipelineNameSegment(projectName, 'project');
    const repoSegment = sanitizePipelineNameSegment(repositoryName || projectName, 'repo');
    const environmentSegment = sanitizePipelineNameSegment(environment, 'environment');
    const branchSegment = sanitizePipelineNameSegment(branchName?.replace(/^refs\/heads\//, ''), 'branch');
    return `${projectSegment}-${repoSegment}-${environmentSegment}-${branchSegment}.yml`;
  };

  const buildPipelineFilename = ({ projectName, repositoryName, environment, branchName }) => {
    const projectSegment = sanitizePipelineNameSegment(projectName, 'project');
    const repoSegment = sanitizePipelineNameSegment(repositoryName || projectName, 'repo');
    if (!String(environment || '').trim()) {
      throw new Error('Environment is required to build the Pipeline filename.');
    }
    const environmentSegment = sanitizePipelineNameSegment(environment, 'environment').toUpperCase();
    const branchSegment = sanitizePipelineNameSegment(
      branchName?.replace(/^refs\/heads\//, ''),
      'branch',
      { lowercase: false }
    ).replace(/(^|[-_.])([a-z])/g, (_, separator, character) => `${separator}${character.toUpperCase()}`);
    return `${projectSegment}-${repoSegment}-${branchSegment}To${environmentSegment}.yml`;
  };

  const buildPipelineName = (pipelineFilename) => pipelineFilename;

  const buildReleaseName = ({ service, environment }) => {
    const normalizePart = (value, label) => {
      const normalized = String(value || '').trim().replace(/\s+/g, ' ').toUpperCase();
      if (!normalized || normalized.length > 100 || /[\0\r\n]/.test(normalized)) {
        throw new Error(`${label} is required to build the classic Release name.`);
      }
      return normalized;
    };
    return `${normalizePart(service, 'Service name')} ${normalizePart(environment, 'Environment')}`;
  };

  const getProjectRouteSegment = () => {
    const candidate = state.rawProjectName || state.projectName || state.projectId;
    return candidate ? encodeURIComponent(candidate) : '';
  };

  const buildRepositoryFileUrl = ({ repositoryName, filePath }) => {
    const projectRoute = getProjectRouteSegment();
    return `${state.hostUri}${projectRoute}/_git/${encodeURIComponent(repositoryName)}?path=${encodeURIComponent(
      filePath
    )}&version=GB${encodeURIComponent(SCAFFOLD_BRANCH)}&_a=contents`;
  };

  const showCompletionLinks = ({ supportRepositories, pipelineDefinition }) => {
    const nginx = supportRepositories.find((result) => result.kind === 'nginx');
    const docker = supportRepositories.find((result) => result.kind === 'docker');
    if (!nginx?.repo?.name || !docker?.repo?.name || !pipelineDefinition?.id) {
      throw new Error('Provisioning completed but review links could not be constructed.');
    }
    if (nginxResultLink) {
      nginxResultLink.href = buildRepositoryFileUrl({
        repositoryName: nginx.repo.name,
        filePath: nginx.filePath
      });
      nginxResultLink.textContent = `Review ${nginx.filePath}`;
    }
    if (composeResultLink) {
      composeResultLink.href = buildRepositoryFileUrl({
        repositoryName: docker.repo.name,
        filePath: docker.filePath
      });
      composeResultLink.textContent = `Review ${docker.filePath}`;
    }
    if (pipelineResultLink) {
      const projectRoute = getProjectRouteSegment();
      pipelineResultLink.href = `${state.hostUri}${projectRoute}/_build?definitionId=${encodeURIComponent(
        pipelineDefinition.id
      )}`;
      pipelineResultLink.textContent = `Open Pipeline ${pipelineDefinition.name || pipelineDefinition.id}`;
    }
    finishProvisioning();
  };

  const isUnauthorizedError = (error) =>
    error?.status === 401 ||
    error?.status === 403 ||
    /TF400813/i.test(error?.detail || '') ||
    /\b401\b/.test(error?.message || '');

  const buildSignOutUrl = (hostUri) => `${normalizeHostUri(hostUri || state.hostUri || getHostBase())}_signout`;

  const buildExtensionManagementUrl = (hostUri) =>
    `${normalizeHostUri(hostUri || state.hostUri || getHostBase())}_settings/extensions?tab=installed`;

  const navigateHost = async (url) => {
    const sdk = state.sdk || normalizeSdk(window.VSS || window.parent?.VSS);
    const navigationServiceId = sdk?.ServiceIds?.Navigation;
    if (sdk?.getService && navigationServiceId) {
      try {
        const navigationService = await sdk.getService(navigationServiceId);
        if (navigationService?.navigate) {
          navigationService.navigate(url);
          return;
        }
      } catch (error) {
        console.warn('Azure DevOps host navigation service could not open the sign-out page', error);
      }
    }

    try {
      window.top.location.assign(url);
    } catch (error) {
      console.warn('Top-level navigation was unavailable; opening sign-out in the current window', error);
      window.location.assign(url);
    }
  };

  const restartAzureDevOpsSession = async () => {
    if (!state.hostUri) {
      setStatus('Azure DevOps host context is unavailable. Reopen the generator from the target branch.', true);
      return;
    }
    state.accessToken = null;
    state.accessTokenError = null;
    if (reauthenticateButton) {
      reauthenticateButton.disabled = true;
    }
    setStatus('Signing out of Azure DevOps. Complete the login flow, then reopen Generate pipeline from the branch.');
    await navigateHost(buildSignOutUrl(state.hostUri));
  };

  const openExtensionAuthorization = async () => {
    if (!state.hostUri) {
      setStatus('Azure DevOps host context is unavailable. Reopen the generator from the target branch.', true);
      return;
    }
    state.accessToken = null;
    state.accessTokenError = null;
    if (authorizeExtensionButton) {
      authorizeExtensionButton.disabled = true;
    }
    setStatus('Opening Collection Settings → Extensions. Authorize Pipeline Generator, then reopen it from the branch.');
    await navigateHost(buildExtensionManagementUrl(state.hostUri));
  };

  const applyBootstrapPayload = async (payload = {}, source = 'message') => {
    const {
      branch,
      projectId,
      projectName,
      repoId,
      repoName,
      hostUri,
      accessToken,
      accessTokenError
    } = payload;

    const normalizedHost = normalizeHostUri(hostUri || state.hostUri || getHostBase());
    state.sourceBranch = branch || state.sourceBranch;
    state.branch = SCAFFOLD_BRANCH;
    state.projectId = projectId || state.projectId;
    state.rawProjectName = projectName || state.rawProjectName;
    state.projectName = projectName || state.projectName;
    state.repoId = repoId || state.repoId;
    state.rawRepositoryName = repoName || state.rawRepositoryName;
    state.repositoryName = repoName || state.repositoryName;
    state.hostUri = normalizedHost;
    if (accessToken) {
      state.accessToken = accessToken;
      state.accessTokenError = null;
      setReauthenticationVisibility(false);
    } else {
      state.accessTokenError = accessTokenError || state.accessTokenError;
    }

    const targetBranch = state.branch;
    const sourceBranch = state.sourceBranch;
    const branchDescriptor =
      sourceBranch && sourceBranch !== targetBranch
        ? `${targetBranch} (source: ${sourceBranch})`
        : targetBranch;
    branchLabel.textContent = branchDescriptor
      ? `Target branch: ${branchDescriptor}`
      : 'Loading branch context...';
    if (branchInput && targetBranch) {
      branchInput.value = targetBranch;
      branchInput.disabled = true;
    }

    targetRepoInput.value = `${state.projectName || 'project'}_Azure_DevOps`;
    setServiceNameFromRepository(state.repositoryName || state.projectName, state.projectName);

    if (!state.projectId || !state.accessToken || !state.hostUri) {
      let authMessage;
      if (state.accessTokenError) {
        const needsHostAuth = isHostAuthorizationError(state.accessTokenError);
        authMessage = needsHostAuth
          ? 'Azure DevOps could not issue an access token because extension authorization is missing. A Collection Administrator must authorize Pipeline Generator in Collection Settings → Extensions; if no authorization action is available, reinstall this same published version.'
          : `Azure DevOps did not provide an access token (${state.accessTokenError}). Refresh the page or sign in again, then relaunch the generator.`;
      } else {
        authMessage = 'Loaded context from branch action but still waiting for an access token from Azure DevOps. Refresh or try again if this persists.';
      }
      setReauthenticationVisibility(
        true,
        buildTokenRecoveryMessage(authMessage)
      );
      setStatus(authMessage, true);
      setSubmitting(false);
      return;
    }

    try {
      await loadDeploymentTargets({
        hostUri: state.hostUri,
        accessToken: state.accessToken,
        branch: sourceBranch || targetBranch
      });
      await Promise.all([
        loadPools({ hostUri: state.hostUri, projectId: state.projectId, accessToken: state.accessToken }),
        loadContainerRegistries({ hostUri: state.hostUri, projectId: state.projectId, accessToken: state.accessToken }),
        refreshDockerfiles({
          hostUri: state.hostUri,
          projectId: state.projectId,
          repoId: state.repoId,
          branch: sourceBranch || targetBranch,
          accessToken: state.accessToken
        })
      ]);
      setStatus(
        source === 'message'
          ? 'Azure DevOps context received from the branch action. Generate the pipeline when ready.'
          : 'Azure DevOps context ready. Generate the pipeline when you are ready.'
      );
    } catch (error) {
      console.error('Failed to hydrate form from bootstrap payload', error);
      setStatus('Context loaded, but some resources could not be auto-detected. Fill missing values manually.', true);
    } finally {
      setSubmitting(false);
    }
  };

  const normalizeName = (value) => value?.toString().trim().toLowerCase();

  const extractRepositoryName = (value) => {
    if (!value) return '';
    const segments = value.split('/').filter(Boolean);
    return segments.length ? segments[segments.length - 1] : value;
  };

  const setServiceNameFromRepository = (name, projectName) => {
    if (!serviceInput) return;
    const targetName = extractRepositoryName(name) || projectName;
    const compactProject = String(projectName || '').replace(/\s+/g, '');
    const projectPrefixes = [String(projectName || '').trim(), compactProject]
      .filter(Boolean)
      .sort((left, right) => right.length - left.length);
    let serviceName = String(targetName || '').trim();
    for (const prefix of projectPrefixes) {
      if (serviceName.toLowerCase().startsWith(prefix.toLowerCase())) {
        const remainder = serviceName.slice(prefix.length);
        if (/^[\s._-]+/.test(remainder)) {
          serviceName = remainder.replace(/^[\s._-]+/, '');
          break;
        }
      }
    }
    const normalizedTarget = normalizeName(serviceName || targetName);
    if (!normalizedTarget) return;

    const currentValue = normalizeName(serviceInput.value);
    const projectDefault = normalizeName(projectName);
    const wasAutoFilled = serviceInput.dataset.autofilled === 'true';
    const shouldUpdate =
      !currentValue ||
      wasAutoFilled ||
      (projectDefault && currentValue === projectDefault);

    if (shouldUpdate && currentValue !== normalizedTarget) {
      serviceInput.value = normalizedTarget;
      serviceInput.dataset.autofilled = 'true';
    }
  };

  const setKomodoServerFromEnvironment = (environment) => {
    if (!environment || !komodoSelect) return;
    const normalizedEnvironment = environment.toLowerCase().replace(/[^a-z0-9]/g, '');
    const aliases = {
      dev: ['dev', 'development'],
      pro: ['pro', 'prod', 'production'],
      prod: ['prod', 'production'],
      qa: ['qa'],
      demo: ['demo'],
      soc: ['soc']
    };
    const candidates = aliases[normalizedEnvironment] || [normalizedEnvironment];
    const match = Array.from(komodoSelect.options).find((option) => {
      const normalizedServer = option.value.toLowerCase().replace(/[^a-z0-9]/g, '');
      return candidates.some((candidate) => normalizedServer.startsWith(candidate));
    });
    komodoSelect.value = match ? match.value : '';
  };

  const populateDefaults = () => {
    Object.entries(defaultValues).forEach(([key, value]) => {
      const input = document.getElementById(key);
      if (input && !input.value) {
        if (input.tagName.toLowerCase() === 'select') {
          const hasOption = Array.from(input.options).some((option) => option.value === value);
          if (hasOption) {
            input.value = value;
          }
        } else {
          input.value = value;
        }
      }
    });
  };

  const detectEnvironmentFromBranch = (branch) => {
    if (!branch) return undefined;
    const lower = branch.toLowerCase();

    if (lower.includes('master') || lower.includes('main')) {
      return 'pro';
    }

    const candidates = environmentSelect
      ? Array.from(environmentSelect.options).map((option) => option.value.toLowerCase())
      : [];

    return candidates.find((key) => key && lower.includes(key));
  };

  const applyDetectedEnvironment = (branch) => {
    const detected = detectEnvironmentFromBranch(branch);
    if (detected && environmentSelect) {
      const available = Array.from(environmentSelect.options).some(
        (option) => option.value.toLowerCase() === detected.toLowerCase()
      );
      if (available) {
        environmentSelect.value = detected;
        setKomodoServerFromEnvironment(detected);
      }
    }
  };

  const populateSelectOptions = (select, options, placeholder) => {
    if (!select) return;
    select.innerHTML = '';
    if (placeholder) {
      const hint = document.createElement('option');
      hint.value = '';
      hint.textContent = placeholder;
      hint.disabled = !options.length;
      hint.selected = !options.length;
      select.appendChild(hint);
    }
    options.forEach((option) => {
      const node = document.createElement('option');
      node.value = option.value;
      node.textContent = option.label;
      select.appendChild(node);
    });
    if (!select.value && options.length) {
      select.value = options[0].value;
    }
  };

  const getBranchObjectId = async ({ hostUri, projectId, repoId, branch, accessToken }) => {
    const branchName = SCAFFOLD_BRANCH;
    const refUrl = `${hostUri}${encodeURIComponent(projectId)}/_apis/git/repositories/${repoId}/refs?filter=${encodeURIComponent(
      `heads/${branchName}`
    )}&api-version=6.0`;
    const res = await fetch(refUrl, { headers: authHeaders(accessToken) });

    if (res.status === 404) {
      return ZERO_OBJECT_ID;
    }

    if (!res.ok) {
      const detail = await readErrorDetail(res);
      throw buildHttpError('Failed to query branch', res, detail);
    }

    const payload = await res.json();
    return payload.value?.[0]?.objectId || ZERO_OBJECT_ID;
  };

  const getRepositoryFileContent = async ({ hostUri, projectId, repoId, branchName, path, accessToken }) => {
    const normalizedPath = path.startsWith('/') ? path : `/${path}`;
    const url = `${hostUri}${encodeURIComponent(projectId)}/_apis/git/repositories/${repoId}/items?path=${encodeURIComponent(
      normalizedPath
    )}&versionDescriptor.version=${encodeURIComponent(
      branchName
    )}&versionDescriptor.versionType=branch&%24format=text&api-version=6.0`;
    const res = await fetch(url, { headers: authHeaders(accessToken) });

    if (res.status === 404) {
      return null;
    }

    if (!res.ok) {
      const detail = await readErrorDetail(res);
      throw buildHttpError(`Failed to read repository file ${normalizedPath}`, res, detail);
    }

    return res.text();
  };

  const ensureRepositoryByName = async ({ hostUri, projectId, repositoryName, accessToken }) => {
    const url = `${hostUri}${encodeURIComponent(projectId)}/_apis/git/repositories?api-version=6.0`;

    const res = await fetch(url, {
      headers: authHeaders(accessToken)
    });
    if (!res.ok) {
      const detail = await readErrorDetail(res);
      throw buildHttpError('Failed to list repositories', res, detail);
    }
    const payload = await res.json();
    const existing = (payload.value || []).find((repo) => repo.name === repositoryName);
    if (existing) {
      return existing;
    }

    const createRes = await fetch(url, {
      method: 'POST',
      headers: {
        ...authHeaders(accessToken),
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({ name: repositoryName, project: { id: projectId } })
    });
    if (!createRes.ok) {
      const detail = await readErrorDetail(createRes);
      throw buildHttpError('Failed to create repository', createRes, detail);
    }
    const created = await createRes.json();
    if (!created?.id) {
      throw new Error(`Azure DevOps created ${repositoryName} but returned no repository ID.`);
    }
    return created;
  };

  const ensureRepo = async ({ hostUri, projectId, projectName, accessToken }) => {
    const targetName = `${projectName}_Azure_DevOps`;
    targetRepoInput.value = targetName;
    return ensureRepositoryByName({
      hostUri,
      projectId,
      repositoryName: targetName,
      accessToken
    });
  };

  const ensureDefaultBranch = async ({ hostUri, projectId, repoId, branchName, accessToken }) => {
    const defaultBranch = `refs/heads/${branchName}`;
    const url = `${hostUri}${encodeURIComponent(projectId)}/_apis/git/repositories/${repoId}?api-version=6.0`;
    const res = await fetch(url, {
      method: 'PATCH',
      headers: {
        ...authHeaders(accessToken),
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({ defaultBranch })
    });

    if (!res.ok) {
      const detail = await readErrorDetail(res);
      throw buildHttpError('Failed to set default branch', res, detail);
    }

    return res.json();
  };

  const normalizeResourceSegment = (value, label) => {
    const normalized = String(value || '')
      .trim()
      .toLowerCase()
      .replace(/\s+/g, '-')
      .replace(/[^a-z0-9._-]+/g, '-')
      .replace(/[._-]+/g, '-')
      .replace(/^-+|-+$/g, '');
    if (!normalized) {
      throw new Error(`${label} cannot be converted to a safe resource name.`);
    }
    return normalized;
  };

  const isFrontendService = (service) => {
    const normalized = normalizeResourceSegment(service, 'Service name');
    return ['ui', 'front', 'frontend', 'newui'].includes(normalized) || normalized.endsWith('-ui');
  };

  const buildComposeSample = ({ projectKey, serviceKey, environment, repositoryAddress }) => {
    const containerName = `${projectKey}_${serviceKey.replace(/-/g, '_')}_${environment}`;
    const internalPort = isFrontendService(serviceKey) ? 80 : 8080;
    const registry = String(repositoryAddress || defaultValues.repositoryAddress).trim().replace(/\/+$/, '');
    return [
      'services:',
      `  ${containerName}:`,
      `    container_name: ${containerName}`,
      `    image: ${registry}/${projectKey}/${serviceKey}-${environment}:\${IMAGE_TAG:-CHANGE_ME}`,
      '    restart: unless-stopped',
      '    expose:',
      `      - "${internalPort}"`,
      '    networks:',
      '      - nginx-network',
      '',
      'networks:',
      '  nginx-network:',
      '    external: true',
      ''
    ].join('\n');
  };

  const NGINX_MANAGED_ROUTES_START = '    # BEGIN PIPELINE-GENERATOR MANAGED ROUTES';
  const NGINX_MANAGED_ROUTES_END = '    # END PIPELINE-GENERATOR MANAGED ROUTES';

  const buildNginxRouteBlock = ({ projectKey, serviceKey, environment }) => {
    const containerName = `${projectKey}_${serviceKey.replace(/-/g, '_')}_${environment}`;
    const frontend = isFrontendService(serviceKey);
    const location = frontend ? '/' : `/${serviceKey}/`;
    const internalPort = frontend ? 80 : 8080;
    const upstreamDirectives = [
      '        resolver         127.0.0.11         ipv6=off;',
      `        set              $target            ${containerName};`
    ];
    upstreamDirectives.push(
      frontend
        ? `        proxy_pass                          http://$target:${internalPort}/;`
        : `        proxy_pass                          http://$target:${internalPort};`
    );
    return {
      location,
      content: [
        `    # BEGIN PIPELINE-GENERATOR ROUTE ${serviceKey}`,
        `    location ${location} {`,
        ...upstreamDirectives,
        '        proxy_http_version 1.1;',
        '        proxy_set_header Upgrade $http_upgrade;',
        '        proxy_set_header Connection "upgrade";',
        '        proxy_set_header Host $host;',
        '        proxy_set_header X-Real-IP $remote_addr;',
        '        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;',
        '        proxy_set_header X-Forwarded-Proto $scheme;',
        '        proxy_read_timeout 3600s;',
        '        proxy_send_timeout 3600s;',
        '    }',
        `    # END PIPELINE-GENERATOR ROUTE ${serviceKey}`
      ].join('\n')
    };
  };

  const tokenizeNginx = (content) => {
    const tokens = [];
    let index = 0;
    while (index < content.length) {
      const character = content[index];
      if (/\s/.test(character)) {
        index += 1;
        continue;
      }
      if (character === '#') {
        const newline = content.indexOf('\n', index);
        index = newline === -1 ? content.length : newline + 1;
        continue;
      }
      if (character === '"' || character === "'") {
        const quote = character;
        const start = index;
        let value = '';
        index += 1;
        let closed = false;
        while (index < content.length) {
          if (content[index] === '\\' && index + 1 < content.length) {
            value += content[index + 1];
            index += 2;
            continue;
          }
          if (content[index] === quote) {
            index += 1;
            closed = true;
            break;
          }
          value += content[index];
          index += 1;
        }
        if (!closed) {
          throw new Error('Nginx configuration contains an unterminated quoted value.');
        }
        tokens.push({ value, start, end: index });
        continue;
      }
      if ('{};'.includes(character)) {
        tokens.push({ value: character, start: index, end: index + 1 });
        index += 1;
        continue;
      }
      const start = index;
      while (index < content.length && !/[\s{};#]/.test(content[index])) {
        index += 1;
      }
      tokens.push({ value: content.slice(start, index), start, end: index });
    }
    let depth = 0;
    tokens.forEach((token) => {
      if (token.value === '}') depth -= 1;
      if (depth < 0) {
        throw new Error('Nginx configuration has an unmatched closing brace.');
      }
      token.depth = depth;
      if (token.value === '{') depth += 1;
    });
    if (depth !== 0) {
      throw new Error('Nginx configuration has unmatched braces.');
    }
    return tokens;
  };

  const findMatchingBraceToken = (tokens, openIndex) => {
    let nested = 0;
    for (let index = openIndex; index < tokens.length; index += 1) {
      if (tokens[index].value === '{') nested += 1;
      if (tokens[index].value === '}') nested -= 1;
      if (nested === 0) return index;
    }
    return -1;
  };

  const readNginxDirectiveValues = (tokens, directiveIndex, blockDepth) => {
    const values = [];
    for (let index = directiveIndex + 1; index < tokens.length; index += 1) {
      const token = tokens[index];
      if (token.depth !== blockDepth) continue;
      if (token.value === ';') return values;
      if (token.value === '{' || token.value === '}') return [];
      values.push(token.value);
    }
    return [];
  };

  const findNginxHttpsServer = (content, serverName) => {
    const tokens = tokenizeNginx(content);
    const matches = [];
    for (let index = 0; index < tokens.length - 1; index += 1) {
      const token = tokens[index];
      if (token.value.toLowerCase() !== 'server' || tokens[index + 1]?.value !== '{') continue;
      const openIndex = index + 1;
      const closeIndex = findMatchingBraceToken(tokens, openIndex);
      if (closeIndex === -1) {
        throw new Error('Nginx server block has no matching closing brace.');
      }
      const blockDepth = tokens[openIndex].depth + 1;
      const names = [];
      const listens = [];
      const locations = [];
      for (let cursor = openIndex + 1; cursor < closeIndex; cursor += 1) {
        const current = tokens[cursor];
        if (current.depth !== blockDepth) continue;
        const keyword = current.value.toLowerCase();
        if (keyword === 'server_name') {
          names.push(...readNginxDirectiveValues(tokens, cursor, blockDepth));
        } else if (keyword === 'listen') {
          listens.push(...readNginxDirectiveValues(tokens, cursor, blockDepth));
        } else if (keyword === 'location') {
          const values = [];
          for (let valueIndex = cursor + 1; valueIndex < closeIndex; valueIndex += 1) {
            const valueToken = tokens[valueIndex];
            if (valueToken.depth !== blockDepth) continue;
            if (valueToken.value === '{') break;
            if (valueToken.value === ';' || valueToken.value === '}') break;
            values.push(valueToken.value);
          }
          if (values.length && !values[0].startsWith('~')) {
            const location = ['=', '^~'].includes(values[0]) ? values[1] : values[0];
            if (location) locations.push(location);
          }
        }
      }
      const hasServerName = names.some((name) => name.toLowerCase() === serverName.toLowerCase());
      const listensOnHttps = listens.some((listen) => /(?:^|:)443$/.test(listen));
      if (hasServerName && listensOnHttps) {
        matches.push({
          open: tokens[openIndex],
          close: tokens[closeIndex],
          locations
        });
      }
      index = closeIndex;
    }
    if (!matches.length) {
      throw new Error(`Nginx file has no HTTPS server block for ${serverName}.`);
    }
    if (matches.length > 1) {
      throw new Error(`Nginx file has multiple HTTPS server blocks for ${serverName}; merge them manually first.`);
    }
    return matches[0];
  };

  const normalizeNginxManagedRoutes = ({ content, startIndex, endIndex }) => {
    const managedRoutes = content.slice(startIndex, endIndex);
    const routeBlockPattern = () =>
      /^([ \t]*# BEGIN PIPELINE-GENERATOR ROUTE ([^\r\n]+)[ \t]*\r?\n)([\s\S]*?)(^[ \t]*# END PIPELINE-GENERATOR ROUTE \2[ \t]*)(?:\r?\n)?/gm;
    const escapeRegex = (value) => String(value).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    let migratedRoutes = managedRoutes.replace(
      routeBlockPattern(),
      (routeBlock, startMarker, serviceKey, routeBody, endMarker) => {
        const frontend = isFrontendService(serviceKey);
        const canonicalLocation = frontend ? '/' : `/${serviceKey}/`;
        let migratedBody = routeBody;

        if (!frontend) {
          const legacyLocation = `/${serviceKey}`;
          const legacyLocationPattern = new RegExp(
            `^([ \\t]*location[ \\t]+)${escapeRegex(legacyLocation)}([ \\t]*\\{[ \\t]*$)`,
            'm'
          );
          migratedBody = migratedBody.replace(
            legacyLocationPattern,
            (_, prefix, suffix) => `${prefix}${canonicalLocation}${suffix}`
          );
        }

        migratedBody = migratedBody.replace(
          /^([ \t]*)proxy_pass[ \t]+http:\/\/([A-Za-z0-9][A-Za-z0-9._-]*):([0-9]+)\/?;[ \t]*$/gm,
          (_, indentation, containerName, port) => {
            const suffix = frontend ? '/' : '';
            return [
              `${indentation}resolver         127.0.0.11         ipv6=off;`,
              `${indentation}set              $target            ${containerName};`,
              `${indentation}proxy_pass                          http://$target:${port}${suffix};`
            ].join('\n');
          }
        );

        migratedBody = migratedBody.replace(
          /^([ \t]*)proxy_pass[ \t]+http:\/\/\$target:([0-9]+)\/?;[ \t]*$/gm,
          (_, indentation, port) => {
            const suffix = frontend ? '/' : '';
            return `${indentation}proxy_pass                          http://$target:${port}${suffix};`;
          }
        );

        if (!frontend) {
          const rewriteExpression = `^/${serviceKey}/(.*)$`;
          const generatedRewritePattern = new RegExp(
            `^[ \\t]*rewrite[ \\t]+${escapeRegex(rewriteExpression)}[ \\t]+/\\$1[ \\t]+break;[ \\t]*(?:\\r?\\n)?`,
            'm'
          );
          migratedBody = migratedBody.replace(generatedRewritePattern, '');
        }

        return `${startMarker}${migratedBody}${endMarker}\n`;
      }
    );

    const routeBlocks = [];
    const blockScanner = routeBlockPattern();
    let blockMatch;
    while ((blockMatch = blockScanner.exec(migratedRoutes))) {
      routeBlocks.push({
        start: blockMatch.index,
        end: blockScanner.lastIndex,
        content: blockMatch[0],
        root: /^[ \t]*location[ \t]+\/[ \t]*\{/m.test(blockMatch[0])
      });
    }
    const rootBlock = routeBlocks.find((block) => block.root);
    if (rootBlock && routeBlocks.some((block) => !block.root && block.start > rootBlock.start)) {
      const withoutRoot = `${migratedRoutes.slice(0, rootBlock.start)}${migratedRoutes.slice(rootBlock.end)}`;
      const separator = withoutRoot.endsWith('\n\n') ? '' : withoutRoot.endsWith('\n') ? '\n' : '\n\n';
      migratedRoutes = `${withoutRoot}${separator}${rootBlock.content}`;
    }
    if (migratedRoutes === managedRoutes) return content;
    return `${content.slice(0, startIndex)}${migratedRoutes}${content.slice(endIndex)}`;
  };

  const findManagedRootRouteIndex = ({ content, startIndex, endIndex }) => {
    const managedRoutes = content.slice(startIndex, endIndex);
    const routePattern =
      /^([ \t]*# BEGIN PIPELINE-GENERATOR ROUTE ([^\r\n]+)[ \t]*\r?\n)([\s\S]*?)(^[ \t]*# END PIPELINE-GENERATOR ROUTE \2[ \t]*)(?:\r?\n)?/gm;
    let match;
    while ((match = routePattern.exec(managedRoutes))) {
      if (/^[ \t]*location[ \t]+\/[ \t]*\{/m.test(match[0])) {
        return startIndex + match.index;
      }
    }
    return -1;
  };

  const mergeNginxServiceRoute = ({ content, serverName, projectKey, serviceKey, environment }) => {
    let mergedContent = content;
    let server = findNginxHttpsServer(mergedContent, serverName);
    const route = buildNginxRouteBlock({ projectKey, serviceKey, environment });

    let startIndex = mergedContent.indexOf(NGINX_MANAGED_ROUTES_START, server.open.end);
    let endIndex = mergedContent.indexOf(NGINX_MANAGED_ROUTES_END, server.open.end);
    const startInsideServer = startIndex !== -1 && startIndex < server.close.start;
    const endInsideServer = endIndex !== -1 && endIndex < server.close.start;
    if (startInsideServer !== endInsideServer || (startInsideServer && startIndex > endIndex)) {
      throw new Error('Nginx managed-route markers are incomplete or out of order.');
    }

    if (startInsideServer && endInsideServer) {
      mergedContent = normalizeNginxManagedRoutes({
        content: mergedContent,
        startIndex,
        endIndex
      });
      if (mergedContent !== content) {
        server = findNginxHttpsServer(mergedContent, serverName);
        startIndex = mergedContent.indexOf(NGINX_MANAGED_ROUTES_START, server.open.end);
        endIndex = mergedContent.indexOf(NGINX_MANAGED_ROUTES_END, server.open.end);
      }
    }

    if (server.locations.includes(route.location)) {
      return mergedContent;
    }

    if (startInsideServer && endInsideServer) {
      const rootRouteIndex =
        route.location === '/' ? -1 : findManagedRootRouteIndex({ content: mergedContent, startIndex, endIndex });
      const insertionIndex =
        rootRouteIndex === -1 ? mergedContent.lastIndexOf('\n', endIndex) + 1 : rootRouteIndex;
      return `${mergedContent.slice(0, insertionIndex)}${route.content}\n\n${mergedContent.slice(insertionIndex)}`;
    }

    const beforeClose = mergedContent.slice(0, server.close.start);
    const separator = beforeClose.endsWith('\n') ? '\n' : '\n\n';
    const managedBlock = [
      NGINX_MANAGED_ROUTES_START,
      route.content,
      NGINX_MANAGED_ROUTES_END,
      ''
    ].join('\n');
    return `${beforeClose}${separator}${managedBlock}${mergedContent.slice(server.close.start)}`;
  };

  const buildNginxSample = ({ projectHost, projectKey, serviceKey, environment, domain }) => {
    const route = buildNginxRouteBlock({ projectKey, serviceKey, environment });
    const certificateName = domain.split('.')[0];
    const serverName = `${projectHost}.${domain}`;
    return [
      'server {',
      '    listen 80;',
      `    server_name ${serverName};`,
      '    return 301 https://$host$request_uri;',
      '}',
      '',
      'server {',
      '    listen 443 ssl;',
      `    server_name ${serverName};`,
      '',
      '    client_max_body_size 0;',
      `    ssl_certificate /etc/nginx/conf.d/${certificateName}.pem;`,
      `    ssl_certificate_key /etc/nginx/conf.d/${certificateName}.key;`,
      '',
      NGINX_MANAGED_ROUTES_START,
      route.content,
      NGINX_MANAGED_ROUTES_END,
      '}',
      ''
    ].join('\n');
  };

  const buildSupportRepositorySpecs = ({
    projectName,
    environment,
    domain,
    service,
    repositoryAddress
  }) => {
    const compactProject = String(projectName || '').replace(/\s+/g, '');
    const normalizedEnvironment = normalizeResourceSegment(environment, 'Environment');
    const serviceKey = normalizeResourceSegment(service, 'Service name');
    if (!compactProject || /[\\/\0\r\n]/.test(compactProject)) {
      throw new Error('Project name cannot be converted to a safe DevOps repository name.');
    }
    if (!/^(?=.{1,253}$)(?:[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?\.)+[a-z]{2,63}$/i.test(domain || '')) {
      throw new Error(`Environment ${environment || '(empty)'} has no valid domain.`);
    }
    const compactProjectLower = compactProject.toLowerCase();
    const projectHost = normalizeResourceSegment(compactProjectLower, 'Project hostname');
    const composeDirectory = `${normalizedEnvironment}_${compactProjectLower}`;
    const nginxDirectory = normalizedEnvironment;
    return [
      {
        kind: 'docker',
        name: `${compactProject}_Docker_DevOps`,
        directory: composeDirectory,
        filePath: `/${composeDirectory}/compose.yml`,
        content: buildComposeSample({
          projectKey: compactProjectLower,
          serviceKey,
          environment: normalizedEnvironment,
          repositoryAddress
        })
      },
      {
        kind: 'nginx',
        name: `${compactProject}_Nginx_DevOps`,
        directory: nginxDirectory,
        filePath: `/${nginxDirectory}/${projectHost}-${normalizedEnvironment}.conf`,
        content: buildNginxSample({
          projectHost,
          projectKey: compactProjectLower,
          serviceKey,
          environment: normalizedEnvironment,
          domain: String(domain).toLowerCase()
        }),
        mergeExisting: (content) => mergeNginxServiceRoute({
          content,
          serverName: `${projectHost}.${String(domain).toLowerCase()}`,
          projectKey: compactProjectLower,
          serviceKey,
          environment: normalizedEnvironment
        })
      }
    ];
  };

  const ensureRepositoryBootstrapFiles = async ({
    hostUri,
    projectId,
    repo,
    directory,
    sampleFile,
    accessToken
  }) => {
    const branchName = SCAFFOLD_BRANCH;
    const branchRef = `refs/heads/${branchName}`;
    const oldObjectId = await getBranchObjectId({
      hostUri,
      projectId,
      repoId: repo.id,
      branch: branchName,
      accessToken
    });
    const environmentFile = { path: '/environments', content: 'mattermost_channel=changeme' };
    const desiredFiles = [environmentFile, sampleFile];
    let existingFiles = desiredFiles.map(() => null);
    if (oldObjectId !== ZERO_OBJECT_ID) {
      existingFiles = await Promise.all(
        desiredFiles.map((file) => getRepositoryFileContent({
          hostUri,
          projectId,
          repoId: repo.id,
          branchName,
          path: file.path,
          accessToken
        }))
      );
    }
    const changes = [];
    desiredFiles.forEach((file, index) => {
      const existing = existingFiles[index];
      if (existing === null) {
        changes.push({ ...file, changeType: 'add' });
        return;
      }
      if (typeof file.mergeExisting === 'function') {
        const merged = file.mergeExisting(existing);
        if (merged !== existing) {
          changes.push({ ...file, content: merged, changeType: 'edit' });
        }
      }
    });
    if (!changes.length) {
      return { skipped: true, unchanged: true };
    }

    const url = `${hostUri}${encodeURIComponent(projectId)}/_apis/git/repositories/${repo.id}/pushes?api-version=6.0`;
    const body = {
      refUpdates: [{ name: branchRef, oldObjectId }],
      commits: [
        {
          comment: `Initialize ${repo.name} for ${directory}`,
          changes: changes.map((file) => ({
            changeType: file.changeType,
            item: { path: file.path },
            newContent: { content: file.content, contentType: 'rawtext' }
          }))
        }
      ]
    };
    const res = await fetch(url, {
      method: 'POST',
      headers: {
        ...authHeaders(accessToken),
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(body)
    });
    if (!res.ok) {
      const detail = await readErrorDetail(res);
      throw buildHttpError(`Failed to initialize repository ${repo.name}`, res, detail);
    }
    return { skipped: false, paths: changes.map((file) => file.path) };
  };

  const ensureSupportRepositories = async ({
    hostUri,
    projectId,
    projectName,
    environment,
    domain,
    service,
    repositoryAddress,
    accessToken
  }) => {
    const specs = buildSupportRepositorySpecs({
      projectName,
      environment,
      domain,
      service,
      repositoryAddress
    });
    const results = [];
    for (const spec of specs) {
      const repo = await ensureRepositoryByName({
        hostUri,
        projectId,
        repositoryName: spec.name,
        accessToken
      });
      const bootstrap = await ensureRepositoryBootstrapFiles({
        hostUri,
        projectId,
        repo,
        directory: spec.directory,
        sampleFile: {
          path: spec.filePath,
          content: spec.content,
          mergeExisting: spec.mergeExisting
        },
        accessToken
      });
      await ensureDefaultBranch({
        hostUri,
        projectId,
        repoId: repo.id,
        branchName: SCAFFOLD_BRANCH,
        accessToken
      });
      results.push({
        kind: spec.kind,
        repo,
        directory: spec.directory,
        filePath: spec.filePath,
        bootstrap
      });
    }
    return results;
  };

  const postScaffold = async ({
    hostUri,
    projectId,
    repoId,
    branch,
    accessToken,
    content,
    pipelineFilename = 'project-repo-environment-branch.yml'
  }) => {
    const branchName = SCAFFOLD_BRANCH;
    const branchRef = `refs/heads/${branchName}`;
    const url = `${hostUri}${encodeURIComponent(projectId)}/_apis/git/repositories/${repoId}/pushes?api-version=6.0`;
    const pipelineContent = content || '';
    const oldObjectId = await getBranchObjectId({ hostUri, projectId, repoId, branch: branchName, accessToken });
    const filePath = `/${pipelineFilename}`;
    const existingContent =
      oldObjectId === ZERO_OBJECT_ID
        ? null
        : await getRepositoryFileContent({
          hostUri,
          projectId,
          repoId,
        branchName,
        path: filePath,
          accessToken
        });
    const fileExists = existingContent !== null;
    if (fileExists && existingContent === pipelineContent) {
      return { skipped: true, unchanged: true, path: filePath, branch: branchRef };
    }
    const body = {
      refUpdates: [
        {
          name: branchRef,
          oldObjectId
        }
      ],
      commits: [
        {
          comment: `${fileExists ? 'Update' : 'Add'} pipeline generator defaults`,
          changes: [
            {
              changeType: fileExists ? 'edit' : 'add',
              item: { path: filePath },
              newContent: { content: pipelineContent, contentType: 'rawtext' }
            }
          ]
        }
      ]
    };

    const res = await fetch(url, {
      method: 'POST',
      headers: {
        ...authHeaders(accessToken),
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(body)
    });

    if (!res.ok) {
      const detail = await readErrorDetail(res);
      const error = new Error(`Failed to push scaffold (${res.status})${detail ? `: ${detail}` : ''}`);
      error.status = res.status;
      error.detail = detail;
      throw error;
    }
  };

  const buildPipelineConfiguration = ({ repoId, repositoryName, pipelinePath, branch }) => ({
    type: 'yaml',
    path: pipelinePath.startsWith('/') ? pipelinePath : `/${pipelinePath}`,
    repository: {
      id: repoId,
      name: repositoryName,
      type: 'azureReposGit',
      defaultBranch: `refs/heads/${branch}`
    }
  });

  const buildPipelinesApiUrl = ({ hostUri, projectId, pipelineId, repositoryId }) => {
    const pipelineSegment = pipelineId ? `/${encodeURIComponent(pipelineId)}` : '';
    const searchParams = new URLSearchParams();
    if (repositoryId) {
      searchParams.set('repositoryId', repositoryId);
    }
    searchParams.set('api-version', PIPELINE_API_VERSION);
    return `${hostUri}${encodeURIComponent(projectId)}/_apis/pipelines${pipelineSegment}?${searchParams.toString()}`;
  };

  const readPipelineResponse = async (res, failureMessage) => {
    if (!res.ok) {
      const detail = await readErrorDetail(res);
      throw markErrorDomain(buildHttpError(failureMessage, res, detail), 'pipeline');
    }

    const pipeline = await res.json();
    if (!pipeline?.id) {
      throw markErrorDomain(
        new Error(`${failureMessage.replace(/^Failed to /, 'Azure DevOps reported success for ')} but returned no pipeline ID.`),
        'pipeline'
      );
    }
    return pipeline;
  };

  const getPipelineByName = async ({ hostUri, projectId, pipelineName, legacyPipelineNames = [], accessToken }) => {
    const url = buildPipelinesApiUrl({ hostUri, projectId });
    const res = await fetch(url, { headers: authHeaders(accessToken) });
    if (!res.ok) {
      const detail = await readErrorDetail(res);
      throw markErrorDomain(buildHttpError('Failed to list Azure Pipelines', res, detail), 'pipeline');
    }
    const payload = await res.json();
    const pipelines = payload.value || [];
    const desired = pipelines.find((pipeline) => pipeline.name === pipelineName);
    if (desired) return desired;
    for (const legacyPipelineName of legacyPipelineNames) {
      if (legacyPipelineName && legacyPipelineName !== pipelineName) {
        const legacy = pipelines.find((pipeline) => pipeline.name === legacyPipelineName);
        if (legacy) return legacy;
      }
    }
    return undefined;
  };

  const normalizeComparableFolder = (folder) => normalizePipelineFolder(folder, '\\').toLowerCase();

  const normalizeComparableYamlPath = (path = '') => {
    const normalized = String(path || '').trim().replace(/\\/g, '/');
    return normalized.startsWith('/') ? normalized : `/${normalized}`;
  };

  const buildDefinitionsApiUrl = ({ hostUri, projectId, definitionId, query = {} }) => {
    const definitionSegment = definitionId ? `/${encodeURIComponent(definitionId)}` : '';
    const searchParams = new URLSearchParams(query);
    searchParams.set('api-version', BUILD_API_VERSION);
    return `${hostUri}${encodeURIComponent(projectId)}/_apis/build/definitions${definitionSegment}?${searchParams.toString()}`;
  };

  const readBuildDefinitionResponse = async (res, failureMessage) => {
    if (!res.ok) {
      const detail = await readErrorDetail(res);
      throw markErrorDomain(buildHttpError(failureMessage, res, detail), 'pipeline');
    }

    const definition = await res.json();
    if (!definition?.id) {
      throw markErrorDomain(new Error(`${failureMessage} but Azure DevOps returned no Build Definition ID.`), 'pipeline');
    }
    return definition;
  };

  const getBuildDefinitionById = async ({ hostUri, projectId, definitionId, accessToken }) => {
    const url = buildDefinitionsApiUrl({ hostUri, projectId, definitionId });
    const res = await fetch(url, { headers: authHeaders(accessToken) });
    return readBuildDefinitionResponse(res, `Failed to load Build Definition ${definitionId}`);
  };

  const getBuildDefinitionByYamlPath = async ({
    hostUri,
    projectId,
    repoId,
    pipelinePath,
    accessToken
  }) => {
    const desiredPath = normalizeComparableYamlPath(pipelinePath);
    const url = buildDefinitionsApiUrl({
      hostUri,
      projectId,
      query: {
        repositoryId: repoId,
        repositoryType: 'TfsGit',
        yamlFilename: desiredPath,
        includeAllProperties: 'true'
      }
    });
    const res = await fetch(url, { headers: authHeaders(accessToken) });
    if (!res.ok) {
      const detail = await readErrorDetail(res);
      throw markErrorDomain(buildHttpError('Failed to find a Pipeline by its YAML path', res, detail), 'pipeline');
    }

    const payload = await res.json();
    const references = [...(payload.value || [])].sort((left, right) => Number(left.id) - Number(right.id));
    for (const reference of references) {
      const definition =
        reference?.process?.yamlFilename && reference?.repository?.id
          ? reference
          : await getBuildDefinitionById({
              hostUri,
              projectId,
              definitionId: reference.id,
              accessToken
            });
      const samePath = normalizeComparableYamlPath(definition?.process?.yamlFilename) === desiredPath;
      const sameRepository = String(definition?.repository?.id || '') === String(repoId);
      if (samePath && sameRepository) {
        return definition;
      }
    }
    return undefined;
  };

  const updateBuildDefinition = async ({
    hostUri,
    projectId,
    definition,
    repo,
    pipelineName,
    desiredConfig,
    accessToken
  }) => {
    // Azure DevOps requires the current revision and recommends GET-modify-PUT
    // with the complete Build Definition document. A list response can include
    // a revision without containing every field, so always fetch the full
    // definition immediately before the update.
    const current = await getBuildDefinitionById({
      hostUri,
      projectId,
      definitionId: definition.id,
      accessToken
    });
    const updated = {
      ...current,
      name: pipelineName,
      path: PIPELINE_FOLDER,
      comment: 'Updated by Pipeline Generator.',
      process: {
        ...(current.process || {}),
        type: current.process?.type ?? 2,
        yamlFilename: desiredConfig.path
      },
      repository: {
        ...(current.repository || {}),
        id: desiredConfig.repository.id,
        name: repo.name,
        type: current.repository?.type || 'TfsGit',
        defaultBranch: desiredConfig.repository.defaultBranch
      }
    };
    const url = buildDefinitionsApiUrl({ hostUri, projectId, definitionId: current.id });
    const res = await fetch(url, {
      method: 'PUT',
      headers: {
        ...authHeaders(accessToken),
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(updated)
    });
    return readBuildDefinitionResponse(res, `Failed to update Build Definition ${current.id}`);
  };

  const pipelineBindingMatches = ({ pipeline, pipelineName, desiredConfig }) => {
    const configuration = pipeline?.configuration;
    const repository = configuration?.repository || pipeline?.repository;
    const yamlPath = configuration?.path || pipeline?.process?.yamlFilename;
    const folder = pipeline?.folder || pipeline?.path;
    return (
      pipeline?.name === pipelineName &&
      normalizeComparableFolder(folder) === normalizeComparableFolder(PIPELINE_FOLDER) &&
      normalizeComparableYamlPath(yamlPath) === normalizeComparableYamlPath(desiredConfig.path) &&
      String(repository?.id || '') === String(desiredConfig.repository.id) &&
      repository?.defaultBranch === desiredConfig.repository.defaultBranch
    );
  };

  const upsertPipelineDefinition = async ({
    hostUri,
    projectId,
    repo,
    pipelineName,
    pipelinePath,
    legacyPipelineNames = [],
    legacyPipelinePaths = [],
    branch,
    accessToken
  }) => {
    const repositoryName = `${state.projectName || projectId}/${repo.name}`;
    const desiredConfig = buildPipelineConfiguration({
      repoId: repo.id,
      repositoryName,
      pipelinePath,
      branch
    });

    const existing = await getPipelineByName({
      hostUri,
      projectId,
      pipelineName,
      legacyPipelineNames,
      accessToken
    });
    if (existing?.id) {
      // The Pipelines API's by-ID response can omit repository.defaultBranch
      // on Azure DevOps Server. Treating that sparse response as a mismatch
      // caused a no-op rerun to PUT the Build Definition and increment its
      // revision. Read the canonical full Build Definition before deciding
      // whether migration is necessary.
      const current = await getBuildDefinitionById({
        hostUri,
        projectId,
        definitionId: existing.id,
        accessToken
      });
      if (pipelineBindingMatches({ pipeline: current, pipelineName, desiredConfig })) {
        return current || existing;
      }
      return updateBuildDefinition({
        hostUri,
        projectId,
        definition: existing,
        repo,
        pipelineName,
        desiredConfig,
        accessToken
      });
    }

    const existingForYaml = await getBuildDefinitionByYamlPath({
      hostUri,
      projectId,
      repoId: repo.id,
      pipelinePath: desiredConfig.path,
      accessToken
    });
    if (existingForYaml?.id) {
      const alreadyDesired = pipelineBindingMatches({
        pipeline: existingForYaml,
        pipelineName,
        desiredConfig
      });
      if (alreadyDesired) {
        return existingForYaml;
      }
      return updateBuildDefinition({
        hostUri,
        projectId,
        definition: existingForYaml,
        repo,
        pipelineName,
        desiredConfig,
        accessToken
      });
    }

    for (const legacyPipelinePath of legacyPipelinePaths) {
      if (!legacyPipelinePath || normalizeComparableYamlPath(legacyPipelinePath) === desiredConfig.path) continue;
      const legacyForYaml = await getBuildDefinitionByYamlPath({
        hostUri,
        projectId,
        repoId: repo.id,
        pipelinePath: legacyPipelinePath,
        accessToken
      });
      if (legacyForYaml?.id) {
        return updateBuildDefinition({
          hostUri,
          projectId,
          definition: legacyForYaml,
          repo,
          pipelineName,
          desiredConfig,
          accessToken
        });
      }
    }

    const createUrl = buildPipelinesApiUrl({
      hostUri,
      projectId,
      repositoryId: repo.id
    });
    const res = await fetch(createUrl, {
      method: 'POST',
      headers: {
        ...authHeaders(accessToken),
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({ name: pipelineName, folder: PIPELINE_FOLDER, configuration: desiredConfig })
    });

    return readPipelineResponse(res, 'Failed to create pipeline');
  };

  const buildReleaseDefinitionsApiUrl = ({ hostUri, projectId, definitionId, query = {} }) => {
    const definitionSegment = definitionId ? `/${encodeURIComponent(definitionId)}` : '';
    const searchParams = new URLSearchParams(query);
    searchParams.set('api-version', RELEASE_API_VERSION);
    return `${hostUri}${encodeURIComponent(projectId)}/_apis/release/definitions${definitionSegment}?${searchParams.toString()}`;
  };

  const readReleaseDefinitionResponse = async (res, failureMessage) => {
    if (!res.ok) {
      const detail = await readErrorDetail(res);
      const error = markErrorDomain(buildHttpError(failureMessage, res, detail), 'release');
      error.responseDetail = detail;
      throw error;
    }
    const definition = await res.json();
    if (!definition?.id) {
      throw markErrorDomain(new Error(`${failureMessage} but Azure DevOps returned no Release Definition ID.`), 'release');
    }
    return definition;
  };

  const getReleaseDefinitionByName = async ({ hostUri, projectId, releaseName, accessToken }) => {
    const searchParams = new URLSearchParams({
      searchText: releaseName,
      isExactNameMatch: 'true',
      searchTextContainsFolderName: 'false',
      '$top': '100'
    });
    const url = buildReleaseDefinitionsApiUrl({
      hostUri,
      projectId,
      query: Object.fromEntries(searchParams.entries())
    });
    const res = await fetch(url, { headers: authHeaders(accessToken) });
    if (!res.ok) {
      const detail = await readErrorDetail(res);
      throw markErrorDomain(buildHttpError('Failed to list classic Release definitions', res, detail), 'release');
    }
    const payload = await res.json();
    return (payload.value || []).find((definition) => definition.name === releaseName);
  };

  const getReleaseDefinitionById = async ({ hostUri, projectId, definitionId, accessToken }) => {
    const url = buildReleaseDefinitionsApiUrl({ hostUri, projectId, definitionId });
    const res = await fetch(url, { headers: authHeaders(accessToken) });
    return readReleaseDefinitionResponse(res, `Failed to load classic Release definition ${definitionId}`);
  };

  const getReleaseDefinitionByPipelineId = async ({
    hostUri,
    projectId,
    pipelineId,
    accessToken
  }) => {
    const url = buildReleaseDefinitionsApiUrl({
      hostUri,
      projectId,
      query: {
        '$expand': 'Artifacts',
        artifactType: 'Build',
        artifactSourceId: `${projectId}:${pipelineId}`,
        '$top': '100'
      }
    });
    const res = await fetch(url, { headers: authHeaders(accessToken) });
    if (!res.ok) {
      const detail = await readErrorDetail(res);
      throw markErrorDomain(buildHttpError('Failed to find a Release definition by Pipeline artifact', res, detail), 'release');
    }
    const payload = await res.json();
    return (payload.value || []).find((definition) =>
      (definition.artifacts || []).some(
        (artifact) => String(artifact?.definitionReference?.definition?.id || '') === String(pipelineId)
      )
    );
  };

  const resolveReleaseAgentQueue = async ({ hostUri, projectId, queueName, accessToken }) => {
    if (!queueName) {
      throw markErrorDomain(new Error('No agent queue was selected for the classic Release job.'), 'release');
    }

    const url = `${hostUri}${encodeURIComponent(projectId)}/_apis/distributedtask/queues?api-version=6.0`;
    const res = await fetch(url, { headers: authHeaders(accessToken) });
    if (!res.ok) {
      const detail = await readErrorDetail(res);
      throw markRequiredExtensionScope(
        markErrorDomain(buildHttpError(`Failed to load agent queue ${queueName}`, res, detail), 'release'),
        'vso.agentpools'
      );
    }

    const payload = await res.json();
    const queue = (payload.value || []).find((item) => item.name === queueName);
    if (!queue?.id) {
      throw markErrorDomain(
        new Error(`Release agent queue was not found: ${queueName}. Check Project settings → Agent pools/queues.`),
        'release'
      );
    }
    return queue;
  };

  const resolveReleaseScriptRepository = async ({ hostUri, source, accessToken }) => {
    const scriptProject = source.project || state.projectId;
    const repositoryName = source.repository;
    const rawPath = source.path;
    const branch = source.branch || SCAFFOLD_BRANCH;
    if (!scriptProject || !repositoryName || !rawPath) {
      throw markErrorDomain(
        new Error('Release script source is incomplete. azureReposFile requires project, repository, and path.'),
        'release'
      );
    }

    const repositoriesUrl = `${hostUri}${encodeURIComponent(scriptProject)}/_apis/git/repositories?api-version=6.0`;
    const repositoriesResponse = await fetch(repositoriesUrl, { headers: authHeaders(accessToken) });
    if (!repositoriesResponse.ok) {
      const detail = await readErrorDetail(repositoriesResponse);
      throw markErrorDomain(buildHttpError('Failed to list the configured Release script repository', repositoriesResponse, detail), 'release');
    }

    const repositories = await repositoriesResponse.json();
    const scriptRepository = (repositories.value || []).find(
      (item) => item.id === repositoryName || item.name === repositoryName
    );
    if (!scriptRepository?.id) {
      throw markErrorDomain(
        new Error(`Release script repository was not found: ${scriptProject}/${repositoryName}.`),
        'release'
      );
    }

    const scriptPath = rawPath.startsWith('/') ? rawPath : `/${rawPath}`;
    const fileUrl = `${hostUri}${encodeURIComponent(scriptProject)}/_apis/git/repositories/${encodeURIComponent(
      scriptRepository.id
    )}/items?path=${encodeURIComponent(scriptPath)}&versionDescriptor.version=${encodeURIComponent(
      branch
    )}&versionDescriptor.versionType=branch&%24format=text&api-version=6.0`;
    const fileResponse = await fetch(fileUrl, { headers: authHeaders(accessToken) });
    if (!fileResponse.ok) {
      const detail = await readErrorDetail(fileResponse);
      throw markErrorDomain(buildHttpError(`Failed to load Release Bash script ${scriptPath}`, fileResponse, detail), 'release');
    }

    const script = await fileResponse.text();
    if (!script.trim()) {
      throw markErrorDomain(new Error(`Release Bash script is empty: ${scriptProject}/${repositoryName}${scriptPath}.`), 'release');
    }
    return script;
  };

  const resolveReleaseInlineScript = async ({ releaseConfig, hostUri, accessToken }) => {
    const source = releaseConfig.scriptSource || {};
    if (source.type === 'inline') {
      if (typeof source.content !== 'string' || !source.content.trim()) {
        throw markErrorDomain(
          new Error('Release Bash inline script is empty. Set scriptSource.content in dist/release-config.js.'),
          'release'
        );
      }
      return source.content;
    }
    if (source.type === 'packagedFile') {
      const packagedPath = String(source.path || '').trim();
      if (!packagedPath) {
        throw markErrorDomain(
          new Error('Release Bash packagedFile path is empty in dist/release-config.js.'),
          'release'
        );
      }
      const packagedUrl = new URL(packagedPath, window.location.href).toString();
      const packagedResponse = await fetch(packagedUrl, { cache: 'no-store' });
      if (!packagedResponse.ok) {
        const detail = await readErrorDetail(packagedResponse);
        throw markErrorDomain(
          buildHttpError(`Failed to load packaged Release Bash script ${packagedPath}`, packagedResponse, detail),
          'release'
        );
      }
      const script = await packagedResponse.text();
      if (!script.trim()) {
        throw markErrorDomain(new Error(`Packaged Release Bash script is empty: ${packagedPath}.`), 'release');
      }
      return script;
    }
    if (source.type === 'azureReposFile') {
      return resolveReleaseScriptRepository({ hostUri, source, accessToken });
    }
    throw markErrorDomain(
      new Error(
        `Unsupported Release script source type: ${source.type || 'missing'}. Use inline, packagedFile, or azureReposFile.`
      ),
      'release'
    );
  };

  const resolveReleaseVariableGroup = async ({
    hostUri,
    projectId,
    groupName,
    requiredVariableNames,
    accessToken
  }) => {
    if (!groupName) {
      throw markErrorDomain(new Error('Release variableGroupName is empty in dist/release-config.js.'), 'release');
    }
    const searchParams = new URLSearchParams({
      groupName,
      actionFilter: 'Use',
      'api-version': '7.1'
    });
    const url = `${hostUri}${encodeURIComponent(projectId)}/_apis/distributedtask/variablegroups?${searchParams}`;
    const res = await fetch(url, { headers: authHeaders(accessToken) });
    if (!res.ok) {
      const detail = await readErrorDetail(res);
      throw markRequiredExtensionScope(
        markErrorDomain(buildHttpError(`Failed to load Release variable group ${groupName}`, res, detail), 'release'),
        'vso.variablegroups_read'
      );
    }

    const payload = await res.json();
    const group = (payload.value || [])
      .filter((item) => item?.name === groupName && item?.id != null)
      .sort((left, right) => Number(left.id) - Number(right.id))[0];
    if (!group) {
      throw markErrorDomain(
        new Error(`Release variable group was not found or cannot be used: ${groupName}. Check Pipelines → Library.`),
        'release'
      );
    }

    const variables = group.variables || {};
    const missingVariables = (requiredVariableNames || []).filter(
      (name) => !Object.prototype.hasOwnProperty.call(variables, name)
    );
    if (missingVariables.length) {
      throw markErrorDomain(
        new Error(
          `Release variable group ${groupName} is missing required variables: ${missingVariables.join(', ')}.`
        ),
        'release'
      );
    }
    return { id: Number(group.id), name: group.name };
  };

  const buildReleaseDefinitionPayload = ({
    releaseName,
    releaseConfig,
    inlineScript,
    projectId,
    projectName,
    repo,
    pipelineDefinition,
    pipelineName,
    branch,
    agentQueue,
    variableGroup
  }) => {
    const defaultBranch = `refs/heads/${branch}`;
    const artifactAlias = `_${pipelineName}`;
    return {
      name: releaseName,
      path: releaseConfig.folder,
      description: 'Generated by Pipeline Generator.',
      releaseNameFormat: 'Release-$(rev:r)',
      artifacts: [
        {
          sourceId: `${projectId}:${pipelineDefinition.id}`,
          type: 'Build',
          alias: artifactAlias,
          definitionReference: {
            definition: { id: String(pipelineDefinition.id), name: pipelineName },
            project: { id: projectId, name: projectName },
            repository: { id: repo.id, name: repo.name },
            defaultVersionBranch: { id: defaultBranch, name: defaultBranch },
            defaultVersionType: { id: 'latestType', name: 'Latest' },
            defaultVersionSpecific: { id: '', name: '' },
            defaultVersionTags: { id: '', name: '' },
            artifactSourceDefinitionUrl: { id: '', name: '' }
          },
          isPrimary: true,
          isRetained: false
        }
      ],
      environments: [
        {
          name: releaseConfig.environmentName,
          rank: 1,
          variables: {},
          variableGroups: [],
          demands: [],
          conditions: [{ name: 'ReleaseStarted', conditionType: 'event', value: '', result: null }],
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
              executionOrder: 'beforeGates'
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
              executionOrder: 'afterSuccessfulGates'
            }
          },
          deployPhases: [
            {
              name: 'Agent job',
              phaseType: 'agentBasedDeployment',
              rank: 1,
              workflowTasks: [
                {
                  taskId: BASH_TASK_ID,
                  version: '3.*',
                  name: releaseConfig.bashTaskName,
                  refName: '',
                  enabled: true,
                  alwaysRun: false,
                  continueOnError: false,
                  timeoutInMinutes: 0,
                  definitionType: 'task',
                  condition: 'succeeded()',
                  inputs: {
                    targetType: 'inline',
                    script: inlineScript,
                    workingDirectory: '',
                    failOnStderr: 'false',
                    noProfile: 'true',
                    noRc: 'true'
                  }
                }
              ],
              deploymentInput: {
                queueId: Number(agentQueue.id),
                queueName: agentQueue.name,
                demands: [],
                enableAccessToken: false,
                skipArtifactsDownload: false,
                timeoutInMinutes: 0,
                jobCancelTimeoutInMinutes: 1,
                condition: 'succeeded()',
                overrideInputs: {},
                parallelExecution: { parallelExecutionType: 'none' },
                artifactsDownloadInput: { downloadInputs: [] }
              }
            }
          ]
        }
      ],
      variables: {},
      variableGroups: [Number(variableGroup.id)],
      triggers: [],
      properties: {}
    };
  };

  const getReleaseWorkflowTask = (definition) =>
    definition?.environments?.[0]?.deployPhases?.[0]?.workflowTasks?.[0];

  const getReleaseDeploymentInput = (definition) =>
    definition?.environments?.[0]?.deployPhases?.[0]?.deploymentInput;

  const releaseDefinitionMatches = ({ current, desired }) => {
    const currentArtifact = current?.artifacts?.[0];
    const desiredArtifact = desired?.artifacts?.[0];
    const currentEnvironment = current?.environments?.[0];
    const desiredEnvironment = desired?.environments?.[0];
    const currentTask = getReleaseWorkflowTask(current);
    const desiredTask = getReleaseWorkflowTask(desired);
    const currentDeployment = getReleaseDeploymentInput(current);
    const desiredDeployment = getReleaseDeploymentInput(desired);
    const currentVariableGroupIds = new Set(
      (current?.variableGroups || []).map((groupId) => String(groupId))
    );
    const hasDesiredVariableGroups = (desired?.variableGroups || []).every((groupId) =>
      currentVariableGroupIds.has(String(groupId))
    );
    return (
      current?.name === desired?.name &&
      normalizeComparableFolder(current?.path) === normalizeComparableFolder(desired?.path) &&
      String(currentArtifact?.definitionReference?.definition?.id || '') ===
        String(desiredArtifact?.definitionReference?.definition?.id || '') &&
      String(currentArtifact?.definitionReference?.repository?.id || '') ===
        String(desiredArtifact?.definitionReference?.repository?.id || '') &&
      currentEnvironment?.name === desiredEnvironment?.name &&
      Number(currentDeployment?.queueId) === Number(desiredDeployment?.queueId) &&
      currentTask?.taskId === desiredTask?.taskId &&
      currentTask?.version === desiredTask?.version &&
      currentTask?.name === desiredTask?.name &&
      currentTask?.inputs?.targetType === 'inline' &&
      currentTask?.inputs?.script === desiredTask?.inputs?.script &&
      hasDesiredVariableGroups &&
      currentEnvironment?.conditions?.some((condition) => condition?.name === 'ReleaseStarted') &&
      currentEnvironment?.preDeployApprovals?.approvals?.some((approval) => approval?.isAutomated === true) &&
      currentEnvironment?.postDeployApprovals?.approvals?.some((approval) => approval?.isAutomated === true)
    );
  };

  const updateReleaseDefinition = async ({
    hostUri,
    projectId,
    definition,
    desired,
    accessToken
  }) => {
    const current = await getReleaseDefinitionById({
      hostUri,
      projectId,
      definitionId: definition.id,
      accessToken
    });
    if (releaseDefinitionMatches({ current, desired })) {
      return { ...current, created: false, updated: false };
    }
    const mergedEnvironments = (desired.environments || []).map((desiredEnvironment, index) => {
      const currentEnvironment = current.environments?.[index] || {};
      const mergedDeployPhases = (desiredEnvironment.deployPhases || []).map((desiredPhase, phaseIndex) => ({
        ...(currentEnvironment.deployPhases?.[phaseIndex] || {}),
        ...desiredPhase
      }));
      return {
        ...currentEnvironment,
        ...desiredEnvironment,
        ...(currentEnvironment.id ? { id: currentEnvironment.id } : {}),
        deployPhases: mergedDeployPhases
      };
    });
    const mergedVariableGroups = Array.from(
      new Set(
        [...(current.variableGroups || []), ...(desired.variableGroups || [])]
          .map((groupId) => Number(groupId))
          .filter(Number.isFinite)
      )
    );
    const updated = {
      ...current,
      ...desired,
      id: current.id,
      revision: current.revision,
      comment: 'Updated by Pipeline Generator.',
      environments: mergedEnvironments,
      variableGroups: mergedVariableGroups
    };
    const url = buildReleaseDefinitionsApiUrl({ hostUri, projectId });
    const res = await fetch(url, {
      method: 'PUT',
      headers: {
        ...authHeaders(accessToken),
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(updated)
    });
    const result = await readReleaseDefinitionResponse(
      res,
      `Failed to update classic Release definition ${current.id}`
    );
    return { ...result, created: false, updated: true };
  };

  const ensureReleaseDefinition = async ({
    hostUri,
    projectId,
    projectName,
    repo,
    pipelineDefinition,
    pipelineName,
    service,
    environment,
    branch,
    queueName,
    accessToken
  }) => {
    const releaseConfig = getReleaseConfig();
    if (!releaseConfig.enabled) {
      return { skipped: true, reason: 'Release creation is disabled in dist/release-config.js.' };
    }
    if (!releaseConfig.environmentName) {
      throw markErrorDomain(new Error('Release environmentName is empty in dist/release-config.js.'), 'release');
    }
    if (!releaseConfig.variableGroupName) {
      throw markErrorDomain(new Error('Release variableGroupName is empty in dist/release-config.js.'), 'release');
    }

    const releaseName = buildReleaseName({ service, environment });
    const existingByName = await getReleaseDefinitionByName({ hostUri, projectId, releaseName, accessToken });
    const existingByPipeline = existingByName?.id
      ? undefined
      : await getReleaseDefinitionByPipelineId({
          hostUri,
          projectId,
          pipelineId: pipelineDefinition.id,
          accessToken
        });

    const [inlineScript, agentQueue, variableGroup] = await Promise.all([
      resolveReleaseInlineScript({ releaseConfig, hostUri, accessToken }),
      resolveReleaseAgentQueue({ hostUri, projectId, queueName, accessToken }),
      resolveReleaseVariableGroup({
        hostUri,
        projectId,
        groupName: releaseConfig.variableGroupName,
        requiredVariableNames: releaseConfig.requiredVariableNames,
        accessToken
      })
    ]);
    const body = buildReleaseDefinitionPayload({
      releaseName,
      releaseConfig,
      inlineScript,
      projectId,
      projectName,
      repo,
      pipelineDefinition,
      pipelineName,
      branch,
      agentQueue,
      variableGroup
    });
    const existing = existingByName || existingByPipeline;
    if (existing?.id) {
      return updateReleaseDefinition({
        hostUri,
        projectId,
        definition: existing,
        desired: body,
        accessToken
      });
    }

    const url = buildReleaseDefinitionsApiUrl({ hostUri, projectId });
    const res = await fetch(url, {
      method: 'POST',
      headers: {
        ...authHeaders(accessToken),
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(body)
    });
    if (!res.ok) {
      const detail = await readErrorDetail(res);
      if (res.status === 409 || /already exists|same name|duplicate/i.test(detail)) {
        const duplicate = await getReleaseDefinitionByName({ hostUri, projectId, releaseName, accessToken });
        if (duplicate?.id) {
          return updateReleaseDefinition({
            hostUri,
            projectId,
            definition: duplicate,
            desired: body,
            accessToken
          });
        }
      }
      throw markErrorDomain(buildHttpError(`Failed to create classic Release definition ${releaseName}`, res, detail), 'release');
    }
    return { ...(await res.json()), created: true, updated: false };
  };

  const fetchAgentQueues = async ({ hostUri, projectId, accessToken }) => {
    const url = `${hostUri}${encodeURIComponent(projectId)}/_apis/distributedtask/queues?api-version=6.0`;
    const res = await fetch(url, { headers: authHeaders(accessToken) });
    if (!res.ok) {
      const detail = await readErrorDetail(res);
      throw buildHttpError('Failed to load pools', res, detail);
    }
    const payload = await res.json();
    return Array.from(new Set((payload.value || []).map((queue) => queue.name).filter(Boolean)));
  };

  const fetchContainerRegistries = async ({ hostUri, projectId, accessToken }) => {
    const url = `${hostUri}${encodeURIComponent(projectId)}/_apis/serviceendpoint/endpoints?type=dockerregistry&projectIds=${encodeURIComponent(
      projectId
    )}&api-version=6.0`;
    const res = await fetch(url, { headers: authHeaders(accessToken) });
    if (!res.ok) {
      const detail = await readErrorDetail(res);
      throw buildHttpError('Failed to load container registries', res, detail);
    }
    const payload = await res.json();
    return (payload.value || []).map((endpoint) => endpoint.name || endpoint.id).filter(Boolean);
  };


  const normalizeDockerfileDir = (path = '') => {
    const normalized = path.split('\\').join('/');
    const withoutFile = normalized.replace(/\/?Dockerfile$/i, '');
    const trimmed = withoutFile.replace(/^\/+/, '').replace(/^\//, '');
    return trimmed || '.';
  };

  const buildPipelineYaml = (payload, options = {}) => {
    const sourceBranchName = (options.sourceBranch || 'main').replace(/^refs\/heads\//, '');
    const sourceRepositoryName =
      options.rawRepositoryName || options.repositoryName || options.sourceRepositoryName || 'repository';
    const projectName = options.rawProjectName || options.projectName || 'PROJECTNAME';
    const projectRepoName = `${projectName}/${sourceRepositoryName}`;
    return [
      "trigger: none",
      '',
      'resources:',
      '  repositories:',
      '    - repository: SharedTemplatesRepo',
      '      type: git',
      '      endpoint: ShonizCollection',
      '      name: SharedTemplates/SharedTemplates',
      '      ref: main',
      '',
      `    - repository: otherRepo`,
      '      type: git',
      `      name: "${projectRepoName}"`,
      `      ref: refs/heads/${sourceBranchName}`,
      '      trigger:',
      '        branches:',
      '          include:',
      `            - ${sourceBranchName}`,
      '#        paths:',
      '#          exclude:',
      '#            - server/**',
      '#          include:',
      '#            - client/**',
      '',
      'variables:',
      '- group: KomodoAPI',
      '',
      'stages:',
      '- template: build-push-komodo.yml@SharedTemplatesRepo',
      '  parameters:',
      `    pool: '${payload.pool || ''}'`,
      `    service: '${payload.service || ''}'                # service name`,
      `    environment: '${payload.environment || ''}'           # selected deployment environment`,
      `    dockerfileDir: '${payload.dockerfileDir || '**'}'  # path of Dockerfile, Default is '**'`,
      `    repositoryAddress: '${payload.repositoryAddress || ''}'`,
      `    containerRegistryService: '${payload.containerRegistryService || ''}'`,
      "    tag: '1.0.$(Build.BuildId)'",
      `    komodoServer: '${payload.komodoServer || ''}' # selected Komodo server`,
      "    komodoApiKey: '$(KOMODO_API_KEY)'",
      "    komodoApiSecret: '$(KOMODO_API_SECRET)'",
      '    sourceRepo: otherRepo',
      ''
    ].join('\n');
  };

  const handleSubmit = async (event) => {
    event.preventDefault();
    if (state.provisioningComplete) return;
    setCompletionVisibility(false);
    if (initializationPromise) {
      try {
        await initializationPromise;
      } catch (error) {
        console.error('Initialization failed before submit', error);
      }
    }
    if (!state.deploymentTargetsReady) {
      setStatus(
        `Deployment targets are unavailable. Verify ${DEPLOYMENT_TARGETS_CONFIG.collection}/${DEPLOYMENT_TARGETS_CONFIG.project}/${DEPLOYMENT_TARGETS_CONFIG.repository}:${DEPLOYMENT_TARGETS_CONFIG.path} on ${DEPLOYMENT_TARGETS_CONFIG.branch} and reopen the generator.`,
        true
      );
      setSubmitting(false);
      return;
    }
    const payload = Object.fromEntries(new FormData(form).entries());
    const environmentConfig = state.deploymentTargets?.environmentConfigs?.find(
      ({ name }) => name.toLowerCase() === String(payload.environment || '').toLowerCase()
    );
    if (!environmentConfig?.domain) {
      setStatus(`No domain is configured for environment ${payload.environment || '(empty)'}.`, true);
      setSubmitting(false);
      return;
    }
    const yaml = buildPipelineYaml(payload, {
      sourceBranch: state.sourceBranch,
      rawProjectName: state.rawProjectName,
      projectName: state.projectName,
      rawRepositoryName: state.rawRepositoryName,
      repositoryName: state.repositoryName,
      sourceRepositoryName: state.rawRepositoryName || state.repositoryName || state.projectName
    });

    setStatus('Generating pipeline template...');
    setSubmitting(true);

    if (!state.accessToken) {
      const errorMessage = state.accessTokenError || 'Azure DevOps did not issue an extension access token.';
      const message = buildTokenRecoveryMessage(errorMessage);
      setReauthenticationVisibility(true, message);
      setStatus(message, true);
      setSubmitting(false);
      return;
    }

    if (!state.projectId && state.sdk?.getWebContext) {
      const context = state.sdk.getWebContext();
      state.projectId = context?.project?.id || state.projectId;
      const contextProjectName = context?.project?.name;
      state.rawProjectName = contextProjectName || state.rawProjectName;
      state.projectName = contextProjectName || state.projectId || state.projectName;
      state.repoId = context?.repository?.id || state.repoId;
      const contextRepositoryName = context?.repository?.name;
      state.rawRepositoryName = contextRepositoryName || state.rawRepositoryName;
      state.repositoryName = contextRepositoryName || state.repositoryName;
    }

    const sourceRepositoryName = state.rawRepositoryName || state.repositoryName || state.projectName;
    const pipelineFilename = buildPipelineFilename({
      projectName: state.projectName,
      repositoryName: sourceRepositoryName,
      environment: payload.environment,
      branchName: state.sourceBranch
    });
    const legacyPipelineFilename = buildLegacyPipelineFilename({
      projectName: state.projectName,
      repositoryName: sourceRepositoryName,
      branchName: state.sourceBranch
    });
    const legacyEnvironmentFirstPipelineFilename = buildLegacyEnvironmentFirstPipelineFilename({
      projectName: state.projectName,
      repositoryName: sourceRepositoryName,
      environment: payload.environment,
      branchName: state.sourceBranch
    });
    const pipelineName = buildPipelineName(pipelineFilename);
    const releaseName = buildReleaseName({
      service: payload.service,
      environment: payload.environment
    });

    if (!state.accessToken || !state.projectId) {
      setStatus('Open the extension from Azure DevOps to create the repositories, files, Pipeline, and Release.', true);
      setSubmitting(false);
      return yaml;
    }

    const targetBranch = SCAFFOLD_BRANCH;
    try {
      const provisioningProjectName = state.rawProjectName || state.projectName;
      let supportRepositories = [];
      const repo = await runProvisioningStep(
        'Step 1/5: creating or reusing the Azure, Docker, and Nginx DevOps repositories...',
        async () => {
          const pipelineRepo = await ensureRepo({
            hostUri: state.hostUri,
            projectId: state.projectId,
            projectName: provisioningProjectName,
            accessToken: state.accessToken
          });
          supportRepositories = await ensureSupportRepositories({
            hostUri: state.hostUri,
            projectId: state.projectId,
            projectName: provisioningProjectName,
            environment: payload.environment,
            domain: environmentConfig.domain,
            service: payload.service,
            repositoryAddress: payload.repositoryAddress,
            accessToken: state.accessToken
          });
          return pipelineRepo;
        }
      );
      state.generatedRepoId = repo.id || state.generatedRepoId;
      state.generatedRepositoryName = repo.name || state.generatedRepositoryName;
      state.branch = targetBranch;
      await runProvisioningStep(`Step 2/5: saving YAML file /${pipelineFilename}...`, () =>
        postScaffold({
          hostUri: state.hostUri,
          projectId: state.projectId,
          repoId: repo.id,
          branch: targetBranch,
          accessToken: state.accessToken,
          content: yaml,
          pipelineFilename
        })
      );
      await runProvisioningStep('Step 3/5: setting the generated repository default branch...', () =>
        ensureDefaultBranch({
          hostUri: state.hostUri,
          projectId: state.projectId,
          repoId: repo.id,
          branchName: targetBranch,
          accessToken: state.accessToken
        })
      );

      const pipelineDefinition = await runProvisioningStep(
        `Step 4/5: creating or updating Pipeline ${pipelineName} in ${PIPELINE_FOLDER}...`,
        () =>
          upsertPipelineDefinition({
            hostUri: state.hostUri,
            projectId: state.projectId,
            repo,
            pipelineName,
            pipelinePath: `/${pipelineFilename}`,
            legacyPipelineNames: [legacyEnvironmentFirstPipelineFilename, legacyPipelineFilename],
            legacyPipelinePaths: [
              `/${legacyEnvironmentFirstPipelineFilename}`,
              `/${legacyPipelineFilename}`
            ],
            branch: targetBranch,
            accessToken: state.accessToken
          })
      );

      const releaseDefinition = await runProvisioningStep(
        `Step 5/5: creating or reusing the classic Release definition ${releaseName}...`,
        () =>
          ensureReleaseDefinition({
            hostUri: state.hostUri,
            projectId: state.projectId,
            projectName: state.rawProjectName || state.projectName,
            repo,
            pipelineDefinition,
            pipelineName,
            service: payload.service,
            environment: payload.environment,
            branch: targetBranch,
            queueName: payload.pool,
            accessToken: state.accessToken
          })
      );

      const releaseMessage = releaseDefinition.skipped
        ? `Release skipped: ${releaseDefinition.reason}`
        : `Release definition ${
            releaseDefinition.created ? 'created' : releaseDefinition.updated ? 'updated' : 'already up to date'
          } (ID: ${releaseDefinition.id}).`;
      setStatus(
        `Done. Pipeline ${pipelineName} is linked to /${pipelineFilename} in ${PIPELINE_FOLDER} (ID: ${pipelineDefinition?.id || 'unknown'}). ${releaseMessage} Review the generated Nginx and Compose files below, then run the Pipeline manually.`,
        false
      );
      showCompletionLinks({ supportRepositories, pipelineDefinition });
    } catch (error) {
      console.error(error);
      const detail = sanitizeErrorDetail(error?.detail || error?.message || '');
      const step = error?.provisioningStep || 'Provisioning';
      const permissionHint = error?.requiredExtensionScope
        ? ` The installed extension token is missing or has not been reauthorized for ${error.requiredExtensionScope}. A Collection Administrator must authorize the updated Pipeline Generator scopes.`
        : error?.domain === 'release'
          ? ' Ask a project administrator to grant Manage release definitions, View releases, and Use the selected agent queue.'
          : error?.domain === 'pipeline'
            ? ' Ask a project administrator to grant Create/Edit pipeline permission under Project settings → Pipelines → Security.'
            : ' Ask a project administrator to grant the required Repos permissions.';
      const unauthorizedMessage = `Access denied during ${step}.${detail ? ` Details: ${detail}.` : ''}${permissionHint}`;
      const detailMessage = `Failed during ${step}.${detail ? ` Details: ${detail}` : ' No error details were returned by Azure DevOps.'}`;
      if (isUnauthorizedError(error)) {
        state.accessToken = null;
        setReauthenticationVisibility(
          true,
          error?.requiredExtensionScope
            ? `The installed extension token cannot access ${error.requiredExtensionScope}. Use Open extension authorization as a Collection Administrator, authorize the updated scopes, then reopen the generator from the branch.`
            : 'The Azure DevOps host token was denied. Sign out and authenticate again, then reopen the generator from the branch.'
        );
      }
      setStatus(isUnauthorizedError(error) ? unauthorizedMessage : detailMessage, true);
    }

    setSubmitting(false);
    return yaml;
  };

  form?.addEventListener('submit', handleSubmit);

  const fetchDockerfileDirectories = async ({ hostUri, projectId, repoId, branch, accessToken }) => {
    if (!repoId) return [];
    const versionDescriptor = branch
      ? `&versionDescriptor.version=${encodeURIComponent(branch)}&versionDescriptor.versionType=branch`
      : '';
    const url = `${hostUri}${encodeURIComponent(projectId)}/_apis/git/repositories/${repoId}/items?recursionLevel=Full&includeContentMetadata=true${versionDescriptor}&api-version=6.0`;
    const res = await fetch(url, { headers: authHeaders(accessToken) });
    if (!res.ok) {
      const detail = await readErrorDetail(res);
      throw buildHttpError('Failed to scan repository for Dockerfiles', res, detail);
    }
    const payload = await res.json();
    return (payload.value || [])
      .filter((item) => !item.isFolder && /(?:^|\/|\\)Dockerfile$/i.test(item.path || item.serverItem || ''))
      .map((item) => normalizeDockerfileDir(item.path || item.serverItem))
      .filter(Boolean);
  };

  const init = async () => {
    setSubmitting(true);
    setStatus('Loading Azure DevOps context...');
    populateDefaults();
    const query = new URLSearchParams(window.location.search);
    const branchFromQuery = getQueryValue(query.get('branch'));
    const projectIdFromQuery = getQueryValue(query.get('projectId'));
    const projectNameFromQuery = getQueryValue(query.get('projectName')) || projectIdFromQuery;
    const repoIdFromQuery = getQueryValue(query.get('repoId'));
    const repoNameFromQuery = getQueryValue(query.get('repoName'));
    const initialBranch = branchFromQuery || '(unknown branch)';
    const isFramed = window.parent !== window;

    const hostLooksLikeAzureDevOps = (() => {
      const candidateOrigins = new Set();
      const addOrigin = (value) => {
        try {
          if (value) {
            candidateOrigins.add(new URL(value).origin);
          }
        } catch {
          /* ignore invalid URLs */
        }
      };

      addOrigin(window.location.origin);
      addOrigin(document.referrer);
      if (window.location.ancestorOrigins) {
        try {
          const rawAncestors = window.location.ancestorOrigins;
          const ancestors = [];

          if (typeof rawAncestors.forEach === 'function') {
            rawAncestors.forEach((value) => ancestors.push(value));
          } else {
            const length = Number(rawAncestors.length) || 0;
            for (let i = 0; i < length; i += 1) {
              ancestors.push(rawAncestors[i]);
            }
          }

          ancestors.forEach(addOrigin);
        } catch (error) {
          console.warn('Skipping ancestorOrigins inspection', error);
        }
      }
      if (candidateOrigins.size === 0) return false;

      return Array.from(candidateOrigins).some((origin) => {
        try {
          const { hostname } = new URL(origin);
          return (
            origin === window.location.origin ||
            hostname.toLowerCase().endsWith('dev.azure.com') ||
            hostname.toLowerCase().endsWith('visualstudio.com')
          );
        } catch {
          return false;
        }
      });
    })();

    // Only attempt SDK initialization when the extension is running inside the
    // Azure DevOps dialog or hub iframe. Direct asset URLs remain in offline
    // mode to avoid noisy VSS handshake errors.
    const shouldAttemptSdk = isFramed && hostLooksLikeAzureDevOps;

    state.sourceBranch = initialBranch;
    state.projectId = projectIdFromQuery;
    state.rawProjectName = projectNameFromQuery;
    state.projectName = projectNameFromQuery;
    state.repoId = repoIdFromQuery;
    state.rawRepositoryName = repoNameFromQuery;
    state.repositoryName = repoNameFromQuery;
    state.hostUri = `${getHostBase().replace(/\/+$/, '')}/`;

    branchLabel.textContent = branchFromQuery
      ? `Target branch: ${SCAFFOLD_BRANCH} (source: ${initialBranch})`
      : 'Loading branch context...';
    if (branchInput && branchFromQuery) {
      branchInput.value = SCAFFOLD_BRANCH;
      branchInput.disabled = true;
    }
    targetRepoInput.value = `${projectNameFromQuery || 'project'}_Azure_DevOps`;
    setServiceNameFromRepository(repoNameFromQuery || projectNameFromQuery, projectNameFromQuery);
    // A host dialog is a trusted candidate even when an on-premises
    // referrer-policy removes document.referrer. The VSS handshake itself is
    // the authority; an unrelated parent cannot complete it successfully.
    const hasHostContext = isFramed;
    if (!hasHostContext || !shouldAttemptSdk) {
      setStatus(
        'Running outside Azure DevOps. Open the extension from a branch action to create the repository and pipeline file automatically.',
        true
      );
      setSubmitting(false);
      return;
    }

    try {
      const sdk = await loadVssSdk();
      sdk.init({ usePlatformScripts: true, explicitNotifyLoaded: true });
      await waitForSdkReady(sdk);

      const context = sdk.getWebContext();
      const dialogConfiguration = getDialogConfiguration(sdk);
      const hostNavigationState = await getHostNavigationState(sdk);
      const hostedConfiguration = { ...hostNavigationState, ...dialogConfiguration };

      const branch =
        getQueryValue(hostedConfiguration.branch) ||
        branchFromQuery ||
        context?.repository?.defaultBranch?.replace(/^refs\/heads\//, '') ||
        '(unknown branch)';
      state.sourceBranch = branch;

      const projectId = getQueryValue(hostedConfiguration.projectId) || projectIdFromQuery || context?.project?.id;
      const projectName =
        getQueryValue(hostedConfiguration.projectName) || projectNameFromQuery || context?.project?.name || projectId;
      const repoId = getQueryValue(hostedConfiguration.repoId) || repoIdFromQuery || context?.repository?.id;
      let repositoryName =
        getQueryValue(hostedConfiguration.repoName) || repoNameFromQuery || context?.repository?.name;
      state.sdk = sdk;
      state.projectId = projectId;
      state.rawProjectName = projectName || state.rawProjectName;
      state.projectName = projectName;
      state.repoId = repoId;
      state.rawRepositoryName = repositoryName || state.rawRepositoryName;
      state.repositoryName = repositoryName;

      branchLabel.textContent = `Target branch: ${SCAFFOLD_BRANCH} (source: ${branch})`;
      if (branchInput) {
        branchInput.value = SCAFFOLD_BRANCH;
        branchInput.disabled = true;
      }
      targetRepoInput.value = `${projectName || 'project'}_Azure_DevOps`;
      setServiceNameFromRepository(repositoryName || projectName, projectName);

      if (!projectId) {
        setStatus('Project context was not provided by the branch action or hub.', true);
        sdk.notifyLoadFailed('Missing project context');
        return;
      }

      const hostUri = normalizeHostUri(hostedConfiguration.hostUri || context.collection?.uri || getHostBase());
      state.hostUri = hostUri;
      let accessToken = state.accessToken;
      let accessTokenError = null;

      if (!accessToken) {
        try {
          accessToken = await getAccessTokenWithRetry(sdk);
        } catch (tokenError) {
          console.error('Failed to acquire Azure DevOps access token', tokenError);
          accessTokenError = normalizeAccessTokenError(tokenError);
        }
      }

      state.accessTokenError = accessTokenError;

      if (!accessToken) {
        const errorMessage =
          accessTokenError ||
          'Failed to acquire access token from Azure DevOps. Reload the page and relaunch the generator from a branch action.';
        setReauthenticationVisibility(
          true,
          buildTokenRecoveryMessage(errorMessage)
        );
        setStatus(errorMessage, true);
        sdk.notifyLoadSucceeded?.();
        return;
      }

      try {
        state.accessToken = accessToken;
        setReauthenticationVisibility(false);
        if (!repositoryName && repoId) {
          try {
            const repoUrl = `${hostUri}${encodeURIComponent(projectId)}/_apis/git/repositories/${encodeURIComponent(
              repoId
            )}?api-version=6.0`;
            const repoRes = await fetch(repoUrl, { headers: authHeaders(accessToken) });
            if (repoRes.ok) {
              const repoPayload = await repoRes.json();
              repositoryName = repoPayload?.name || repositoryName;
              state.rawRepositoryName = repositoryName || state.rawRepositoryName;
              state.repositoryName = repositoryName;
              setServiceNameFromRepository(repositoryName, projectName);
            }
          } catch (repoError) {
            console.warn('Failed to fetch repository metadata', repoError);
          }
        }
        await loadDeploymentTargets({ hostUri, accessToken, branch });
        await Promise.all([
          loadPools({ hostUri, projectId, accessToken }),
          loadContainerRegistries({ hostUri, projectId, accessToken }),
          refreshDockerfiles({ hostUri, projectId, repoId, branch, accessToken })
        ]);
        setStatus('Azure DevOps context ready. Generate the pipeline when you are ready.');
      } catch (tokenError) {
        console.error('Failed to initialize Azure DevOps context', tokenError);
        const detail = sanitizeErrorDetail(tokenError?.detail || tokenError?.message || '');
        let detailMessage;
        if (state.deploymentTargetsReady) {
          detailMessage =
            accessTokenError || 'Failed to initialize Azure DevOps resources. Reload the page and try again.';
        } else if (tokenError?.domain === 'komodo') {
          detailMessage = `Could not load active Komodo servers. ${
            detail || 'Verify the central credential file, Komodo API access, TLS certificate, and CORS origin.'
          }`;
        } else {
          detailMessage = `Could not load ${DEPLOYMENT_TARGETS_CONFIG.collection}/${DEPLOYMENT_TARGETS_CONFIG.project}/${DEPLOYMENT_TARGETS_CONFIG.repository}:${DEPLOYMENT_TARGETS_CONFIG.path}. ${
            detail || 'Verify the file, branch, YAML structure, and repository Read permission.'
          }`;
        }
        if (isUnauthorizedError(tokenError)) {
          setReauthenticationVisibility(
            true,
            'The Azure DevOps host token could not access the required APIs. Sign out and authenticate again, then reopen the generator.'
          );
        }
        setStatus(detailMessage, true);
        sdk.notifyLoadSucceeded?.();
        return;
      }

      sdk.notifyLoadSucceeded();
    } catch (error) {
      console.error('Failed to initialize extension frame', error);
      const fallbackMessage = /Timed out waiting for Azure DevOps host/i.test(error?.message || '')
        ? 'Could not connect to the Azure DevOps host. If you opened this page directly, use the form to generate the YAML and copy it below.'
        : 'Failed to initialize extension frame. Check extension permissions and reload, or copy the template below.';
      if (state.projectId && state.hostUri) {
        setReauthenticationVisibility(
          true,
          `${fallbackMessage} Sign out and authenticate again below to rebuild the Azure DevOps host session.`
        );
      }
      setStatus(fallbackMessage, true);
      const sdk = normalizeSdk(window.VSS || window.parent?.VSS);
      sdk?.notifyLoadSucceeded?.();
    } finally {
      setSubmitting(false);
    }
  };

  const startInitialization = () => {
    if (initializationPromise) {
      return initializationPromise;
    }
    setStatus('Loading Azure DevOps context...');
    initializationPromise = init();
    return initializationPromise;
  };

  if (environmentSelect) {
    environmentSelect.addEventListener('change', (event) => {
      setKomodoServerFromEnvironment(event.target.value);
    });
  }

  reauthenticateButton?.addEventListener('click', () => {
    restartAzureDevOpsSession();
  });

  authorizeExtensionButton?.addEventListener('click', () => {
    openExtensionAuthorization();
  });

  window.addEventListener('message', (event) => {
    if (!event?.data || event.origin !== window.location.origin) return;
    if (event.data.type === 'pipeline-bootstrap') {
      applyBootstrapPayload(event.data.payload || {}, 'message');
      event.source?.postMessage({ type: 'pipeline-bootstrap-ack' }, event.origin);
    }
  });

  startInitialization();
})();
