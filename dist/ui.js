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
    komodoServer: 'DEMO-192.168.62.91',
    dockerfileDir: '**'
  };
  const defaultPoolOptions = ['PublishDockerAgent', 'Default'];
  const defaultRegistryOptions = ['BulutReg', 'DockerReg'];
  const environmentKomodoMap = {
    dev: 'Development-192.168.62.19',
    demo: 'DEMO-192.168.62.91',
    qa: 'QA-192.168.62.153',
    pro: 'Production-31.7.65.195'
  };

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
    nameSuffix: '_Release',
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
      nameSuffix: String(configured.nameSuffix || DEFAULT_RELEASE_CONFIG.nameSuffix),
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
      submitButton.disabled = isSubmitting;
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

  const buildPipelineFilename = ({ projectName, repositoryName, branchName }) => {
    const sanitizeSegment = (segment, fallback) => {
      const fallbackValue = fallback?.toString().toLowerCase() || '';
      const value = segment?.toString().trim().toLowerCase();
      const base = value || fallbackValue;

      const cleaned = base
        .replace(/[\\/]+/g, '-')
        .replace(/[^\w.-]+/g, '-')
        .replace(/-+/g, '-')
        .replace(/^-+|-+$/g, '');

      return cleaned || fallbackValue || 'segment';
    };

    const projectSegment = sanitizeSegment(projectName, 'project');
    const repoSegment = sanitizeSegment(repositoryName || projectName, 'repo');
    const branchSegment = sanitizeSegment(branchName?.replace(/^refs\/heads\//, ''), 'branch');
    return `${projectSegment}-${repoSegment}-${branchSegment}.yml`;
  };

  const buildPipelineName = (pipelineFilename) => pipelineFilename;

  const getProjectRouteSegment = () => {
    const candidate = state.rawProjectName || state.projectName || state.projectId;
    return candidate ? encodeURIComponent(candidate) : '';
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
    applyDetectedEnvironment(sourceBranch || targetBranch);
    setKomodoServerFromEnvironment(environmentSelect?.value);

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
    const normalizedTarget = normalizeName(targetName);
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
    const target = environmentKomodoMap[environment.toLowerCase()];
    if (!target) return;
    const match = Array.from(komodoSelect.options).find((option) => option.value === target);
    if (match) {
      komodoSelect.value = target;
    }
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
      : ['dev', 'demo', 'qa', 'pro'];

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
      throw buildHttpError('Failed to read the existing generated YAML', res, detail);
    }

    return res.text();
  };

  const ensureRepo = async ({ hostUri, projectId, projectName, accessToken }) => {
    const targetName = `${projectName}_Azure_DevOps`;
    targetRepoInput.value = targetName;
    const url = `${hostUri}${encodeURIComponent(projectId)}/_apis/git/repositories?api-version=6.0`;

    const res = await fetch(url, {
      headers: authHeaders(accessToken)
    });
    if (!res.ok) {
      const detail = await readErrorDetail(res);
      throw buildHttpError('Failed to list repositories', res, detail);
    }
    const payload = await res.json();
    const existing = (payload.value || []).find((repo) => repo.name === targetName);
    if (existing) {
      return existing;
    }

    const createRes = await fetch(url, {
      method: 'POST',
      headers: {
        ...authHeaders(accessToken),
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({ name: targetName, project: { id: projectId } })
    });
    if (!createRes.ok) {
      const detail = await readErrorDetail(createRes);
      throw buildHttpError('Failed to create repository', createRes, detail);
    }
    return createRes.json();
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

  const postScaffold = async ({
    hostUri,
    projectId,
    repoId,
    branch,
    accessToken,
    content,
    pipelineFilename = 'project-repo-branch.yml'
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

  const getPipelineByName = async ({ hostUri, projectId, pipelineName, accessToken }) => {
    const url = buildPipelinesApiUrl({ hostUri, projectId });
    const res = await fetch(url, { headers: authHeaders(accessToken) });
    if (!res.ok) {
      const detail = await readErrorDetail(res);
      throw markErrorDomain(buildHttpError('Failed to list Azure Pipelines', res, detail), 'pipeline');
    }
    const payload = await res.json();
    return (payload.value || []).find((pipeline) => pipeline.name === pipelineName);
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

  const upsertPipelineDefinition = async ({ hostUri, projectId, repo, pipelineName, pipelinePath, branch, accessToken }) => {
    const repositoryName = `${state.projectName || projectId}/${repo.name}`;
    const desiredConfig = buildPipelineConfiguration({
      repoId: repo.id,
      repositoryName,
      pipelinePath,
      branch
    });

    const existing = await getPipelineByName({ hostUri, projectId, pipelineName, accessToken });
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

    const releaseName = `${pipelineName}${releaseConfig.nameSuffix}`;
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
      `    environment: '${payload.environment || ''}'           # dev/demo/pro/qa`,
      `    dockerfileDir: '${payload.dockerfileDir || '**'}'  # path of Dockerfile, Default is '**'`,
      `    repositoryAddress: '${payload.repositoryAddress || ''}'`,
      `    containerRegistryService: '${payload.containerRegistryService || ''}'`,
      "    tag: '1.0.$(Build.BuildId)'",
      `    komodoServer: '${payload.komodoServer || ''}' # or 'Development-192.168.62.19' or 'Production-31.7.65.195'`,
      "    komodoApiKey: '$(KOMODO_API_KEY)'",
      "    komodoApiSecret: '$(KOMODO_API_SECRET)'",
      '    sourceRepo: otherRepo',
      ''
    ].join('\n');
  };

  const handleSubmit = async (event) => {
    event.preventDefault();
    if (initializationPromise) {
      try {
        await initializationPromise;
      } catch (error) {
        console.error('Initialization failed before submit', error);
      }
    }
    const payload = Object.fromEntries(new FormData(form).entries());
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
      branchName: state.sourceBranch || payload.environment
    });
    const pipelineName = buildPipelineName(pipelineFilename);

    if (!state.accessToken || !state.projectId) {
      setStatus('Open the extension from Azure DevOps to push the template and open it automatically.', true);
      setSubmitting(false);
      return yaml;
    }

    const targetBranch = SCAFFOLD_BRANCH;
    try {
      const repo = await runProvisioningStep('Step 1/5: creating or reusing the pipeline repository...', () =>
        ensureRepo({
          hostUri: state.hostUri,
          projectId: state.projectId,
          projectName: state.projectName,
          accessToken: state.accessToken
        })
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
            branch: targetBranch,
            accessToken: state.accessToken
          })
      );

      const releaseDefinition = await runProvisioningStep(
        `Step 5/5: creating or reusing the classic Release definition for ${pipelineName}...`,
        () =>
          ensureReleaseDefinition({
            hostUri: state.hostUri,
            projectId: state.projectId,
            projectName: state.rawProjectName || state.projectName,
            repo,
            pipelineDefinition,
            pipelineName,
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
        `Done. Pipeline ${pipelineName} is linked to /${pipelineFilename} in ${PIPELINE_FOLDER} (ID: ${pipelineDefinition?.id || 'unknown'}). ${releaseMessage} Opening Pipelines...`,
        false
      );
      const projectRoute = getProjectRouteSegment();
      window.setTimeout(() => {
        navigateHost(`${state.hostUri}${projectRoute}/_build?definitionId=${encodeURIComponent(pipelineDefinition.id)}`);
      }, 700);
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
    applyDetectedEnvironment(initialBranch);
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
      applyDetectedEnvironment(branch);
      setKomodoServerFromEnvironment(environmentSelect?.value);

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
        await Promise.all([
          loadPools({ hostUri, projectId, accessToken }),
          loadContainerRegistries({ hostUri, projectId, accessToken }),
          refreshDockerfiles({ hostUri, projectId, repoId, branch, accessToken })
        ]);
        setStatus('Azure DevOps context ready. Generate the pipeline when you are ready.');
      } catch (tokenError) {
        console.error('Failed to initialize Azure DevOps context', tokenError);
        const detailMessage =
          accessTokenError || 'Failed to acquire access token from Azure DevOps. Reload the page and try again.';
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
