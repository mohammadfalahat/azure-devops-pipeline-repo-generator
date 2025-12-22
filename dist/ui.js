(() => {
  const getHostBase = () => {
    if (!document.referrer) {
      return window.location.origin;
    }
    const referrer = new URL(document.referrer);
    const segments = referrer.pathname.split('/').filter(Boolean);
    const hasTfsVirtualDir = segments[0]?.toLowerCase() === 'tfs';
    return `${referrer.origin}${hasTfsVirtualDir ? '/tfs' : ''}`;
  };

  const loadScript = (src) =>
    new Promise((resolve, reject) => {
      const script = document.createElement('script');
      script.src = src;
      script.async = false;
      script.crossOrigin = 'use-credentials';
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

    const hostSdk = `${getHostBase()}/_content/MS.VSS.SDK/scripts/VSS.SDK.min.js`;
    const localSdk = new URL('./lib/VSS.SDK.min.js', window.location.href).toString();
    const localSdkFallback = new URL('./lib/VSS.SDK.js', window.location.href).toString();
    // Prefer bundled SDK assets first because some Azure DevOps hosts block direct downloads
    // of the platform SDK (e.g., returning an HTML login page with a text/html MIME type).
    // Trying local files first avoids those MIME-type failures while keeping the host SDK
    // as a last-resort option for environments that rely on it being served directly.
    const candidates = [localSdk, localSdkFallback, hostSdk];

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

  const getAuthHeader = (token) => {
    const tokenValue = typeof token === 'string' ? token : token?.token;
    if (!tokenValue) {
      throw new Error('Extension access token was unavailable.');
    }
    const isLikelyJwt = tokenValue.split('.').length === 3;
    if (isLikelyJwt) {
      return `Bearer ${tokenValue}`;
    }
    const encoded = btoa(`:${tokenValue}`);
    return `Basic ${encoded}`;
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

  const buildHttpError = (baseMessage, response, detail) => {
    const message = `${baseMessage} (${response.status})${detail ? `: ${detail}` : ''}`;
    const error = new Error(message);
    error.status = response.status;
    if (detail) {
      error.detail = detail;
    }
    return error;
  };

  const delay = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

  const getAccessTokenFromSdk = async (sdk, maxAttempts = 3, delayMs = 800) => {
    if (!sdk?.getAccessToken) {
      throw new Error('Azure DevOps access token API is unavailable.');
    }

    // Explicitly request the scopes required to create pipelines so on-premises
    // servers issue a token that can manage build definitions (TF400813/401
    // otherwise occur when the returned token only covers Repos).
    const requestedScope = ['vso.code', 'vso.code_manage', 'vso.project', 'vso.build'].join(' ');

    let lastError;
    for (let attempt = 1; attempt <= maxAttempts; attempt += 1) {
      try {
        let token;
        try {
          token = await sdk.getAccessToken({ scope: requestedScope });
        } catch (scopeError) {
          // Older SDKs do not support the scoped call; fall back to the default
          // behavior so the token acquisition still succeeds.
          token = await sdk.getAccessToken();
          if (!token) {
            throw scopeError;
          }
        }
        if (token) {
          return token;
        }
        lastError = new Error('Azure DevOps did not provide an access token.');
      } catch (error) {
        lastError = error;
      }

      if (attempt < maxAttempts) {
        await delay(delayMs * attempt);
      }
    }

    throw lastError;
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

  const buildPipelineName = ({ projectName, repositoryName, environment }) => {
    const projectSegment = projectName || 'project';
    const repoSegment = repositoryName || projectName || 'repo';
    const environmentSegment = environment || 'env';
    return `${projectSegment}_${repoSegment}_${environmentSegment}`;
  };

  const isUnauthorizedError = (error) =>
    error?.status === 401 ||
    error?.status === 403 ||
    /TF400813/i.test(error?.detail || '') ||
    /\b401\b/.test(error?.message || '');

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

    const normalizedHost = (hostUri || state.hostUri || getHostBase()).replace(/\/+$/, '') + '/';
    state.sourceBranch = branch || state.sourceBranch;
    state.branch = SCAFFOLD_BRANCH;
    state.projectId = projectId || state.projectId;
    state.rawProjectName = projectName || state.rawProjectName;
    state.projectName = projectName || state.projectName;
    state.repoId = repoId || state.repoId;
    state.rawRepositoryName = repoName || state.rawRepositoryName;
    state.repositoryName = repoName || state.repositoryName;
    state.hostUri = normalizedHost;
    state.accessToken = accessToken || state.accessToken;
    state.accessTokenError = accessTokenError || state.accessTokenError;

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
        const needsHostAuth = /HostAuthorizationNotFound/i.test(state.accessTokenError);
        authMessage = needsHostAuth
          ? 'Azure DevOps could not issue an access token because host authorization was not found. Confirm the extension is installed and enabled for this collection/project (Organization/Collection settings → Extensions → Manage) and that your account can access it, then relaunch the generator.'
          : `Azure DevOps did not provide an access token (${state.accessTokenError}). Refresh the page or sign in again, then relaunch the generator.`;
      } else {
        authMessage = 'Loaded context from branch action but still waiting for an access token from Azure DevOps. Refresh or try again if this persists.';
      }
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
          branch: targetBranch,
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
      return '0000000000000000000000000000000000000000';
    }

    if (!res.ok) {
      const detail = await readErrorDetail(res);
      throw buildHttpError('Failed to query branch', res, detail);
    }

    const payload = await res.json();
    return payload.value?.[0]?.objectId || '0000000000000000000000000000000000000000';
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
    const authHeader = getAuthHeader(accessToken);
    const oldObjectId = await getBranchObjectId({ hostUri, projectId, repoId, branch: branchName, accessToken });
    const body = {
      refUpdates: [
        {
          name: branchRef,
          oldObjectId
        }
      ],
      commits: [
        {
          comment: 'Add pipeline generator defaults',
          changes: [
            {
              changeType: 'add',
              item: { path: `/${pipelineFilename}` },
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

  const getPipelineByName = async ({ hostUri, projectId, pipelineName, accessToken }) => {
    const url = `${hostUri}${encodeURIComponent(projectId)}/_apis/pipelines?api-version=7.1`;
    const res = await fetch(url, { headers: authHeaders(accessToken) });
    if (!res.ok) {
      return undefined;
    }
    const payload = await res.json();
    return (payload.value || []).find((pipeline) => pipeline.name === pipelineName);
  };

  const getPipelineById = async ({ hostUri, projectId, pipelineId, accessToken }) => {
    const url = `${hostUri}${encodeURIComponent(projectId)}/_apis/pipelines/${pipelineId}?api-version=7.1`;
    const res = await fetch(url, { headers: authHeaders(accessToken) });
    if (!res.ok) {
      return undefined;
    }
    return res.json();
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
      const current = await getPipelineById({ hostUri, projectId, pipelineId: existing.id, accessToken });
      const needsUpdate =
        current?.configuration?.path !== desiredConfig.path ||
        current?.configuration?.repository?.id !== desiredConfig.repository.id ||
        current?.configuration?.repository?.defaultBranch !== desiredConfig.repository.defaultBranch;

      if (!needsUpdate) {
        return current || existing;
      }

      const updateUrl = `${hostUri}${encodeURIComponent(projectId)}/_apis/pipelines/${existing.id}?api-version=7.1`;
      const res = await fetch(updateUrl, {
        method: 'PUT',
        headers: {
          ...authHeaders(accessToken),
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({ name: pipelineName, configuration: desiredConfig })
      });

      if (!res.ok) {
        const detail = await readErrorDetail(res);
        throw buildHttpError('Failed to update pipeline', res, detail);
      }

      return res.json();
    }

    const createUrl = `${hostUri}${encodeURIComponent(projectId)}/_apis/pipelines?api-version=7.1`;
    const res = await fetch(createUrl, {
      method: 'POST',
      headers: {
        ...authHeaders(accessToken),
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({ name: pipelineName, folder: '\\', configuration: desiredConfig })
    });

    if (!res.ok) {
      const detail = await readErrorDetail(res);
      throw buildHttpError('Failed to create pipeline', res, detail);
    }

    return res.json();
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
      "trigger: none                      # always none",
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
      `      name: "${projectRepoName}"             # PROJECTNAME/REPONAME`,
      `      ref: refs/heads/${sourceBranchName}        # refs/heads/BRANCH`,
      '      trigger:',
      '        branches:',
      '          include:',
      `            - ${sourceBranchName}                # BRANCH`,
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
      sourceRepositoryName: state.repositoryName || state.projectName
    });

    setStatus('Generating pipeline template...');
    setSubmitting(true);

    if (!state.accessToken && state.sdk?.getAccessToken) {
      try {
        state.accessToken = await getAccessTokenFromSdk(state.sdk);
      } catch (error) {
        console.error('Failed to refresh access token during submit', error);
      }
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

    const pipelineFilename = buildPipelineFilename({
      projectName: state.projectName,
      repositoryName: state.repositoryName,
      branchName: state.sourceBranch || payload.environment
    });
    const pipelineName = buildPipelineName({
      projectName: state.projectName,
      repositoryName: state.repositoryName,
      environment: payload.environment
    });

    if (!state.accessToken || !state.projectId) {
      setStatus('Open the extension from Azure DevOps to push the template and create the pipeline automatically.', true);
      setSubmitting(false);
      return yaml;
    }

    const targetBranch = SCAFFOLD_BRANCH;
    try {
      const repo = await ensureRepo({
        hostUri: state.hostUri,
        projectId: state.projectId,
        projectName: state.projectName,
        accessToken: state.accessToken
      });
      state.repoId = repo.id || state.repoId;
      state.repositoryName = repo.name || state.repositoryName;
      state.branch = targetBranch;
      await postScaffold({
        hostUri: state.hostUri,
        projectId: state.projectId,
        repoId: repo.id,
        branch: targetBranch,
        accessToken: state.accessToken,
        content: yaml,
        pipelineFilename
      });
      await ensureDefaultBranch({
        hostUri: state.hostUri,
        projectId: state.projectId,
        repoId: repo.id,
        branchName: targetBranch,
        accessToken: state.accessToken
      });

      await upsertPipelineDefinition({
        hostUri: state.hostUri,
        projectId: state.projectId,
        repo,
        pipelineName,
        pipelinePath: pipelineFilename,
        branch: targetBranch,
        accessToken: state.accessToken
      });

      setStatus(`Pipeline ${pipelineName} created. Redirecting...`, false);
      window.location.href = `${state.hostUri}${encodeURIComponent(state.projectId)}/_build`;
    } catch (error) {
      console.error(error);
      const detail = sanitizeErrorDetail(error?.detail || error?.message || '');
      const manualPath = `/${pipelineFilename}`;
      const createApiUrl = `${state.hostUri}${encodeURIComponent(state.projectId)}/_apis/pipelines?api-version=7.1`;
      const curlExample =
        `curl -u :<PAT_WITH_PIPELINE_SCOPE> -H "Content-Type: application/json" ` +
        `-d @pipeline.json "${createApiUrl}"`;
      const unauthorizedMessage =
        `Automatic pipeline creation failed: access was denied${detail ? ` (${detail})` : ''}. ` +
        `${state.accessToken ? 'The token from Azure DevOps may not include pipeline creation rights for this project; ask a project administrator to grant Create pipeline permission or retry with a token that includes that scope.' : 'Open the extension from Azure DevOps so we can request a project-scoped token with pipeline creation rights.'} ` +
        `You can still create the pipeline manually with the generated YAML at ${manualPath}: ` +
        `Pipelines > New pipeline > Azure Repos Git > Existing Azure Pipelines YAML (or open ${state.hostUri}${encodeURIComponent(
          state.projectId
        )}/_build?view=pipelines and choose that option), then select branch '${targetBranch}' and path '${manualPath}'. ` +
        `If you want to test the REST API directly with your own credentials, POST to ${createApiUrl} (for example: ${curlExample}). ` +
        `Include the pipeline body (name + configuration) in pipeline.json; an empty file returns "Value cannot be null. Parameter name: inputParameters". ` +
        `The repository.id in that body must be the GUID of the YAML repository (for example from /_apis/git/repositories); leaving it blank yields "repositoryId must not be Guid.Empty."`;
      const detailMessage = error?.message ? `Automatic pipeline creation failed: ${error.message}` : 'Automatic pipeline creation failed.';
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
    const res = await fetch(url, { headers: { Authorization: getAuthHeader(accessToken) } });
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
    const hasReferrer = Boolean(document.referrer);
    const isFramed = window.parent !== window;
    const hasOpener = Boolean(window.opener);

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
    // Azure DevOps iframe host. Opening the form in a new tab (for example via
    // the window.open fallback) should remain in offline mode to avoid noisy
    // VSS handshake errors.
    const shouldAttemptSdk = isFramed && hostLooksLikeAzureDevOps;

    state.sourceBranch = initialBranch;
    state.projectId = projectIdFromQuery;
    state.projectName = projectNameFromQuery;
    state.repoId = repoIdFromQuery;
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
    const hasHostContext = Boolean(isFramed && (hasReferrer || projectIdFromQuery || repoIdFromQuery || hasOpener));
    if (!hasHostContext || !shouldAttemptSdk) {
      setStatus(
        'Running outside Azure DevOps. Open the extension from a branch action to create the repository and pipeline automatically.',
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

      const branch =
        branchFromQuery ||
        context?.repository?.defaultBranch?.replace(/^refs\/heads\//, '') ||
        '(unknown branch)';
      state.sourceBranch = branch;

      const projectId = projectIdFromQuery || context?.project?.id;
      const projectName = projectNameFromQuery || context?.project?.name || projectId;
      const repoId = repoIdFromQuery || context?.repository?.id;
      let repositoryName = repoNameFromQuery || context?.repository?.name;
      state.sdk = sdk;
      state.projectId = projectId;
      state.projectName = projectName;
      state.repoId = repoId;
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

      const hostUri = (context.collection?.uri || getHostBase()).replace(/\/+$/, '') + '/';
      state.hostUri = hostUri;
      let accessToken;

      try {
        accessToken = await getAccessTokenFromSdk(sdk);
        state.accessToken = accessToken;
        if (!repositoryName && repoId) {
          try {
            const repoUrl = `${hostUri}${encodeURIComponent(projectId)}/_apis/git/repositories/${encodeURIComponent(
              repoId
            )}?api-version=6.0`;
            const repoRes = await fetch(repoUrl, { headers: authHeaders(accessToken) });
            if (repoRes.ok) {
              const repoPayload = await repoRes.json();
              repositoryName = repoPayload?.name || repositoryName;
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
        console.error('Failed to acquire Azure DevOps access token', tokenError);
        setStatus('Failed to acquire access token from Azure DevOps. Reload the page and try again.', true);
        sdk.notifyLoadFailed?.('Access token unavailable');
        return;
      }

      sdk.notifyLoadSucceeded();
    } catch (error) {
      console.error('Failed to initialize extension frame', error);
      const fallbackMessage = /Timed out waiting for Azure DevOps host/i.test(error?.message || '')
        ? 'Could not connect to the Azure DevOps host. If you opened this page directly, use the form to generate the YAML and copy it below.'
        : 'Failed to initialize extension frame. Check extension permissions and reload, or copy the template below.';
      setStatus(fallbackMessage, true);
      const sdk = normalizeSdk(window.VSS || window.parent?.VSS);
      sdk?.notifyLoadFailed?.(error?.message || 'Initialization failed');
    } finally {
      setSubmitting(false);
    }
  };

  if (environmentSelect) {
    environmentSelect.addEventListener('change', (event) => {
      setKomodoServerFromEnvironment(event.target.value);
    });
  }

  window.addEventListener('message', (event) => {
    if (!event?.data || event.origin !== window.location.origin) return;
    if (event.data.type === 'pipeline-bootstrap') {
      applyBootstrapPayload(event.data.payload || {}, 'message');
      event.source?.postMessage({ type: 'pipeline-bootstrap-ack' }, event.origin);
    }
  });

  initializationPromise = init();
})();
