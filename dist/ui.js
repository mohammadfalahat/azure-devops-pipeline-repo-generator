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
      script.crossOrigin = 'anonymous';
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
    const candidates = [localSdk, localSdkFallback, `${getHostBase()}/_content/MS.VSS.SDK/scripts/VSS.SDK.min.js`];

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

  const defaultValues = {
    pool: 'PublishDockerAgent',
    environment: 'demo',
    repositoryAddress: 'registry.buluttakin.com',
    containerRegistryService: 'BulutReg',
    komodoServer: 'DEMO-192.168.62.91'
  };
  const defaultPoolOptions = ['PublishDockerAgent', 'Default'];
  const defaultRegistryOptions = ['BulutReg', 'DockerReg'];
  const tokenStorageKey = 'pipeline-generator.settings.token';
  const hostUriRef = { current: `${getHostBase().replace(/\/+$/, '')}/` };
  const onTokenUpdatedRef = { handler: () => {} };

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
  const pipelineSection = document.getElementById('pipeline-section');
  const tokenGate = document.getElementById('token-gate');
  const tokenGateMessage = document.getElementById('token-gate-message');
  const settingsPanel = document.getElementById('settings-panel');
  const openSettingsButton = document.getElementById('open-settings');
  const openSettingsFromGateButton = document.getElementById('open-settings-from-gate');
  const closeSettingsButton = document.getElementById('close-settings');
  const settingsForm = document.getElementById('settings-form');
  const tokenInput = document.getElementById('personalToken');
  const clearTokenButton = document.getElementById('clear-token');
  const tokenStatus = document.getElementById('token-status');

  const setStatus = (message, isError = false) => {
    status.textContent = message;
    status.className = isError ? 'status-error' : 'status-success';
  };

  const setTokenGateStatus = (message, isError = false) => {
    if (!tokenGateMessage) return;
    tokenGateMessage.textContent = message || '';
    tokenGateMessage.className = isError ? 'status-error' : '';
  };

  const showPipelineForm = () => {
    pipelineSection?.classList.remove('hidden');
    tokenGate?.classList.add('hidden');
  };

  const hidePipelineForm = () => {
    pipelineSection?.classList.add('hidden');
    tokenGate?.classList.remove('hidden');
  };

  const openSettingsPanel = () => {
    settingsPanel?.classList.remove('hidden');
    openSettingsButton?.setAttribute('aria-expanded', 'true');
  };

  const closeSettingsPanel = () => {
    settingsPanel?.classList.add('hidden');
    openSettingsButton?.setAttribute('aria-expanded', 'false');
  };

  const readStoredToken = () => {
    try {
      const raw = window.localStorage?.getItem(tokenStorageKey);
      if (!raw) return {};
      const parsed = JSON.parse(raw);
      if (typeof parsed === 'string') {
        return { token: parsed };
      }
      if (parsed && typeof parsed.token === 'string') {
        return { token: parsed.token, savedAt: parsed.savedAt };
      }
    } catch (error) {
      console.warn('Failed to read stored token', error);
    }
    return {};
  };

  const persistToken = (token) => {
    if (!token) return false;
    try {
      window.localStorage?.setItem(tokenStorageKey, JSON.stringify({ token, savedAt: new Date().toISOString() }));
      return true;
    } catch (error) {
      console.error('Failed to persist token', error);
      return false;
    }
  };

  const clearStoredToken = () => {
    try {
      window.localStorage?.removeItem(tokenStorageKey);
      return true;
    } catch (error) {
      console.error('Failed to clear stored token', error);
      return false;
    }
  };

  const setTokenStatus = (message, isError = false) => {
    if (!tokenStatus) return;
    tokenStatus.textContent = message;
    tokenStatus.className = isError ? 'status-error' : 'status-success';
  };

  const verifyPersonalToken = async ({ hostUri, token }) => {
    if (!token) return false;
    const normalizedHost = `${(hostUri || hostUriRef.current || '').replace(/\/+$/, '')}/`;
    if (!normalizedHost) return false;
    try {
      const res = await fetch(
        `${normalizedHost}_apis/connectionData?connectOptions=1&lastChangeId=-1&lastChangeId64=-1`,
        {
          headers: { Authorization: getAuthHeader(token) }
        }
      );
      return res.ok;
    } catch (error) {
      console.error('Failed to verify personal access token', error);
      return false;
    }
  };

  const requireVerifiedToken = async ({ hostUri }) => {
    const { token, savedAt } = readStoredToken();
    if (!token) {
      hidePipelineForm();
      setTokenStatus('A personal access token is required to continue.', true);
      setTokenGateStatus('Please add a personal access token in Settings to unlock the form.', true);
      openSettingsPanel();
      return null;
    }

    setTokenGateStatus('Verifying saved personal access token...');
    const verified = await verifyPersonalToken({ hostUri, token });
    if (!verified) {
      hidePipelineForm();
      setTokenStatus('The saved personal access token could not be verified. Please update it in Settings.', true);
      setTokenGateStatus('Token verification failed. Replace it in Settings to continue.', true);
      openSettingsPanel();
      return null;
    }

    const savedMessage = savedAt ? ` (saved ${new Date(savedAt).toLocaleString()})` : '';
    setTokenStatus(`Personal access token verified${savedMessage}.`);
    setTokenGateStatus('');
    showPipelineForm();
    closeSettingsPanel();
    return token;
  };

  const wireSettingsForm = ({ hostUriRef, getOnTokenUpdated }) => {
    const toggleSettings = () => {
      const isOpen = !settingsPanel?.classList.contains('hidden');
      if (isOpen) {
        closeSettingsPanel();
      } else {
        openSettingsPanel();
      }
    };

    openSettingsButton?.addEventListener('click', toggleSettings);
    closeSettingsButton?.addEventListener('click', closeSettingsPanel);
    openSettingsFromGateButton?.addEventListener('click', () => {
      openSettingsPanel();
      tokenInput?.focus();
    });

    settingsForm?.addEventListener('submit', async (event) => {
      event.preventDefault();
      const hostUri = hostUriRef?.current;
      const value = tokenInput?.value.trim();
      if (!value) {
        setTokenStatus('Enter a personal access token to continue.', true);
        hidePipelineForm();
        setTokenGateStatus('Personal access token required.', true);
        openSettingsPanel();
        return;
      }

      setTokenStatus('Verifying personal access token...');
      setTokenGateStatus('Verifying personal access token...');
      const verified = await verifyPersonalToken({ hostUri, token: value });
      if (!verified) {
        setTokenStatus('Token is invalid or lacks required permissions.', true);
        setTokenGateStatus('Token verification failed. Update it to proceed.', true);
        hidePipelineForm();
        openSettingsPanel();
        return;
      }

      const saved = persistToken(value);
      if (saved && tokenInput) {
        tokenInput.value = '';
        tokenInput.type = 'password';
      }
      if (!saved) {
        setTokenStatus('Unable to save token. Check browser storage permissions.', true);
        setTokenGateStatus('Unable to save token. Check browser storage permissions.', true);
        return;
      }

      setTokenStatus('Token verified and saved.');
      setTokenGateStatus('');
      showPipelineForm();
      closeSettingsPanel();
      const handler = getOnTokenUpdated?.();
      handler?.(value);
    });

    clearTokenButton?.addEventListener('click', () => {
      const handler = getOnTokenUpdated?.();
      const cleared = clearStoredToken();
      if (tokenInput) {
        tokenInput.value = '';
        tokenInput.type = 'password';
      }
      hidePipelineForm();
      setTokenGateStatus('Personal access token is required to continue.', true);
      setTokenStatus(cleared ? 'Saved token cleared. Add a new one to continue.' : 'Unable to clear stored token.', !cleared);
      openSettingsPanel();
      handler?.(null);
    });
  };

  const sanitizeProjectName = (name) => name.replace(/[^A-Za-z0-9]/g, '_');

  const setServiceNameFromRepository = (name) => {
    if (!serviceInput || !name) return;
    const normalized = name.toString().trim().toLowerCase();
    if (normalized) {
      serviceInput.value = normalized;
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
      }
    }
  };

  const getAuthHeader = (token) => {
    const tokenValue = typeof token === 'string' ? token : token?.token;
    if (!tokenValue) {
      throw new Error('Extension access token was unavailable.');
    }
    const encoded = btoa(`:${tokenValue}`);
    return `Basic ${encoded}`;
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
    const branchName = branch?.replace(/^refs\/heads\//, '') || 'main';
    const refUrl = `${hostUri}${encodeURIComponent(projectId)}/_apis/git/repositories/${repoId}/refs?filter=${encodeURIComponent(
      `heads/${branchName}`
    )}&api-version=7.1-preview.1`;
    const authHeader = getAuthHeader(accessToken);
    const res = await fetch(refUrl, { headers: { Authorization: authHeader } });

    if (res.status === 404) {
      return '0000000000000000000000000000000000000000';
    }

    if (!res.ok) {
      throw new Error(`Failed to query branch (${res.status})`);
    }

    const payload = await res.json();
    return payload.value?.[0]?.objectId || '0000000000000000000000000000000000000000';
  };

  const ensureRepo = async ({ hostUri, projectId, projectName, accessToken }) => {
    const sanitized = sanitizeProjectName(projectName);
    const targetName = `${sanitized}_Azure_DevOps`;
    targetRepoInput.value = targetName;
    const url = `${hostUri}${encodeURIComponent(projectId)}/_apis/git/repositories?api-version=7.1-preview.1`;

    const authHeader = getAuthHeader(accessToken);
    const res = await fetch(url, {
      headers: { Authorization: authHeader }
    });
    if (!res.ok) {
      throw new Error(`Failed to list repositories (${res.status})`);
    }
    const payload = await res.json();
    const existing = (payload.value || []).find((repo) => repo.name === targetName);
    if (existing) {
      return existing;
    }

    const createRes = await fetch(url, {
      method: 'POST',
      headers: {
        Authorization: authHeader,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({ name: targetName, project: { id: projectId } })
    });
    if (!createRes.ok) {
      throw new Error(`Failed to create repository (${createRes.status})`);
    }
    return createRes.json();
  };

  const postScaffold = async ({ hostUri, projectId, repoId, branch, accessToken, payload }) => {
    const branchName = branch?.replace(/^refs\/heads\//, '') || 'main';
    const branchRef = `refs/heads/${branchName}`;
    const url = `${hostUri}${encodeURIComponent(projectId)}/_apis/git/repositories/${repoId}/pushes?api-version=7.1-preview.1`;
    const content = `pool: ${payload.pool}\nservice: ${payload.service}\nenvironment: ${payload.environment}\ndockerfileDir: ${payload.dockerfileDir}\nrepositoryAddress: ${payload.repositoryAddress}\ncontainerRegistryService: ${payload.containerRegistryService}\nkomodoServer: ${payload.komodoServer}\n`;
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
              item: { path: '/pipeline-template.yml' },
              newContent: { content, contentType: 'rawtext' }
            }
          ]
        }
      ]
    };

    const res = await fetch(url, {
      method: 'POST',
      headers: {
        Authorization: authHeader,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(body)
    });

    if (!res.ok) {
      throw new Error(`Failed to push scaffold (${res.status})`);
    }
  };

  const fetchAgentQueues = async ({ hostUri, projectId, accessToken }) => {
    const url = `${hostUri}${encodeURIComponent(projectId)}/_apis/distributedtask/queues?api-version=7.1-preview.1`;
    const res = await fetch(url, { headers: { Authorization: getAuthHeader(accessToken) } });
    if (!res.ok) {
      throw new Error(`Failed to load pools (${res.status})`);
    }
    const payload = await res.json();
    return Array.from(new Set((payload.value || []).map((queue) => queue.name).filter(Boolean)));
  };

  const fetchContainerRegistries = async ({ hostUri, projectId, accessToken }) => {
    const url = `${hostUri}${encodeURIComponent(projectId)}/_apis/serviceendpoint/endpoints?type=dockerregistry&projectIds=${encodeURIComponent(
      projectId
    )}&api-version=7.1-preview.4`;
    const res = await fetch(url, { headers: { Authorization: getAuthHeader(accessToken) } });
    if (!res.ok) {
      throw new Error(`Failed to load container registries (${res.status})`);
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

  const fetchDockerfileDirectories = async ({ hostUri, projectId, repoId, branch, accessToken }) => {
    if (!repoId) return [];
    const versionDescriptor = branch
      ? `&versionDescriptor.version=${encodeURIComponent(branch)}&versionDescriptor.versionType=branch`
      : '';
    const url = `${hostUri}${encodeURIComponent(projectId)}/_apis/git/repositories/${repoId}/items?recursionLevel=Full&includeContentMetadata=true${versionDescriptor}&api-version=7.1-preview.1`;
    const res = await fetch(url, { headers: { Authorization: getAuthHeader(accessToken) } });
    if (!res.ok) {
      throw new Error(`Failed to scan repository for Dockerfiles (${res.status})`);
    }
    const payload = await res.json();
    return (payload.value || [])
      .filter((item) => !item.isFolder && /(?:^|\/|\\)Dockerfile$/i.test(item.path || item.serverItem || ''))
      .map((item) => normalizeDockerfileDir(item.path || item.serverItem))
      .filter(Boolean);
  };

  const init = async () => {
    populateDefaults();
    const query = new URLSearchParams(window.location.search);
    const branchFromQuery = getQueryValue(query.get('branch'));
    const projectIdFromQuery = getQueryValue(query.get('projectId'));
    const projectNameFromQuery = getQueryValue(query.get('projectName')) || projectIdFromQuery;
    const repoIdFromQuery = getQueryValue(query.get('repoId'));
    const repoNameFromQuery = getQueryValue(query.get('repoName'));
    const initialBranch = branchFromQuery || '(unknown branch)';

    branchLabel.textContent = branchFromQuery ? `Target branch: ${initialBranch}` : 'Loading branch context...';
    if (branchInput && branchFromQuery) {
      branchInput.value = initialBranch;
      branchInput.disabled = true;
    }
    targetRepoInput.value = `${sanitizeProjectName(projectNameFromQuery || 'project')}_Azure_DevOps`;
    setServiceNameFromRepository(repoNameFromQuery);
    applyDetectedEnvironment(initialBranch);

    wireSettingsForm({
      hostUriRef,
      getOnTokenUpdated: () => onTokenUpdatedRef.handler
    });

    try {
      await loadVssSdk();
      VSS.init({ usePlatformScripts: true, explicitNotifyLoaded: true });
      await VSS.ready();

      const context = VSS.getWebContext();

      const branch =
        branchFromQuery ||
        context?.repository?.defaultBranch?.replace(/^refs\/heads\//, '') ||
        '(unknown branch)';

      const projectId = projectIdFromQuery || context?.project?.id;
      const projectName = projectNameFromQuery || context?.project?.name || projectId;
      const repoId = repoIdFromQuery || context?.repository?.id;
      const repositoryName = repoNameFromQuery || context?.repository?.name;

      branchLabel.textContent = `Target branch: ${branch}`;
      if (branchInput) {
        branchInput.value = branch;
        branchInput.disabled = true;
      }
      targetRepoInput.value = `${sanitizeProjectName(projectName || 'project')}_Azure_DevOps`;
      setServiceNameFromRepository(repositoryName);
      applyDetectedEnvironment(branch);

      if (!projectId) {
        setStatus('Project context was not provided by the branch action or hub.', true);
        VSS.notifyLoadFailed('Missing project context');
        return;
      }

      const hostUri = (context.collection?.uri || getHostBase()).replace(/\/+$/, '') + '/';
      hostUriRef.current = hostUri;
      let accessToken;
      let cachedDockerfiles = [];

      const loadPools = async () => {
        if (!poolSelect || !accessToken) return;
        try {
          const poolNames = await fetchAgentQueues({ hostUri, projectId, accessToken });
          const poolOptionNames = mergeWithDefaults(defaultPoolOptions, poolNames);
          const options = poolOptionNames.map((name) => ({
            value: name,
            label: name
          }));
          populateSelectOptions(poolSelect, options, options.length ? 'Select a pool' : 'No accessible pools found');
          if (defaultValues.pool && options.some((option) => option.value === defaultValues.pool)) {
            poolSelect.value = defaultValues.pool;
          }
        } catch (error) {
          console.error(error);
          const fallbackOptions = defaultPoolOptions.map((name) => ({ value: name, label: name }));
          populateSelectOptions(poolSelect, fallbackOptions, 'Unable to load pools');
          poolSelect.value = defaultValues.pool;
        }
      };

      const loadContainerRegistries = async () => {
        if (!registrySelect || !accessToken) return;
        try {
          const registries = await fetchContainerRegistries({ hostUri, projectId, accessToken });
          const registryOptionNames = mergeWithDefaults(defaultRegistryOptions, registries);
          const options = registryOptionNames.map((name) => ({
            value: name,
            label: name
          }));
          populateSelectOptions(
            registrySelect,
            options,
            options.length ? 'Select a container registry service connection' : 'No container registries found'
          );
          if (
            defaultValues.containerRegistryService &&
            options.some((option) => option.value === defaultValues.containerRegistryService)
          ) {
            registrySelect.value = defaultValues.containerRegistryService;
          }
        } catch (error) {
          console.error(error);
          const fallbackOptions = defaultRegistryOptions.map((name) => ({ value: name, label: name }));
          populateSelectOptions(registrySelect, fallbackOptions, 'Unable to load container registries');
          registrySelect.value = defaultValues.containerRegistryService;
        }
      };

      const refreshDockerfiles = async () => {
        if (!dockerfileInput || !accessToken) return;
        dockerfileInput.value = '';
        try {
          cachedDockerfiles = await fetchDockerfileDirectories({ hostUri, projectId, repoId, branch, accessToken });
          if (cachedDockerfiles.length) {
            const defaultPath = cachedDockerfiles[0];
            dockerfileInput.value = defaultPath;
          } else {
            dockerfileInput.value = '';
            setStatus('No Dockerfile was found in this branch. Please provide the directory manually.', true);
          }
        } catch (error) {
          console.error(error);
          dockerfileInput.value = '';
          setStatus('Could not auto-detect Dockerfile location. Please fill it manually.', true);
        }
      };

      const initializeData = async () => Promise.all([loadPools(), loadContainerRegistries(), refreshDockerfiles()]);

      const applyAccessToken = async (token) => {
        accessToken = token;
        await initializeData();
      };

      onTokenUpdatedRef.handler = async (token) => {
        if (!token) {
          accessToken = undefined;
          setStatus('Add a personal access token to continue.', true);
          return;
        }
        await applyAccessToken(token);
      };

      const verifiedToken = await requireVerifiedToken({ hostUri });
      if (!verifiedToken) {
        setStatus('Add a personal access token in Settings to continue.', true);
        VSS.notifyLoadFailed('Personal access token missing or invalid');
        return;
      }

      await applyAccessToken(verifiedToken);

      form.addEventListener('submit', async (event) => {
        event.preventDefault();
        setStatus('Working on repository...');
        form.querySelector('button[type="submit"]').disabled = true;
        const payload = Object.fromEntries(new FormData(form).entries());

        if (!accessToken) {
          setStatus('A verified personal access token is required. Open Settings to update it.', true);
          hidePipelineForm();
          openSettingsPanel();
          form.querySelector('button[type="submit"]').disabled = false;
          return;
        }

        try {
          const repo = await ensureRepo({ hostUri, projectId, projectName, accessToken });
          const defaultBranch = repo.defaultBranch?.replace(/^refs\/heads\//, '') || branch || 'main';
          await postScaffold({ hostUri, projectId, repoId: repo.id, branch: defaultBranch, accessToken, payload });
          setStatus(`Repository ${repo.name} is ready with pipeline template on ${defaultBranch}.`, false);
        } catch (error) {
          console.error(error);
          setStatus(error.message, true);
        }

        form.querySelector('button[type="submit"]').disabled = false;
      });

      VSS.notifyLoadSucceeded();
    } catch (error) {
      console.error('Failed to initialize extension frame', error);
      setStatus('Failed to initialize extension frame. Check extension permissions and reload.', true);
      const sdk = normalizeSdk(window.VSS || window.parent?.VSS);
      sdk?.notifyLoadFailed?.(error?.message || 'Initialization failed');
      setTokenGateStatus('Failed to initialize extension frame. Open settings after fixing the issue.', true);
      openSettingsPanel();
    }
  };

  init();
})();
