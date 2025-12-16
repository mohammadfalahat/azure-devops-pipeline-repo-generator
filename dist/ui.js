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

    const candidates = [
      `${getHostBase()}/_content/MS.VSS.SDK/scripts/VSS.SDK.min.js`,
      new URL('./lib/VSS.SDK.min.js', window.location.href).toString()
    ];

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
    service: 'api',
    environment: 'demo',
    dockerfileDir: 'src/TMS.API',
    repositoryAddress: 'registry.buluttakin.com',
    containerRegistryService: 'BulutReg',
    komodoServer: 'DEMO-192.168.62.91'
  };

  const getQueryValue = (value) => (value && value !== 'undefined' && value !== 'null' ? value : undefined);

  const branchLabel = document.getElementById('branch-label');
  const branchInput = document.getElementById('branch');
  const environmentSelect = document.getElementById('environment');
  const poolSelect = document.getElementById('pool');
  const registrySelect = document.getElementById('containerRegistryService');
  const dockerfileInput = document.getElementById('dockerfileDir');
  const form = document.getElementById('pipeline-form');
  const status = document.getElementById('status');
  const targetRepoInput = document.getElementById('targetRepo');

  const setStatus = (message, isError = false) => {
    status.textContent = message;
    status.className = isError ? 'status-error' : 'status-success';
  };

  const sanitizeProjectName = (name) => name.replace(/[^A-Za-z0-9]/g, '_');

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
    return `Bearer ${tokenValue}`;
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

    try {
      await loadVssSdk();
      VSS.init({ usePlatformScripts: true, explicitNotifyLoaded: true });
      await VSS.ready();

      const context = VSS.getWebContext();
      const query = new URLSearchParams(window.location.search);

      const branchFromQuery = getQueryValue(query.get('branch'));
      const branch =
        branchFromQuery ||
        context?.repository?.defaultBranch?.replace(/^refs\/heads\//, '') ||
        '(unknown branch)';

      const projectId = getQueryValue(query.get('projectId')) || context?.project?.id;
      const projectName = getQueryValue(query.get('projectName')) || context?.project?.name || projectId;
      const repoId = getQueryValue(query.get('repoId')) || context?.repository?.id;

      branchLabel.textContent = `Target branch: ${branch}`;
      if (branchInput) {
        branchInput.value = branch;
        branchInput.disabled = true;
      }
      targetRepoInput.value = `${sanitizeProjectName(projectName || 'project')}_Azure_DevOps`;
      applyDetectedEnvironment(branch);

      if (!projectId) {
        setStatus('Project context was not provided by the branch action or hub.', true);
        VSS.notifyLoadFailed('Missing project context');
        return;
      }

      const hostUri = (context.collection?.uri || getHostBase()).replace(/\/+$/, '') + '/';
      let accessToken;

      try {
        accessToken = await VSS.getAccessToken();
      } catch (tokenError) {
        console.error('Access token request was rejected', tokenError);
        setStatus(
          'The extension was denied access to an Azure DevOps token. Ask an admin to approve extension permissions in Organization settings → Extensions.',
          true
        );
        VSS.notifyLoadFailed(tokenError?.message || 'Access token rejected');
        return;
      }

      const loadPools = async () => {
        if (!poolSelect) return;
        try {
          const poolNames = await fetchAgentQueues({ hostUri, projectId, accessToken });
          const options = poolNames.map((name) => ({ value: name, label: name }));
          if (!options.length && defaultValues.pool) {
            options.push({ value: defaultValues.pool, label: defaultValues.pool });
          }
          populateSelectOptions(poolSelect, options, options.length ? 'Select a pool' : 'No accessible pools found');
          if (defaultValues.pool && poolNames.includes(defaultValues.pool)) {
            poolSelect.value = defaultValues.pool;
          }
        } catch (error) {
          console.error(error);
          populateSelectOptions(poolSelect, [], 'Unable to load pools');
        }
      };

      const loadContainerRegistries = async () => {
        if (!registrySelect) return;
        try {
          const registries = await fetchContainerRegistries({ hostUri, projectId, accessToken });
          const options = registries.map((name) => ({ value: name, label: name }));
          if (!options.length && defaultValues.containerRegistryService) {
            options.push({ value: defaultValues.containerRegistryService, label: defaultValues.containerRegistryService });
          }
          populateSelectOptions(
            registrySelect,
            options,
            options.length ? 'Select a container registry service connection' : 'No container registries found'
          );
          if (defaultValues.containerRegistryService && registries.includes(defaultValues.containerRegistryService)) {
            registrySelect.value = defaultValues.containerRegistryService;
          }
        } catch (error) {
          console.error(error);
          populateSelectOptions(registrySelect, [], 'Unable to load container registries');
        }
      };

      let cachedDockerfiles = [];
      const refreshDockerfiles = async () => {
        if (!dockerfileInput) return;
        try {
          cachedDockerfiles = await fetchDockerfileDirectories({ hostUri, projectId, repoId, branch, accessToken });
          if (cachedDockerfiles.length) {
            const defaultPath = cachedDockerfiles[0];
            dockerfileInput.value = defaultPath;
          } else {
            setStatus('No Dockerfile was found in this branch. Please provide the directory manually.', true);
          }
        } catch (error) {
          console.error(error);
          setStatus('Could not auto-detect Dockerfile location. Please fill it manually.', true);
        }
      };

      await Promise.all([loadPools(), loadContainerRegistries(), refreshDockerfiles()]);

      form.addEventListener('submit', async (event) => {
        event.preventDefault();
        setStatus('Working on repository...');
        form.querySelector('button[type="submit"]').disabled = true;
        const payload = Object.fromEntries(new FormData(form).entries());

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
      VSS.notifyLoadFailed(error?.message || 'Initialization failed');
    }
  };

  init();
})();
