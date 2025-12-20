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

  const gallerySdkUrl =
    'https://azure.buluttakin.com/_apis/public/gallery/publisher/localdev/extension/pipeline-generator/0.1.10/assetbyname/dist/lib/VSS.SDK.min.js';

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
    const candidates = [gallerySdkUrl, hostSdk, localSdk, localSdkFallback];

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

  const waitForSdkReady = async (sdk, timeoutMs = 5000) => {
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
  const previewPanel = document.getElementById('preview-panel');
  const yamlOutput = document.getElementById('yaml-output');
  const copyYamlButton = document.getElementById('copy-yaml');
  const submitButton = form?.querySelector('button[type="submit"]');

  const state = {
    sdk: null,
    accessToken: null,
    hostUri: null,
    projectId: null,
    projectName: null,
    repoId: null,
    repositoryName: null,
    branch: null
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

  const getAccessTokenFromSdk = async (sdk) => {
    if (!sdk?.getAccessToken) {
      throw new Error('Azure DevOps access token API is unavailable.');
    }
    const token = await sdk.getAccessToken();
    if (!token) {
      throw new Error('Azure DevOps did not provide an access token.');
    }
    return token;
  };

  const sanitizeProjectName = (name) => name.replace(/[^A-Za-z0-9]/g, '_');

  const slugifyName = (value, fallback) => {
    const slug = value?.toString().trim().toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, '');
    return slug || fallback;
  };

  const buildPipelineFilename = ({ projectName, repositoryName, environment }) => {
    const projectSlug = slugifyName(projectName, 'project');
    const repoSlug = slugifyName(repositoryName || projectName, 'repo');
    const environmentSlug = slugifyName(environment, 'env');
    return `${projectSlug}-${repoSlug}-${environmentSlug}.yml`;
  };

  const applyBootstrapPayload = async (payload = {}, source = 'message') => {
    const {
      branch,
      projectId,
      projectName,
      repoId,
      repoName,
      hostUri,
      accessToken
    } = payload;

    const normalizedHost = (hostUri || state.hostUri || getHostBase()).replace(/\/+$/, '') + '/';
    state.branch = branch || state.branch;
    state.projectId = projectId || state.projectId;
    state.projectName = projectName || state.projectName;
    state.repoId = repoId || state.repoId;
    state.repositoryName = repoName || state.repositoryName;
    state.hostUri = normalizedHost;
    state.accessToken = accessToken || state.accessToken;

    const targetBranch = state.branch;
    branchLabel.textContent = targetBranch ? `Target branch: ${targetBranch}` : 'Loading branch context...';
    if (branchInput && targetBranch) {
      branchInput.value = targetBranch;
      branchInput.disabled = true;
    }

    targetRepoInput.value = `${sanitizeProjectName(state.projectName || 'project')}_Azure_DevOps`;
    if (!serviceInput.value) {
      setServiceNameFromRepository(state.repositoryName || state.projectName);
    }
    applyDetectedEnvironment(targetBranch);
    setKomodoServerFromEnvironment(environmentSelect?.value);

    if (!state.projectId || !state.accessToken || !state.hostUri) {
      setStatus('Loaded context from branch action. Waiting for Azure DevOps host to provide an access token...', true);
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

  const setServiceNameFromRepository = (name) => {
    if (!serviceInput || !name) return;
    const normalized = name.toString().trim().toLowerCase();
    if (normalized) {
      serviceInput.value = normalized;
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
    const branchName = branch?.replace(/^refs\/heads\//, '') || 'main';
    const refUrl = `${hostUri}${encodeURIComponent(projectId)}/_apis/git/repositories/${repoId}/refs?filter=${encodeURIComponent(
      `heads/${branchName}`
    )}&api-version=7.1-preview.1`;
    const res = await fetch(refUrl, { headers: authHeaders(accessToken) });

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

    const res = await fetch(url, {
      headers: authHeaders(accessToken)
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
        ...authHeaders(accessToken),
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({ name: targetName, project: { id: projectId } })
    });
    if (!createRes.ok) {
      throw new Error(`Failed to create repository (${createRes.status})`);
    }
    return createRes.json();
  };

  const postScaffold = async ({
    hostUri,
    projectId,
    repoId,
    branch,
    accessToken,
    content,
    pipelineFilename = 'pipeline-template.yml'
  }) => {
    const branchName = branch?.replace(/^refs\/heads\//, '') || 'main';
    const branchRef = `refs/heads/${branchName}`;
    const url = `${hostUri}${encodeURIComponent(projectId)}/_apis/git/repositories/${repoId}/pushes?api-version=7.1-preview.1`;
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
      throw new Error(`Failed to push scaffold (${res.status})`);
    }
  };

  const fetchAgentQueues = async ({ hostUri, projectId, accessToken }) => {
    const url = `${hostUri}${encodeURIComponent(projectId)}/_apis/distributedtask/queues?api-version=7.1-preview.1`;
    const res = await fetch(url, { headers: authHeaders(accessToken) });
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
    const res = await fetch(url, { headers: authHeaders(accessToken) });
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

  const buildPipelineYaml = (payload) =>
    [
      `pool: ${payload.pool || ''}`,
      `service: ${payload.service || ''}`,
      `environment: ${payload.environment || ''}`,
      `dockerfileDir: ${payload.dockerfileDir || ''}`,
      `repositoryAddress: ${payload.repositoryAddress || ''}`,
      `containerRegistryService: ${payload.containerRegistryService || ''}`,
      `komodoServer: ${payload.komodoServer || ''}`,
      ''
    ].join('\n');

  const showYamlPreview = (payload) => {
    if (!yamlOutput) return '';
    const yaml = buildPipelineYaml(payload);
    yamlOutput.textContent = yaml;
    previewPanel?.classList.remove('hidden');
    return yaml;
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
    const yaml = showYamlPreview(payload);

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
      state.projectName = context?.project?.name || state.projectId || state.projectName;
      state.repoId = context?.repository?.id || state.repoId;
      state.repositoryName = context?.repository?.name || state.repositoryName;
    }

    const pipelineFilename = buildPipelineFilename({
      projectName: state.projectName,
      repositoryName: state.repositoryName,
      environment: payload.environment
    });

    if (!state.accessToken || !state.projectId) {
      setStatus('Template generated below. Open the extension from Azure DevOps to save it automatically.', true);
      setSubmitting(false);
      return yaml;
    }

    try {
      const repo = await ensureRepo({
        hostUri: state.hostUri,
        projectId: state.projectId,
        projectName: state.projectName,
        accessToken: state.accessToken
      });
      const defaultBranch = repo.defaultBranch?.replace(/^refs\/heads\//, '') || state.branch || 'main';
      await postScaffold({
        hostUri: state.hostUri,
        projectId: state.projectId,
        repoId: repo.id,
        branch: defaultBranch,
        accessToken: state.accessToken,
        content: yaml,
        pipelineFilename
      });
      setStatus(
        `Repository ${repo.name} is ready with ${pipelineFilename} on ${defaultBranch}.`,
        false
      );
    } catch (error) {
      console.error(error);
      setStatus(`Template generated below, but automatic push failed: ${error.message}`, true);
    }

    setSubmitting(false);
    return yaml;
  };

  form?.addEventListener('submit', handleSubmit);

  copyYamlButton?.addEventListener('click', async () => {
    if (!yamlOutput?.textContent) return;
    try {
      await navigator.clipboard?.writeText(yamlOutput.textContent);
      setStatus('YAML copied to clipboard.');
    } catch (error) {
      console.error(error);
      setStatus('Could not copy YAML. Please copy it manually from the preview.', true);
    }
  });

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
    const shouldAttemptSdk = hasReferrer || isFramed || hasOpener;

    state.branch = initialBranch;
    state.projectId = projectIdFromQuery;
    state.projectName = projectNameFromQuery;
    state.repoId = repoIdFromQuery;
    state.repositoryName = repoNameFromQuery;
    state.hostUri = `${getHostBase().replace(/\/+$/, '')}/`;

    branchLabel.textContent = branchFromQuery ? `Target branch: ${initialBranch}` : 'Loading branch context...';
    if (branchInput && branchFromQuery) {
      branchInput.value = initialBranch;
      branchInput.disabled = true;
    }
    targetRepoInput.value = `${sanitizeProjectName(projectNameFromQuery || 'project')}_Azure_DevOps`;
    setServiceNameFromRepository(repoNameFromQuery || projectNameFromQuery);
    applyDetectedEnvironment(initialBranch);
    const hasHostContext = Boolean(hasReferrer || projectIdFromQuery || repoIdFromQuery || hasOpener);
    if (!hasHostContext || !shouldAttemptSdk) {
      setStatus(
        'Running outside Azure DevOps. Fill the form to preview the YAML, then copy it below. Open the extension from a branch action to enable automatic push.',
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
      state.branch = branch;

      const projectId = projectIdFromQuery || context?.project?.id;
      const projectName = projectNameFromQuery || context?.project?.name || projectId;
      const repoId = repoIdFromQuery || context?.repository?.id;
      const repositoryName = repoNameFromQuery || context?.repository?.name;
      state.sdk = sdk;
      state.projectId = projectId;
      state.projectName = projectName;
      state.repoId = repoId;
      state.repositoryName = repositoryName;

      branchLabel.textContent = `Target branch: ${branch}`;
      if (branchInput) {
        branchInput.value = branch;
        branchInput.disabled = true;
      }
      targetRepoInput.value = `${sanitizeProjectName(projectName || 'project')}_Azure_DevOps`;
      if (!serviceInput.value) {
        setServiceNameFromRepository(repositoryName || projectName);
      }
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
