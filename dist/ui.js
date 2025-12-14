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

  const loadVssSdk = async () => {
    if (window.VSS) {
      return window.VSS;
    }

    const candidates = [
      new URL('./lib/VSS.SDK.min.js', window.location.href).toString(),
      `${getHostBase()}/_content/MS.VSS.SDK/scripts/VSS.SDK.min.js`
    ];

    let lastError;
    for (const src of candidates) {
      try {
        await loadScript(src);
        if (window.VSS) {
          return window.VSS;
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

  const branchLabel = document.getElementById('branch-label');
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
      if (input) {
        if (input.tagName.toLowerCase() === 'select') {
          input.value = value;
        } else {
          input.value = value;
        }
      }
    });
  };

  const getAuthHeader = (token) => {
    const tokenValue = typeof token === 'string' ? token : token?.token;
    if (!tokenValue) {
      throw new Error('Extension access token was unavailable.');
    }
    return `Bearer ${tokenValue}`;
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

  const postScaffold = async ({ hostUri, projectId, repoId, accessToken, payload }) => {
    const url = `${hostUri}${encodeURIComponent(projectId)}/_apis/git/repositories/${repoId}/pushes?api-version=7.1-preview.1`;
    const content = `pool: ${payload.pool}\nservice: ${payload.service}\nenvironment: ${payload.environment}\ndockerfileDir: ${payload.dockerfileDir}\nrepositoryAddress: ${payload.repositoryAddress}\ncontainerRegistryService: ${payload.containerRegistryService}\nkomodoServer: ${payload.komodoServer}\n`;
    const authHeader = getAuthHeader(accessToken);
    const body = {
      refUpdates: [
        {
          name: 'refs/heads/main',
          oldObjectId: '0000000000000000000000000000000000000000'
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

  const init = async () => {
    populateDefaults();

    try {
      await loadVssSdk();
      VSS.init({ usePlatformScripts: true, explicitNotifyLoaded: true });
      await VSS.ready();

      const context = VSS.getWebContext();
      const query = new URLSearchParams(window.location.search);
      const branch = query.get('branch') || '(unknown branch)';
      const projectId = query.get('projectId');
      const projectName = query.get('projectName') || projectId;
      const repoId = query.get('repoId');

      branchLabel.textContent = `Target branch: ${branch}`;
      targetRepoInput.value = `${sanitizeProjectName(projectName || 'project')}_Azure_DevOps`;

      if (!projectId) {
        setStatus('Project context was not provided by the branch action.', true);
        VSS.notifyLoadFailed('Missing project context');
        return;
      }

      const hostContext = VSS.getHostContext();
      const hostUri = (hostContext?.host?.uri || context.collection?.uri || '').replace(/\/+$/, '') + '/';
      let accessToken;

      try {
        accessToken = await VSS.getAccessToken();
      } catch (tokenError) {
        console.error('Access token request was rejected', tokenError);
        setStatus('The extension was denied access to an Azure DevOps token. Ask an admin to approve extension permissions in Organization settings → Extensions.', true);
        VSS.notifyLoadFailed(tokenError?.message || 'Access token rejected');
        return;
      }

      form.addEventListener('submit', async (event) => {
        event.preventDefault();
        setStatus('Working on repository...');
        form.querySelector('button[type="submit"]').disabled = true;
        const payload = Object.fromEntries(new FormData(form).entries());

        try {
          const repo = await ensureRepo({ hostUri, projectId, projectName, accessToken });
          await postScaffold({ hostUri, projectId, repoId: repo.id, accessToken, payload });
          setStatus(`Repository ${repo.name} is ready with pipeline template.`, false);
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
