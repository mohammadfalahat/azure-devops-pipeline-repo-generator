(() => {
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

  const ensureRepo = async ({ hostUri, projectId, projectName, accessToken }) => {
    const sanitized = sanitizeProjectName(projectName);
    const targetName = `${sanitized}_Azure_DevOps`;
    targetRepoInput.value = targetName;
    const url = `${hostUri}${encodeURIComponent(projectId)}/_apis/git/repositories?api-version=7.1-preview.1`;

    const res = await fetch(url, {
      headers: { Authorization: `Bearer ${accessToken}` }
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
        Authorization: `Bearer ${accessToken}`,
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
        Authorization: `Bearer ${accessToken}`,
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
      await SDK.init({ loaded: false });
      const context = await SDK.ready();
      const query = new URLSearchParams(window.location.search);
      const branch = query.get('branch') || '(unknown branch)';
      const projectId = query.get('projectId');
      const projectName = query.get('projectName') || projectId;
      const repoId = query.get('repoId');

      branchLabel.textContent = `Target branch: ${branch}`;
      targetRepoInput.value = `${sanitizeProjectName(projectName || 'project')}_Azure_DevOps`;

      if (!projectId) {
        setStatus('Project context was not provided by the branch action.', true);
        SDK.notifyLoadFailed('Missing project context');
        return;
      }

      const host = SDK.getHost();
      const accessToken = await SDK.getAccessToken();
      const hostUri = host ? `${host.uri}/` : `${context.webContext.collection.uri}`;

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

      SDK.notifyLoadSucceeded();
    } catch (error) {
      console.error('Failed to initialize extension frame', error);
      setStatus('Failed to initialize extension frame. Check extension permissions and reload.', true);
      SDK.notifyLoadFailed(error?.message || 'Initialization failed');
    }
  };

  init();
})();
