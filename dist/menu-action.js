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
  Boolean(sdk && sdk.init && sdk.ready && sdk.getService && (sdk.getWebContext || sdk.getHostContext));

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
    new URL('./lib/VSS.SDK.min.js', window.location.href).toString(),
    `${getHostBase()}/_content/MS.VSS.SDK/scripts/VSS.SDK.min.js`
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

const prefetchResources = () => {
  const resources = [
    new URL('./index.html', window.location.href).toString(),
    new URL('./ui.js', window.location.href).toString(),
    new URL('./styles.css', window.location.href).toString(),
    new URL('./lib/VSS.SDK.min.js', window.location.href).toString()
  ];

  resources.forEach((href) => {
    const link = document.createElement('link');
    link.rel = 'prefetch';
    link.href = href;
    document.head.appendChild(link);
  });
};

const normalizeBranchName = (name) => {
  if (!name) {
    return undefined;
  }

  const withoutVersionPrefix = name.startsWith('GB') ? name.slice(2) : name;
  return withoutVersionPrefix.replace(/^refs\/heads\//i, '');
};

const getActionContext = (context) => {
  const configuration = VSS.getConfiguration?.();
  return configuration?.actionContext || context || {};
};

const getProject = (context) =>
  context?.project ||
  getRepository(context)?.project ||
  VSS.getWebContext?.()?.project;

const getRepository = (context) =>
  context?.gitRepository ||
  context?.repository ||
  context?.item?.repository ||
  context?.branch?.repository ||
  context?.gitRef?.repository ||
  context?.selectedItem?.repository;

const getBranchName = (context) => {
  const actionContext = getActionContext(context);
  const branchCandidates = [
    actionContext?.branch?.name,
    actionContext?.branch?.fullName,
    actionContext?.branch?.refName,
    actionContext?.gitRef?.name,
    actionContext?.gitRef?.fullName,
    actionContext?.gitRef?.refName,
    actionContext?.ref?.name,
    actionContext?.ref?.refName,
    actionContext?.refName,
    actionContext?.selectedItem?.refName,
    actionContext?.selectedItem?.name,
    actionContext?.item?.refName,
    actionContext?.item?.name,
    actionContext?.branchName,
    actionContext?.name,
    actionContext?.fullName
  ];

  const branchFromContext = branchCandidates.find((value) => typeof value === 'string' && value.trim().length > 0);
  if (branchFromContext) {
    return normalizeBranchName(branchFromContext.trim());
  }

  const url = new URL(window.location.href);
  const version = url.searchParams.get('version');
  const branchFromQuery = normalizeBranchName(version);
  if (branchFromQuery) {
    return branchFromQuery;
  }

  const fallbackBranch = getRepository(actionContext)?.defaultBranch;
  const branchFromWebContext = VSS.getWebContext?.()?.repository?.defaultBranch;
  return normalizeBranchName(fallbackBranch) || normalizeBranchName(branchFromWebContext) || 'Unknown branch';
};

const openGenerator = async (context) => {
  try {
    const actionContext = getActionContext(context);
    const repository = getRepository(actionContext) || VSS.getWebContext?.()?.repository;
    const branchName = getBranchName(actionContext);
    const project = getProject(actionContext);
    const repoId = repository?.id;
    const projectId = project?.id || actionContext?.projectId;
    const projectName = project?.name || projectId;
    const extContext = VSS.getExtensionContext?.();
    const baseUriCandidate = extContext?.baseUri || `${getHostBase()}/`;
    const baseUri = baseUriCandidate.endsWith('/') ? baseUriCandidate : `${baseUriCandidate}/`;
    const params = new URLSearchParams();

    if (branchName) params.set('branch', branchName);
    if (projectId) params.set('projectId', projectId);
    if (projectName) params.set('projectName', projectName);
    if (repoId) params.set('repoId', repoId);

    const targetUrl = `${baseUri}dist/index.html?${params.toString()}`;

    try {
      const hostService = await VSS.getService(VSS.ServiceIds.HostPageLayout);
      if (hostService?.openWindow) {
        hostService.openWindow(targetUrl, {});
        return;
      }
    } catch (serviceError) {
      console.warn('Falling back to window.open because HostPageLayout was unavailable', serviceError);
    }

    window.open(targetUrl, '_blank');
  } catch (error) {
    console.error('Failed to launch pipeline generator', error);
    VSS.handleError?.(error);
  }
};

const initializeAction = () => {
  let sdkInitPromise;

  const ensureSdkReady = () => {
    if (!sdkInitPromise) {
      sdkInitPromise = (async () => {
        const sdk = await loadVssSdk();
        sdk.init({ usePlatformScripts: true, explicitNotifyLoaded: true });
        await sdk.ready();
        prefetchResources();
        sdk.notifyLoadSucceeded();
        return sdk;
      })().catch((error) => {
        console.error('Failed to initialize branch action', error);
        const fallbackSdk = normalizeSdk(window.VSS || window.parent?.VSS);
        fallbackSdk?.notifyLoadFailed?.(error?.message || 'Initialization failed');
        throw error;
      });
    }
    return sdkInitPromise;
  };

  const registerAction = () => {
    const sdk = normalizeSdk(window.VSS || window.parent?.VSS);
    if (!sdk?.register) {
      return false;
    }
    sdk.register('generate-pipeline-action', {
      execute: async (context) => {
        await ensureSdkReady();
        await openGenerator(context);
      }
    });
    return true;
  };

  if (!registerAction()) {
    const intervalId = setInterval(() => {
      if (registerAction()) {
        clearInterval(intervalId);
      }
    }, 50);
  }

  ensureSdkReady().catch(() => {});
};

initializeAction();
