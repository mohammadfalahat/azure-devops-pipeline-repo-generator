const getHostBase = () => {
  if (!document.referrer) {
    return window.location.origin;
  }
  const referrer = new URL(document.referrer);
  const segments = referrer.pathname.split('/').filter(Boolean);
  const hasTfsVirtualDir = segments[0]?.toLowerCase() === 'tfs';
  return `${referrer.origin}${hasTfsVirtualDir ? '/tfs' : ''}`;
};

const loadScript = async (src) => {
  let contentTypeError;
  try {
    const response = await fetch(src, { credentials: 'include', cache: 'no-cache', redirect: 'follow' });
    if (!response.ok) {
      throw new Error(`HTTP ${response.status}`);
    }

    const contentType = response.headers.get('content-type')?.toLowerCase() || '';
    const isJavaScript = /javascript|ecmascript|ms-vssweb/.test(contentType);
    if (contentType && !isJavaScript) {
      contentTypeError = new Error(`Unexpected content type ${contentType}`);
      throw contentTypeError;
    }
  } catch (error) {
    if (contentTypeError) {
      throw new Error(`Blocked Azure DevOps SDK from ${src}: ${contentTypeError.message}`);
    }
    console.warn('Skipping preflight validation failure, trying script tag load next', src, error);
  }

  return new Promise((resolve, reject) => {
    const script = document.createElement('script');
    script.src = src;
    script.async = false;
    script.crossOrigin = 'use-credentials';
    script.onload = resolve;
    script.onerror = () => reject(new Error(`Failed to load Azure DevOps SDK from ${src}`));
    document.head.appendChild(script);
  });
};

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

const waitForSdkReady = async (sdk) => {
  if (!sdk?.ready) {
    return;
  }

  await new Promise((resolve, reject) => {
    try {
      sdk.ready(resolve);
    } catch (error) {
      reject(error);
    }
  });
};

const warmupAssets = async () => {
  const assets = [
    new URL('./index.html', window.location.href).toString(),
    new URL('./ui.js', window.location.href).toString(),
    new URL('./styles.css', window.location.href).toString(),
    new URL('./lib/VSS.SDK.min.js', window.location.href).toString()
  ];

  await Promise.all(
    assets.map(async (href) => {
      try {
        const res = await fetch(href, { cache: 'force-cache', mode: 'no-cors' });
        if (!res || (res.type === 'opaque' ? false : !res.ok)) {
          return;
        }
      } catch (error) {
        console.warn('Failed to warm up asset', href, error);
      }
    })
  );
};

const prefetchResources = () => {
  // Preload only assets with a supported `as` value to avoid browser warnings.
  const resources = [
    { href: new URL('./ui.js', window.location.href).toString(), as: 'script' },
    { href: new URL('./styles.css', window.location.href).toString(), as: 'style' },
    { href: new URL('./lib/VSS.SDK.min.js', window.location.href).toString(), as: 'script' }
  ];

  resources.forEach(({ href, as }) => {
    const preload = document.createElement('link');
    preload.rel = 'preload';
    preload.as = as || 'script';
    preload.href = href;
    preload.crossOrigin = 'use-credentials';
    document.head.appendChild(preload);

    const prefetch = document.createElement('link');
    prefetch.rel = 'prefetch';
    prefetch.as = as || 'script';
    prefetch.href = href;
    prefetch.crossOrigin = 'use-credentials';
    document.head.appendChild(prefetch);
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

const delay = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

const normalizeAccessTokenError = (error) => {
  const message = error?.message || 'Unknown Azure DevOps authentication error';
  if (/HostAuthorizationNotFound/i.test(message)) {
    return 'Host authorization was not found. Ensure the extension is enabled for this collection/project and that your account can access it.';
  }
  return message;
};

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

      // Some Azure DevOps Server instances may respond with an internal
      // error (HTTP 500) from the WebPlatformAuth SessionToken endpoint
      // when an access token cannot be issued. Retrying those responses
      // only generates more noisy 500 logs without succeeding, so stop
      // immediately and surface the error to the UI instead.
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

const postBootstrapMessage = (targetWindow, targetOrigin, payload) => {
  if (!targetWindow || !targetOrigin) return;

  const message = { type: 'pipeline-bootstrap', payload };
  let attempts = 0;
  const maxAttempts = 10;

  const intervalId = setInterval(() => {
    attempts += 1;
    if (targetWindow.closed || attempts > maxAttempts) {
      clearInterval(intervalId);
      return;
    }
    try {
      targetWindow.postMessage(message, targetOrigin);
    } catch (error) {
      console.warn('Failed to post bootstrap message to generator window', error);
    }
  }, 400);

  const handleAck = (event) => {
    if (event.source === targetWindow && event.origin === targetOrigin && event.data?.type === 'pipeline-bootstrap-ack') {
      clearInterval(intervalId);
      window.removeEventListener('message', handleAck);
    }
  };

  window.addEventListener('message', handleAck);
};

const openGenerator = async (context, sdk) => {
  try {
    const actionContext = getActionContext(context);
    const repository = getRepository(actionContext) || VSS.getWebContext?.()?.repository;
    const branchName = getBranchName(actionContext);
    const project = getProject(actionContext);
    const repoId = repository?.id;
    const repoName = repository?.name;
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
    if (repoName) params.set('repoName', repoName);

    const targetUrl = `${baseUri}dist/index.html?${params.toString()}`;
    const targetOrigin = new URL(targetUrl).origin;

    const hostUri = (VSS.getWebContext?.()?.collection?.uri || getHostBase()).replace(/\/+$/, '') + '/';
    let accessToken;
    let accessTokenError;
    try {
      accessToken = await getAccessTokenWithRetry(sdk);
    } catch (tokenError) {
      console.warn('Could not acquire access token before opening generator', tokenError);
      accessTokenError = normalizeAccessTokenError(tokenError);
    }

    const bootstrapPayload = {
      branch: branchName,
      projectId,
      projectName,
      repoId,
      repoName,
      hostUri,
      accessToken,
      accessTokenError
    };

    try {
      const hostService = await VSS.getService(VSS.ServiceIds.HostPageLayout);
      if (hostService?.openWindow) {
        const generatorWindow = hostService.openWindow(targetUrl, {});
        postBootstrapMessage(generatorWindow, targetOrigin, bootstrapPayload);
        return;
      }
    } catch (serviceError) {
      console.warn('Falling back to window.open because HostPageLayout was unavailable', serviceError);
    }

    const generatorWindow = window.open(targetUrl, '_blank');
    postBootstrapMessage(generatorWindow, targetOrigin, bootstrapPayload);
  } catch (error) {
    console.error('Failed to launch pipeline generator', error);
    VSS.handleError?.(error);
  }
};

const initializeAction = () => {
  let sdkInitPromise;
  let assetWarmupPromise;

  const ensureSdkReady = () => {
    if (!sdkInitPromise) {
      sdkInitPromise = (async () => {
        const sdk = await loadVssSdk();
        sdk.init({ usePlatformScripts: true, explicitNotifyLoaded: true });
        await waitForSdkReady(sdk);
        prefetchResources();
        if (!assetWarmupPromise) {
          assetWarmupPromise = warmupAssets().catch((error) => {
            console.warn('Asset warmup failed', error);
            return undefined;
          });
        }
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

  const registerAction = async () => {
    let readySdk;
    try {
      readySdk = await ensureSdkReady();
    } catch (error) {
      console.error('Failed to initialize SDK before registering action', error);
      return false;
    }

    const sdk = normalizeSdk(window.VSS || window.parent?.VSS || readySdk);
    if (!sdk?.register) {
      console.error('Azure DevOps SDK did not expose register after initialization');
      return false;
    }

    const action = {
      execute: async (context) => {
        const sdkInstance = await ensureSdkReady();
        if (assetWarmupPromise) {
          await assetWarmupPromise.catch(() => {});
        }
        await openGenerator(context, sdkInstance);
        sdkInstance.notifyLoadSucceeded?.();
      }
    };

    try {
      sdk.register('generate-pipeline-action', action);
      readySdk.notifyLoadSucceeded?.();
    } catch (error) {
      console.error('Failed to register branch action', error);
      sdk.notifyLoadFailed?.(error?.message || 'Registration failed');
      return false;
    }

    return true;
  };

  const startRetryLoop = () => {
    const intervalId = setInterval(() => {
      registerAction().then((registered) => {
        if (registered) {
          clearInterval(intervalId);
        }
      });
    }, 50);
  };

  registerAction()
    .then((registered) => {
      if (!registered) {
        startRetryLoop();
      }
    })
    .catch(() => startRetryLoop());
};

initializeAction();
