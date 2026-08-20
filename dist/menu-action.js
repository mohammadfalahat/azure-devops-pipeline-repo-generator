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

const loadScript = async (src) => {
  let contentTypeError;
  try {
    const response = await fetch(src, { credentials: 'include', cache: 'no-cache', redirect: 'follow' });
    if (!response.ok) {
      const wwwAuthenticate = response.headers.get('www-authenticate') || '';
      if (response.status === 401 || response.status === 403 || wwwAuthenticate) {
        console.warn('[pipeline-generator] SDK preflight authorization challenge', {
          status: response.status,
          url: response.url || src,
          wwwAuthenticate,
          fedAuthRedirect: response.headers.get('x-tfs-fedauthredirect') || ''
        });
      }
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

  const localSdk = new URL('./lib/VSS.SDK.min.js', window.location.href).toString();
  const localSdkFallback = new URL('./lib/VSS.SDK.js', window.location.href).toString();
  // Only load bundled SDK assets. Some on-prem Azure DevOps hosts challenge
  // platform SDK requests with browser-level Basic auth and trigger repeated
  // login popups despite a valid extension access token.
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
    document.head.appendChild(preload);

    const prefetch = document.createElement('link');
    prefetch.rel = 'prefetch';
    prefetch.as = as || 'script';
    prefetch.href = href;
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

const normalizeRepositoryName = (name) => {
  if (!name) return undefined;
  try {
    return decodeURIComponent(name).trim();
  } catch (error) {
    return name.trim();
  }
};

const getRepositoryNameFromUrl = () => {
  const extractFrom = (url) => {
    if (!url) return undefined;
    try {
      const { pathname } = new URL(url, window.location.origin);
      const match = pathname.match(/\/[_]git\/([^/?]+)/i);
      return normalizeRepositoryName(match?.[1]);
    } catch {
      return undefined;
    }
  };

  return extractFrom(window.location.href) || extractFrom(document.referrer);
};

const getRepositoryName = (context) => {
  const actionContext = getActionContext(context);
  const repository = getRepository(actionContext);
  const candidates = [
    actionContext?.repositoryName,
    actionContext?.repoName,
    repository?.name,
    VSS.getWebContext?.()?.repository?.name,
  ];

  for (const candidate of candidates) {
    const normalized = normalizeRepositoryName(candidate);
    if (normalized) {
      return normalized;
    }
  }

  return getRepositoryNameFromUrl();
};

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
    actionContext?.branch?.fullName,
    actionContext?.branch?.refName,
    actionContext?.branch?.path,
    actionContext?.branch?.name,
    actionContext?.gitRef?.fullName,
    actionContext?.gitRef?.refName,
    actionContext?.gitRef?.name,
    actionContext?.ref?.fullName,
    actionContext?.ref?.refName,
    actionContext?.ref?.name,
    actionContext?.refName,
    actionContext?.selectedItem?.refName,
    actionContext?.selectedItem?.name,
    actionContext?.item?.refName,
    actionContext?.item?.name,
    actionContext?.branchName,
    actionContext?.fullName,
    actionContext?.name
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

const HOST_PAGE_LAYOUT_SERVICE_ID = 'ms.vss-features.host-page-layout-service';
const HOST_NAVIGATION_SERVICE_ID = 'ms.vss-features.host-navigation-service';
const DIALOG_CONTRIBUTION_ID = 'pipeline-generator-dialog';
const HUB_CONTRIBUTION_ID = 'pipeline-generator-hub';

const getHostPageLayoutServiceId = (sdk) =>
  sdk?.ServiceIds?.HostPageLayoutService ||
  sdk?.ServiceIds?.HostPageLayout ||
  HOST_PAGE_LAYOUT_SERVICE_ID;

const buildContributionId = (extensionContext, contributionId) => {
  const publisherId = extensionContext?.publisherId;
  const extensionId = extensionContext?.extensionId;
  if (!publisherId || !extensionId) {
    throw new Error('Azure DevOps did not provide the extension identity required to open the generator.');
  }
  return `${publisherId}.${extensionId}.${contributionId}`;
};

const buildDialogContributionId = (extensionContext) =>
  buildContributionId(extensionContext, DIALOG_CONTRIBUTION_ID);

const buildHubContributionId = (extensionContext) =>
  buildContributionId(extensionContext, HUB_CONTRIBUTION_ID);

const getHostNavigationService = async (sdk) =>
  sdk.getService(sdk?.ServiceIds?.Navigation || HOST_NAVIGATION_SERVICE_ID);

const buildGeneratorHubUrl = ({ hostUri, projectName, extensionContext, params }) => {
  if (!hostUri || !projectName) {
    throw new Error('Azure DevOps project context is required to open the Pipeline Generator hub.');
  }
  const contributionId = buildHubContributionId(extensionContext);
  const query = params.toString();
  return `${hostUri}${encodeURIComponent(projectName)}/_apps/hub/${encodeURIComponent(contributionId)}${
    query ? `?${query}` : ''
  }`;
};

const normalizeGeneratorMode = (mode) => (mode === 'monorepo' ? 'monorepo' : 'pipeline');

const openGenerator = async (context, sdk, requestedMode = 'pipeline') => {
  try {
    const mode = normalizeGeneratorMode(requestedMode);
    const actionContext = getActionContext(context);
    const repository = getRepository(actionContext) || VSS.getWebContext?.()?.repository;
    const branchName = getBranchName(actionContext);
    const project = getProject(actionContext);
    const repoId = repository?.id;
    const repoName = getRepositoryName(actionContext);
    const projectId = project?.id || actionContext?.projectId;
    const projectName = project?.name || projectId;
    const extContext = VSS.getExtensionContext?.();
    const params = new URLSearchParams();

    if (branchName) params.set('branch', branchName);
    if (projectId) params.set('projectId', projectId);
    if (projectName) params.set('projectName', projectName);
    if (repoId) params.set('repoId', repoId);
    if (repoName) params.set('repoName', repoName);
    params.set('mode', mode);

    const hostUri = (VSS.getWebContext?.()?.collection?.uri || getHostBase()).replace(/\/+$/, '') + '/';
    params.set('hostUri', hostUri);
    const bootstrapPayload = {
      branch: branchName,
      projectId,
      projectName,
      repoId,
      repoName,
      hostUri,
      mode
    };

    try {
      const hostService = await sdk.getService(getHostPageLayoutServiceId(sdk));
      if (hostService?.openCustomDialog) {
        hostService.openCustomDialog(buildDialogContributionId(extContext), {
          title: `${mode === 'monorepo' ? 'Generate MonoRepo' : 'Generate pipeline'} for ${branchName || 'branch'}`,
          lightDismiss: false,
          configuration: bootstrapPayload
        });
        return;
      }
    } catch (serviceError) {
      console.warn('Azure DevOps host dialog was unavailable; opening the in-host Repos hub instead', serviceError);
    }

    const navigationService = await getHostNavigationService(sdk);
    if (!navigationService?.navigate) {
      throw new Error('Azure DevOps host navigation service is unavailable; the generator cannot open in-host.');
    }
    navigationService.navigate(
      buildGeneratorHubUrl({ hostUri, projectName, extensionContext: extContext, params })
    );
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

    const buildAction = (mode) => ({
      execute: async (context) => {
        const sdkInstance = await ensureSdkReady();
        if (assetWarmupPromise) {
          await assetWarmupPromise.catch(() => {});
        }
        await openGenerator(context, sdkInstance, mode);
        sdkInstance.notifyLoadSucceeded?.();
      }
    });

    try {
      sdk.register('generate-pipeline-action', buildAction('pipeline'));
      sdk.register('generate-monorepo-action', buildAction('monorepo'));
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
