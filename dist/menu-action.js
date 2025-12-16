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
    const extContext = VSS.getExtensionContext();
    const params = new URLSearchParams();

    if (branchName) params.set('branch', branchName);
    if (projectId) params.set('projectId', projectId);
    if (projectName) params.set('projectName', projectName);
    if (repoId) params.set('repoId', repoId);

    const targetUrl = `${extContext.baseUri}index.html?${params.toString()}`;
    const hostService = await VSS.getService(VSS.ServiceIds.HostPageLayout);
    if (hostService?.openWindow) {
      hostService.openWindow(targetUrl, {});
    } else {
      window.open(targetUrl, '_blank');
    }
  } catch (error) {
    console.error('Failed to launch pipeline generator', error);
    VSS.handleError?.(error);
  }
};

(async () => {
  await loadVssSdk();
  VSS.init({ usePlatformScripts: true });
  await VSS.ready();

  VSS.register('generate-pipeline-action', {
    execute: openGenerator
  });

  VSS.register('generate-pipeline-repo-action', {
    execute: openGenerator
  });
})();
