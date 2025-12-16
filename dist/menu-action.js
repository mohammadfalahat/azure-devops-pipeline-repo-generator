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

const getBranchName = (context) => {
  const branchFromContext = context?.branch?.name || context?.gitRef?.name;
  if (branchFromContext) {
    return branchFromContext;
  }

  const url = new URL(window.location.href);
  const version = url.searchParams.get('version');
  if (version?.startsWith('GB')) {
    return version.slice(2);
  }

  return context?.gitRepository?.defaultBranch || 'Unknown branch';
};

const openGenerator = async (context) => {
  const branchName = getBranchName(context);
  const project = context?.project || context?.gitRepository?.project;
  const repoId = context?.gitRepository?.id;
  const projectId = project?.id || context?.projectId;
  const projectName = project?.name || projectId;
  const extContext = VSS.getExtensionContext();
  const targetUrl = `${extContext.baseUri}index.html?branch=${encodeURIComponent(branchName)}&projectId=${encodeURIComponent(projectId)}&projectName=${encodeURIComponent(projectName)}&repoId=${encodeURIComponent(repoId || '')}`;
  const hostService = await VSS.getService(VSS.ServiceIds.HostPageLayout);
  if (hostService?.openWindow) {
    hostService.openWindow(targetUrl, {});
  } else {
    window.open(targetUrl, '_blank');
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
