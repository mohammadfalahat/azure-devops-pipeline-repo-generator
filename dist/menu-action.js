const getHostBase = () => {
  if (!document.referrer) {
    return window.location.origin;
  }
  const referrer = new URL(document.referrer);
  const segments = referrer.pathname.split('/').filter(Boolean);
  const hasTfsVirtualDir = segments[0]?.toLowerCase() === 'tfs';
  return `${referrer.origin}${hasTfsVirtualDir ? '/tfs' : ''}`;
};

const loadVssSdk = () =>
  new Promise((resolve, reject) => {
    if (window.VSS) {
      resolve(window.VSS);
      return;
    }

    const script = document.createElement('script');
    script.src = `${getHostBase()}/_content/MS.VSS.SDK/scripts/VSS.SDK.min.js`;
    script.async = false;
    script.onload = () => resolve(window.VSS);
    script.onerror = () => reject(new Error('Failed to load Azure DevOps SDK from host.'));
    document.head.appendChild(script);
  });

(async () => {
  await loadVssSdk();
  VSS.init({ usePlatformScripts: true });
  await VSS.ready();

  VSS.register('generate-pipeline-action', {
    execute: async (context) => {
      const branchName = context?.branch?.name || context?.gitRef?.name || 'Unknown branch';
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
    }
  });
})();
