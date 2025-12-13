(async () => {
  await SDK.init();
  SDK.register('generate-pipeline-action', {
    execute: async (context) => {
      const branchName = context?.branch?.name || context?.gitRef?.name || 'Unknown branch';
      const project = context?.project || context?.gitRepository?.project;
      const repoId = context?.gitRepository?.id;
      const projectId = project?.id || context?.projectId;
      const projectName = project?.name || projectId;
      const extContext = SDK.getExtensionContext();
      const targetUrl = `${extContext.baseUri}index.html?branch=${encodeURIComponent(branchName)}&projectId=${encodeURIComponent(projectId)}&projectName=${encodeURIComponent(projectName)}&repoId=${encodeURIComponent(repoId || '')}`;
      const hostService = await SDK.getService('ms.vss-features.host-page-layout-service');
      if (hostService?.openWindow) {
        hostService.openWindow(targetUrl, {});
      } else {
        window.open(targetUrl, '_blank');
      }
    }
  });
})();
