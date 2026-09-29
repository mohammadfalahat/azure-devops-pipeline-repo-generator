(() => {
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

  const normalizeHostUri = (hostUri) => {
    if (!hostUri) return '';

    const trimmed = hostUri.replace(/\/+$/, '');
    const withoutApis = trimmed.replace(/\/_apis\/?$/i, '');

    return `${withoutApis.replace(/\/+$/, '')}/`;
  };

  const buildCollectionUri = (hostUri, collectionName) => {
    const normalizedCollection = String(collectionName || '').trim();
    if (!normalizedCollection || /[\\/\0\r\n]/.test(normalizedCollection)) {
      throw new Error('Central Azure DevOps collection name is invalid.');
    }

    let parsed;
    try {
      parsed = new URL(normalizeHostUri(hostUri));
    } catch {
      throw new Error('Azure DevOps collection URI is invalid.');
    }
    if (!/^https?:$/.test(parsed.protocol)) {
      throw new Error('Azure DevOps collection URI must use HTTP or HTTPS.');
    }

    const currentSegments = parsed.pathname.split('/').filter(Boolean);
    const serverPathSegments = currentSegments.length > 0 ? currentSegments.slice(0, -1) : [];
    const centralPath = [...serverPathSegments, encodeURIComponent(normalizedCollection)].join('/');
    return `${parsed.origin}/${centralPath}/`;
  };

  const buildCentralGitItemUrl = ({ hostUri, source }) => {
    const collectionUri = buildCollectionUri(hostUri, source.collection);
    return `${collectionUri}${encodeURIComponent(source.project)}/_apis/git/repositories/${encodeURIComponent(
      source.repository
    )}/items?path=${encodeURIComponent(source.path)}&versionDescriptor.version=${encodeURIComponent(
      source.branch
    )}&versionDescriptor.versionType=branch&%24format=text&api-version=6.0`;
  };

  // Azure DevOps Server extension access tokens are issued by the collection
  // hosting the current page. A token that is valid for that collection can be
  // rejected by a sibling collection even when the same signed-in user has
  // repository access there. These two central reads stay on the same origin,
  // so use the existing authenticated browser session without forwarding the
  // collection-scoped Bearer token.
  const centralGitRequestOptions = () => ({
    headers: { 'X-TFS-FedAuthRedirect': 'Suppress' },
    cache: 'no-store',
    credentials: 'same-origin',
    redirect: 'manual'
  });

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

    const localSdk = new URL('./lib/VSS.SDK.min.js', window.location.href).toString();
    const localSdkFallback = new URL('./lib/VSS.SDK.js', window.location.href).toString();
    // Only load bundled SDK assets. Some on-prem Azure DevOps hosts challenge
    // requests to the platform SDK endpoint with browser-level Basic auth, which
    // causes repeated username/password popups even when the extension already
    // has a valid access token.
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

  const waitForSdkReady = async (sdk, timeoutMs = 15000) => {
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
    stack: 'default',
    repositoryAddress: 'registry.buluttakin.com',
    containerRegistryService: 'BulutReg',
    dockerfileDir: '**'
  };
  const defaultPoolOptions = ['PublishDockerAgent', 'Default'];
  const defaultRegistryOptions = ['BulutReg', 'DockerReg'];
  const CENTRAL_COLLECTION_NAME = 'ShonizCollection';
  const DEPLOYMENT_TARGETS_CONFIG = Object.freeze({
    collection: CENTRAL_COLLECTION_NAME,
    project: 'SharedTemplates',
    repository: 'SharedTemplates',
    path: '/pipeline-generator.yml',
    branch: 'main'
  });
  const KOMODO_CREDENTIAL_CONFIG = Object.freeze({
    collection: CENTRAL_COLLECTION_NAME,
    project: 'SharedTemplates',
    repository: 'SharedTemplates',
    path: '/komodo-servers-creds.env',
    branch: 'main'
  });

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

  const getDialogConfiguration = (sdk) => {
    const configuration = sdk?.getConfiguration?.();
    if (!configuration || typeof configuration !== 'object') return {};

    // Current hosts return the object supplied to openCustomDialog directly.
    // Accept the nested shapes as well for compatibility with older wrappers.
    return configuration.pipelineBootstrap || configuration.configuration || configuration;
  };

  const getHostNavigationState = async (sdk) => {
    const serviceIds = [
      sdk?.ServiceIds?.Navigation,
      'ms.vss-features.host-navigation-service'
    ].filter((value, index, values) => value && values.indexOf(value) === index);

    for (const serviceId of serviceIds) {
      try {
        const navigationService = await sdk.getService(serviceId);
        if (navigationService?.getCurrentState) {
          return navigationService.getCurrentState() || {};
        }
        if (navigationService?.getQueryParams) {
          return (await navigationService.getQueryParams()) || {};
        }
      } catch (error) {
        console.warn('Could not read Pipeline Generator context from host navigation state', {
          serviceId,
          message: error?.message
        });
      }
    }

    return {};
  };

  const branchLabel = document.getElementById('branch-label');
  const pageTitle = document.getElementById('page-title');
  const formHint = document.getElementById('form-hint');
  const branchInput = document.getElementById('branch');
  const environmentSelect = document.getElementById('environment');
  const stackInput = document.getElementById('stack');
  const poolSelect = document.getElementById('pool');
  const serviceInput = document.getElementById('service');
  const registrySelect = document.getElementById('containerRegistryService');
  const dockerfileInput = document.getElementById('dockerfileDir');
  const form = document.getElementById('pipeline-form');
  const status = document.getElementById('status');
  const targetRepoInput = document.getElementById('targetRepo');
  const komodoSelect = document.getElementById('komodoServer');
  const reauthPanel = document.getElementById('reauth-panel');
  const reauthMessage = document.getElementById('reauth-message');
  const authorizeExtensionButton = document.getElementById('authorize-extension');
  const reauthenticateButton = document.getElementById('reauthenticate');
  const submitButton = form?.querySelector('button[type="submit"]');
  const completionPanel = document.getElementById('completion-panel');
  const nginxResultItem = document.getElementById('nginx-result-item');
  const nginxResultLink = document.getElementById('nginx-result-link');
  const composeResultLink = document.getElementById('compose-result-link');
  const pipelineResultLink = document.getElementById('pipeline-result-link');
  const contractResultItem = document.getElementById('contract-result-item');
  const contractResultLink = document.getElementById('contract-result-link');
  const completionHint = document.getElementById('completion-hint');
  const serviceField = document.getElementById('service-field');
  const dockerfileField = document.getElementById('dockerfile-field');
  const registryAddressField = document.getElementById('registry-address-field');
  const registryServiceField = document.getElementById('registry-service-field');

  const createEditableCombobox = (select, placeholder) => {
    if (!select || typeof window.TomSelect !== 'function') return null;
    const combobox = new window.TomSelect(select, {
      maxItems: 1,
      create: (input) => {
        const value = String(input || '').trim();
        return value ? { value, text: value } : false;
      },
      createOnBlur: true,
      persist: false,
      openOnFocus: true,
      closeAfterSelect: true,
      hideSelected: false,
      selectOnTab: true,
      allowEmptyOption: false,
      placeholder
    });
    select.setAttribute('aria-hidden', 'true');
    const describedBy = select.getAttribute('aria-describedby');
    if (describedBy) combobox.control_input.setAttribute('aria-describedby', describedBy);
    if (select.required) combobox.control_input.setAttribute('aria-required', 'true');
    return combobox;
  };

  const environmentCombobox = createEditableCombobox(
    environmentSelect,
    'Loading from pipeline-generator.yml...'
  );
  const stackCombobox = createEditableCombobox(stackInput, 'Choose or enter a Stack');

  const setEditableDisabled = (select, combobox, disabled) => {
    if (!select) return;
    select.disabled = disabled;
    if (!combobox) return;
    if (disabled) combobox.disable();
    else combobox.enable();
  };

  const setEditablePlaceholder = (select, combobox, placeholder) => {
    if (!select) return;
    select.dataset.placeholder = placeholder;
    if (!combobox) return;
    combobox.settings.placeholder = placeholder;
    combobox.inputState();
  };

  const setEditableValue = (select, combobox, value, silent = true) => {
    if (!select) return;
    const normalizedValue = String(value || '');
    if (!combobox) {
      select.value = normalizedValue;
      return;
    }
    if (normalizedValue && !Object.prototype.hasOwnProperty.call(combobox.options, normalizedValue)) {
      combobox.addOption({ value: normalizedValue, text: normalizedValue });
    }
    combobox.setValue(normalizedValue, silent);
  };

  const populateEditableOptions = (select, combobox, options, placeholder) => {
    if (!select) return;
    if (!combobox) {
      if (placeholder) setEditablePlaceholder(select, null, placeholder);
      populateSelectOptions(select, options, options.length ? undefined : placeholder);
      return;
    }
    combobox.clear(true);
    combobox.clearOptions();
    combobox.addOptions(options.map((option) => ({ value: option.value, text: option.label })));
    if (placeholder) setEditablePlaceholder(select, combobox, placeholder);
    combobox.refreshOptions(false);
  };

  if (targetRepoInput) {
    targetRepoInput.disabled = true;
  }

  if (serviceInput) {
    serviceInput.addEventListener('input', () => {
      const normalized = normalizeServiceNameForForm(serviceInput.value, { trim: false });
      if (serviceInput.value !== normalized) {
        serviceInput.value = normalized;
      }
      serviceInput.dataset.autofilled = 'false';
    });
  }

  const SCAFFOLD_BRANCH = 'main';
  const ZERO_OBJECT_ID = '0000000000000000000000000000000000000000';
  const PIPELINE_FOLDER = '\\komodo';
  const MONOREPO_PIPELINE_FOLDER = '\\komodo\\MR';
  // The Pipelines API remains a preview contract on supported Azure DevOps
  // Server versions. Keep the repositoryId query parameter on create: without
  // it, some on-prem servers accept the YAML commit but reject pipeline
  // registration because they cannot resolve the input repository.
  const PIPELINE_API_VERSION = '7.1-preview.1';
  const BUILD_API_VERSION = '7.1';
  const RELEASE_API_VERSION = '7.1-preview.4';
  const BASH_TASK_ID = '6c731c3c-3c68-459a-a5c9-bde6e6595b5b';
  const DEFAULT_RELEASE_CONFIG = Object.freeze({
    enabled: true,
    folder: '\\komodo',
    environmentName: 'komodo',
    bashTaskName: 'Run Komodo deployment',
    variableGroupName: 'KomodoAPI',
    requiredVariableNames: Object.freeze(['AZP_TOKEN', 'KOMODO_API_KEY', 'KOMODO_API_SECRET']),
    scriptSource: { type: 'inline', content: '' }
  });

  const normalizePipelineFolder = (folder, fallback) => {
    const candidate = String(folder || fallback || '').trim().replace(/\//g, '\\');
    if (!candidate || candidate === '\\') return '\\';
    return `\\${candidate.replace(/^\\+/, '')}`;
  };

  const getReleaseConfig = (mode = state.mode) => {
    const configured = window.PipelineGeneratorReleaseConfig || {};
    const source = configured.scriptSource || DEFAULT_RELEASE_CONFIG.scriptSource;
    const monorepo = normalizeGeneratorMode(mode) === 'monorepo';
    return {
      enabled: configured.enabled !== false,
      folder: monorepo
        ? normalizePipelineFolder(`${configured.folder || DEFAULT_RELEASE_CONFIG.folder}\\MR`, '\\komodo\\MR')
        : normalizePipelineFolder(configured.folder, DEFAULT_RELEASE_CONFIG.folder),
      environmentName: monorepo
        ? 'MR deployment'
        : String(configured.environmentName || DEFAULT_RELEASE_CONFIG.environmentName).trim(),
      bashTaskName: monorepo
        ? 'Deploy immutable MR images through Komodo'
        : String(configured.bashTaskName || DEFAULT_RELEASE_CONFIG.bashTaskName).trim(),
      variableGroupName: String(
        configured.variableGroupName || DEFAULT_RELEASE_CONFIG.variableGroupName
      ).trim(),
      requiredVariableNames: Array.isArray(configured.requiredVariableNames)
        ? configured.requiredVariableNames.map((name) => String(name).trim()).filter(Boolean)
        : [...DEFAULT_RELEASE_CONFIG.requiredVariableNames],
      scriptSource: monorepo ? { type: 'packagedFile', path: 'monorepo-release-inline-task.sh' } : source
    };
  };

  const state = {
    mode: 'pipeline',
    sdk: null,
    accessToken: null,
    accessTokenError: null,
    hostUri: null,
    projectId: null,
    rawProjectName: null,
    projectName: null,
    repoId: null,
    rawRepositoryName: null,
    repositoryName: null,
    generatedRepoId: null,
    generatedRepositoryName: null,
    deploymentTargets: null,
    deploymentTargetsReady: false,
    provisioningComplete: false,
    branch: SCAFFOLD_BRANCH,
    sourceBranch: null
  };
  let initializationPromise;

  const normalizeGeneratorMode = (mode) => (String(mode || '').toLowerCase() === 'monorepo' ? 'monorepo' : 'pipeline');
  const isMonorepoMode = () => state.mode === 'monorepo';
  const getPipelineFolder = () => (isMonorepoMode() ? MONOREPO_PIPELINE_FOLDER : PIPELINE_FOLDER);

  const applyModePresentation = (mode) => {
    state.mode = normalizeGeneratorMode(mode);
    const monorepo = isMonorepoMode();
    if (pageTitle) pageTitle.textContent = monorepo ? 'Generate MonoRepo' : 'Generate pipeline';
    document.title = monorepo ? 'Generate MonoRepo' : 'Generate pipeline';
    if (formHint) {
      formHint.textContent = monorepo
        ? 'Generate one MR Pipeline and one classic Release for this Nx monorepo. SharedTemplates builds immutable static/BFF images, Compose stores their active tags, and Komodo deploys or rolls them back without project Dockerfiles or host mounts.'
        : 'Fill the fields below, then generate the pipeline. The generator will push the template, ensure the project Docker DevOps files and configured-Environment Nginx files, register the YAML pipeline in \\komodo, and create its classic Release definition. It will then show review links without running or redirecting to the Pipeline.';
    }
    if (serviceField) serviceField.hidden = false;
    if (dockerfileField) dockerfileField.hidden = monorepo;
    if (registryAddressField) registryAddressField.hidden = false;
    if (registryServiceField) registryServiceField.hidden = false;
    if (serviceInput) serviceInput.required = true;
    if (dockerfileInput) dockerfileInput.required = !monorepo;
    const repositoryAddressInput = document.getElementById('repositoryAddress');
    if (repositoryAddressInput) repositoryAddressInput.required = true;
    if (registrySelect) registrySelect.required = true;
    if (submitButton) {
      submitButton.textContent = monorepo
        ? 'Create MR runtime, pipeline, and release'
        : 'Create repositories, pipeline, and release';
    }
    contractResultItem?.classList?.toggle('hidden', !monorepo);
    if (contractResultItem && !contractResultItem.classList) {
      contractResultItem.className = monorepo ? '' : 'hidden';
    }
    if (completionHint) {
      completionHint.textContent = monorepo
        ? 'Review the generated files shown below. Then run the MR Pipeline once; later source changes are detected automatically.'
        : 'The files are only starter templates. Review and edit them, then open and run the Pipeline once.';
    }
  };

  const setStatus = (message, isError = false) => {
    status.textContent = message;
    status.className = isError ? 'status-error' : 'status-success';
  };

  const setSubmitting = (isSubmitting) => {
    if (submitButton) {
      submitButton.disabled = isSubmitting || state.provisioningComplete || !state.deploymentTargetsReady;
    }
  };

  const setReauthenticationVisibility = (isVisible, message) => {
    if (!reauthPanel) return;
    if (message && reauthMessage) {
      reauthMessage.textContent = message;
    }
    reauthPanel.classList?.toggle('hidden', !isVisible);
    if (!reauthPanel.classList) {
      reauthPanel.className = isVisible ? 'auth-fallback' : 'auth-fallback hidden';
    }
  };

  const setCompletionVisibility = (isVisible) => {
    completionPanel?.classList?.toggle('hidden', !isVisible);
    if (completionPanel && !completionPanel.classList) {
      completionPanel.className = isVisible ? 'completion-panel' : 'completion-panel hidden';
    }
  };

  const finishProvisioning = () => {
    state.provisioningComplete = true;
    if (form) {
      form.hidden = true;
    }
    setSubmitting(false);
    setCompletionVisibility(true);
    completionPanel?.focus?.({ preventScroll: true });
    completionPanel?.scrollIntoView?.({ behavior: 'smooth', block: 'start' });
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

  const parseDeploymentTargetScalar = (rawValue, lineNumber) => {
    const raw = String(rawValue || '').trim();
    let value;
    if (raw.startsWith('"')) {
      const match = raw.match(/^("(?:\\.|[^"\\])*")\s*(?:#.*)?$/);
      if (!match) {
        throw new Error(`Invalid double-quoted value in pipeline-generator.yml at line ${lineNumber}.`);
      }
      try {
        value = JSON.parse(match[1]);
      } catch (error) {
        throw new Error(`Invalid double-quoted value in pipeline-generator.yml at line ${lineNumber}.`);
      }
    } else if (raw.startsWith("'")) {
      const match = raw.match(/^('(?:''|[^'])*')\s*(?:#.*)?$/);
      if (!match) {
        throw new Error(`Invalid single-quoted value in pipeline-generator.yml at line ${lineNumber}.`);
      }
      value = match[1].slice(1, -1).replace(/''/g, "'");
    } else {
      value = raw.replace(/\s+#.*$/, '').trim();
    }

    if (typeof value !== 'string' || !value.trim()) {
      throw new Error(`Empty deployment target in pipeline-generator.yml at line ${lineNumber}.`);
    }
    value = value.trim();
    if (value.length > 200 || /[\0\r\n]/.test(value) || value.includes("'")) {
      throw new Error(`Unsafe deployment target in pipeline-generator.yml at line ${lineNumber}.`);
    }
    return value;
  };

  const uniqueCaseInsensitive = (values) => {
    const seen = new Set();
    return values.filter((value) => {
      const key = value.toLowerCase();
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });
  };

  const defaultProjectsRootForEnvironment = (environment) =>
    ['pro', 'prod', 'production'].includes(String(environment || '').trim().toLowerCase())
      ? '/mnt/graid/projects'
      : '/var/data/projects';

  const normalizeProjectsRoot = (projectsRoot, environment) => {
    const normalized = String(projectsRoot || defaultProjectsRootForEnvironment(environment))
      .trim()
      .replace(/\/+$/, '');
    if (!['/mnt/graid/projects', '/var/data/projects'].includes(normalized)) {
      throw new Error(
        `Environment ${environment || '(empty)'} projects_root must be /mnt/graid/projects or /var/data/projects.`
      );
    }
    return normalized;
  };

  const parseDeploymentTargetsYaml = (yamlText) => {
    const result = { servers: [], environments: [], environmentConfigs: [] };
    let section;
    let currentEnvironment;
    String(yamlText || '')
      .replace(/^\uFEFF/, '')
      .split(/\r?\n/)
      .forEach((line, index) => {
        const lineNumber = index + 1;
        const trimmed = line.trim();
        if (!trimmed || trimmed.startsWith('#')) return;

        const sectionMatch = line.match(/^([A-Za-z][A-Za-z0-9_-]*):\s*(?:#.*)?$/);
        if (sectionMatch) {
          section = ['servers', 'environments'].includes(sectionMatch[1]) ? sectionMatch[1] : undefined;
          currentEnvironment = undefined;
          return;
        }

        if (section === 'servers') {
          const listMatch = line.match(/^\s+-\s+(.+?)\s*$/);
          if (!listMatch) {
            throw new Error(`Expected a YAML list item in servers at line ${lineNumber}.`);
          }
          result.servers.push(parseDeploymentTargetScalar(listMatch[1], lineNumber));
          return;
        }

        if (section === 'environments') {
          const listMatch = line.match(/^\s+-\s+(.+?)\s*$/);
          if (listMatch) {
            const inlineName = listMatch[1].match(/^name\s*:\s*(.+)$/i);
            if (inlineName) {
              currentEnvironment = {
                name: parseDeploymentTargetScalar(inlineName[1], lineNumber),
                domain: '',
                projectsRoot: ''
              };
            } else {
              const scalar = parseDeploymentTargetScalar(listMatch[1], lineNumber);
              const separator = scalar.indexOf(':');
              currentEnvironment = {
                name: separator > 0 ? scalar.slice(0, separator).trim() : scalar,
                domain: separator > 0 ? scalar.slice(separator + 1).trim() : '',
                projectsRoot: ''
              };
            }
            result.environmentConfigs.push(currentEnvironment);
            return;
          }

          const propertyMatch = line.match(/^\s+(name|domain|projects_root)\s*:\s*(.+?)\s*$/i);
          if (!propertyMatch || !currentEnvironment) {
            throw new Error(`Expected an environment name/domain entry at line ${lineNumber}.`);
          }
          const rawProperty = propertyMatch[1].toLowerCase();
          const property = rawProperty === 'projects_root' ? 'projectsRoot' : rawProperty;
          if (currentEnvironment[property]) {
            throw new Error(`Duplicate environment ${property} at line ${lineNumber}.`);
          }
          currentEnvironment[property] = parseDeploymentTargetScalar(propertyMatch[2], lineNumber);
        }
      });

    result.servers = uniqueCaseInsensitive(result.servers);
    if (!result.environmentConfigs.length) {
      throw new Error('pipeline-generator.yml must contain a non-empty environments list.');
    }
    const seenEnvironments = new Set();
    result.environmentConfigs.forEach((environment) => {
      const name = String(environment.name || '').trim();
      const domain = String(environment.domain || '').trim().toLowerCase();
      if (!/^[A-Za-z0-9][A-Za-z0-9._-]*$/.test(name)) {
        throw new Error(`Environment contains unsupported path characters: ${name || '(empty)'}.`);
      }
      if (
        !/^(?=.{1,253}$)(?:[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?\.)+[a-z]{2,63}$/i.test(domain)
      ) {
        throw new Error(`Environment ${name} must define a valid domain.`);
      }
      const key = name.toLowerCase();
      if (seenEnvironments.has(key)) {
        throw new Error(`Duplicate environment name in pipeline-generator.yml: ${name}.`);
      }
      seenEnvironments.add(key);
      environment.name = name;
      environment.domain = domain;
      environment.projectsRoot = normalizeProjectsRoot(environment.projectsRoot, name);
    });
    result.environments = result.environmentConfigs.map((environment) => environment.name);
    return result;
  };

  const fetchDeploymentTargets = async ({ hostUri }) => {
    const source = DEPLOYMENT_TARGETS_CONFIG;
    const url = buildCentralGitItemUrl({ hostUri, source });
    const res = await fetch(url, centralGitRequestOptions());
    if (!res.ok) {
      const detail = await readErrorDetail(res);
      const error = buildHttpError(`Failed to load deployment targets from ${source.path}`, res, detail);
      error.authenticationMode = 'browser-session';
      throw error;
    }
    return parseDeploymentTargetsYaml(await res.text());
  };

  const parseKomodoCredentialFile = (text) => {
    const values = {};
    const supported = new Set(['KOMODO_ADDRESS', 'KOMODO_API_KEY', 'KOMODO_API_SECRET']);
    String(text || '')
      .replace(/^\uFEFF/, '')
      .split(/\r?\n/)
      .forEach((line, index) => {
        const trimmed = line.trim();
        if (!trimmed || trimmed.startsWith('#')) return;
        const match = trimmed.match(/^([A-Z][A-Z0-9_]*)\s*=\s*(.*)$/);
        if (!match || !supported.has(match[1])) {
          throw new Error(`Unsupported credential-file entry at line ${index + 1}.`);
        }
        if (Object.prototype.hasOwnProperty.call(values, match[1])) {
          throw new Error(`Duplicate ${match[1]} entry in the Komodo credential file.`);
        }
        let value = match[2].trim();
        if (value.startsWith('"')) {
          try {
            value = JSON.parse(value);
          } catch {
            throw new Error(`Invalid quoted credential-file value at line ${index + 1}.`);
          }
        } else if (value.startsWith("'")) {
          if (!/^'(?:''|[^'])*'$/.test(value)) {
            throw new Error(`Invalid quoted credential-file value at line ${index + 1}.`);
          }
          value = value.slice(1, -1).replace(/''/g, "'");
        }
        if (!value || value.length > 2048 || /[\0\r\n]/.test(value)) {
          throw new Error(`Invalid credential-file value at line ${index + 1}.`);
        }
        values[match[1]] = value;
      });

    const missing = [...supported].filter((name) => !values[name]);
    if (missing.length) {
      throw new Error(`Komodo credential file is missing: ${missing.join(', ')}.`);
    }
    if (!/^https:\/\/[^\s]+$/i.test(values.KOMODO_ADDRESS)) {
      throw new Error('KOMODO_ADDRESS in the credential file must be an HTTPS URL.');
    }
    return Object.freeze({
      address: values.KOMODO_ADDRESS.replace(/\/+$/, ''),
      apiKey: values.KOMODO_API_KEY,
      apiSecret: values.KOMODO_API_SECRET
    });
  };

  const fetchKomodoCredentials = async ({ hostUri }) => {
    const source = KOMODO_CREDENTIAL_CONFIG;
    const url = buildCentralGitItemUrl({ hostUri, source });
    const res = await fetch(url, centralGitRequestOptions());
    if (!res.ok) {
      const detail = await readErrorDetail(res);
      const error = buildHttpError(`Failed to load Komodo credentials from ${source.path}`, res, detail);
      error.authenticationMode = 'browser-session';
      error.domain = 'komodo';
      throw error;
    }
    try {
      return parseKomodoCredentialFile(await res.text());
    } catch (error) {
      error.domain = 'komodo';
      throw error;
    }
  };

  const extractEnabledKomodoServers = (payload) => {
    const records = Array.isArray(payload)
      ? payload
      : Array.isArray(payload?.data)
        ? payload.data
        : Array.isArray(payload?.value)
          ? payload.value
          : Array.isArray(payload?.servers)
            ? payload.servers
            : null;
    if (!records) {
      throw new Error('Komodo returned an unsupported ListFullServers response.');
    }
    const servers = uniqueCaseInsensitive(
      records
        .map((record) => record?.data || record)
        .filter((record) => record && record.template !== true && record.config?.enabled === true)
        .map((record) => String(record.name || '').trim())
        .filter(Boolean)
    ).sort((left, right) => left.localeCompare(right, 'en', { sensitivity: 'base' }));
    const unsafeServer = servers.find((server) => server.length > 200 || /[\0\r\n']/.test(server));
    if (unsafeServer || !servers.length) {
      throw new Error(
        unsafeServer ? 'Komodo returned an unsafe server name.' : 'Komodo returned no enabled servers.'
      );
    }
    return servers;
  };

  const fetchKomodoServers = async ({ hostUri }) => {
    try {
      const credentials = await fetchKomodoCredentials({ hostUri });
      const res = await fetch(`${credentials.address}/read`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'X-Api-Key': credentials.apiKey,
          'X-Api-Secret': credentials.apiSecret
        },
        body: JSON.stringify({ type: 'ListFullServers', params: { query: {} } }),
        cache: 'no-store',
        credentials: 'omit',
        referrerPolicy: 'no-referrer'
      });
      if (!res.ok) {
        const error = new Error(`Komodo ListFullServers returned HTTP ${res.status}.`);
        error.status = res.status;
        throw error;
      }
      return extractEnabledKomodoServers(await res.json());
    } catch (error) {
      error.domain = 'komodo';
      throw error;
    }
  };

  const extractDockerNetworkNames = (payload) => {
    const records = Array.isArray(payload)
      ? payload
      : Array.isArray(payload?.data)
        ? payload.data
        : Array.isArray(payload?.value)
          ? payload.value
          : Array.isArray(payload?.networks)
            ? payload.networks
            : null;
    if (!records) {
      throw new Error('Komodo returned an unsupported ListDockerNetworks response.');
    }
    const names = uniqueCaseInsensitive(
      records
        .map((record) => String(record?.name || '').trim())
        .filter(Boolean)
    );
    const unsafeName = names.find((name) => !/^[A-Za-z0-9][A-Za-z0-9_.-]{0,127}$/.test(name));
    if (unsafeName) {
      throw new Error('Komodo returned an unsafe Docker network name.');
    }
    return names;
  };

  const selectNginxNetworkName = (networkNames) => {
    const names = Array.from(networkNames || [], (name) => String(name || '').trim());
    const selected = ['nginx-network', 'nginx-net'].find((candidate) => names.includes(candidate));
    if (!selected) {
      throw new Error('The selected Komodo server has neither nginx-network nor nginx-net.');
    }
    return selected;
  };

  const fetchKomodoDockerNetworks = async ({ hostUri, server }) => {
    if (!server || !/^[A-Za-z0-9][A-Za-z0-9._:-]{0,199}$/.test(server)) {
      throw new Error('A valid Komodo server is required before resolving its Docker network.');
    }
    try {
      const credentials = await fetchKomodoCredentials({ hostUri });
      const res = await fetch(`${credentials.address}/read`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'X-Api-Key': credentials.apiKey,
          'X-Api-Secret': credentials.apiSecret
        },
        body: JSON.stringify({ type: 'ListDockerNetworks', params: { server } }),
        cache: 'no-store',
        credentials: 'omit',
        referrerPolicy: 'no-referrer'
      });
      if (!res.ok) {
        throw new Error(`Komodo ListDockerNetworks returned HTTP ${res.status}.`);
      }
      return extractDockerNetworkNames(await res.json());
    } catch (error) {
      error.domain = 'komodo';
      throw error;
    }
  };

  const resolveNginxNetworkForServer = async ({ hostUri, server }) =>
    selectNginxNetworkName(await fetchKomodoDockerNetworks({ hostUri, server }));

  const loadDeploymentTargets = async ({ hostUri, branch }) => {
    state.deploymentTargetsReady = false;
    setEditableDisabled(environmentSelect, environmentCombobox, true);
    if (komodoSelect) komodoSelect.disabled = true;
    try {
      const [targets, komodoServers] = await Promise.all([
        fetchDeploymentTargets({ hostUri }),
        fetchKomodoServers({ hostUri })
      ]);
      targets.servers = komodoServers;
      state.deploymentTargets = targets;
      populateEditableOptions(
        environmentSelect,
        environmentCombobox,
        targets.environmentConfigs.map(({ name, domain }) => ({
          value: name,
          label: name + ' — ' + domain
        })),
        'Choose or enter an Environment'
      );
      populateSelectOptions(
        komodoSelect,
        targets.servers.map((value) => ({ value, label: value }))
      );
      if (environmentSelect) {
        const defaultEnvironment = targets.environments.find(
          (value) => value.toLowerCase() === defaultValues.environment.toLowerCase()
        );
        setEditableValue(
          environmentSelect,
          environmentCombobox,
          defaultEnvironment || targets.environments[0]
        );
        setEditableDisabled(environmentSelect, environmentCombobox, false);
      }
      if (komodoSelect) {
        komodoSelect.disabled = false;
      }
      applyDetectedEnvironment(branch);
      setKomodoServerFromEnvironment(environmentSelect?.value);
      state.deploymentTargetsReady = true;
      return targets;
    } catch (error) {
      state.deploymentTargets = null;
      state.deploymentTargetsReady = false;
      populateEditableOptions(
        environmentSelect,
        environmentCombobox,
        [],
        'Deployment environments unavailable'
      );
      populateSelectOptions(komodoSelect, [], 'Active Komodo servers unavailable');
      setEditableDisabled(environmentSelect, environmentCombobox, true);
      if (komodoSelect) komodoSelect.disabled = true;
      throw error;
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

  const delay = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

  const normalizeAccessTokenError = (error) => {
    const message = error?.message || 'Unknown Azure DevOps authentication error';
    if (/HostAuthorizationNotFound/i.test(message)) {
      return 'HostAuthorizationNotFound: A Collection Administrator must open Collection Settings → Extensions, select Pipeline Generator, and authorize its requested scopes. If no authorization action is available, reinstall the same published extension version.';
    }
    return message;
  };

  const isHostAuthorizationError = (value) =>
    /HostAuthorizationNotFound|Host authorization was not found/i.test(value?.message || value || '');

  const buildTokenRecoveryMessage = (errorMessage) =>
    isHostAuthorizationError(errorMessage)
      ? `${errorMessage} Use Open extension authorization below; signing out cannot create the missing extension authorization.`
      : `${errorMessage} Sign out and authenticate again below to rebuild the Azure DevOps host session.`;

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
        console.warn('[pipeline-generator] getAccessToken failed', {
          attempt,
          status: error?.status,
          message: error?.message
        });

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

  const getAuthHeader = (token) => {
    const tokenValue = typeof token === 'string' ? token : token?.token;
    if (!tokenValue) {
      throw new Error('Extension access token was unavailable.');
    }

    // Azure DevOps SDK access tokens are OAuth/session tokens and must be sent
    // as Bearer credentials. Treating opaque (non-JWT) tokens as PATs and
    // forcing Basic auth causes repeated browser username/password prompts on
    // on-prem hosts whenever the server challenges unauthorized requests.
    return `Bearer ${tokenValue}`;
  };

  const authHeaders = (token) => ({
    Authorization: getAuthHeader(token),
    'X-TFS-FedAuthRedirect': 'Suppress'
  });

  const sanitizeErrorDetail = (detail = '') =>
    detail
      .replace(/<script[\s\S]*?<\/script>/gi, ' ')
      .replace(/<style[\s\S]*?<\/style>/gi, ' ')
      .replace(/<[^>]+>/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();

  const readErrorDetail = async (response) => {
    try {
      const text = await response.text();
      const sanitized = sanitizeErrorDetail(text || '');
      return sanitized.length > 500 ? `${sanitized.slice(0, 497)}...` : sanitized;
    } catch (error) {
      console.warn('Failed to read error response body', error);
      return '';
    }
  };

  const logAuthDiagnostics = (response, baseMessage) => {
    if (!response) return;

    const wwwAuthenticate = response.headers?.get('www-authenticate') || '';
    const shouldLog = response.status === 401 || response.status === 403 || Boolean(wwwAuthenticate);
    if (!shouldLog) return;

    console.warn('[pipeline-generator] Authorization challenge detected', {
      operation: baseMessage,
      status: response.status,
      url: response.url,
      wwwAuthenticate,
      fedAuthRedirect: response.headers?.get('x-tfs-fedauthredirect') || ''
    });
  };

  const buildHttpError = (baseMessage, response, detail) => {
    logAuthDiagnostics(response, baseMessage);
    const message = `${baseMessage} (${response.status})${detail ? `: ${detail}` : ''}`;
    const error = new Error(message);
    error.status = response.status;
    error.url = response.url;
    if (detail) {
      error.detail = detail;
    }
    return error;
  };

  const markErrorDomain = (error, domain) => {
    if (error && !error.domain) {
      error.domain = domain;
    }
    return error;
  };

  const markRequiredExtensionScope = (error, scope) => {
    if (error && !error.requiredExtensionScope) {
      error.requiredExtensionScope = scope;
    }
    return error;
  };

  const runProvisioningStep = async (label, work) => {
    setStatus(label);
    try {
      return await work();
    } catch (error) {
      if (error && !error.provisioningStep) {
        error.provisioningStep = label;
      }
      throw error;
    }
  };

  const sanitizePipelineNameSegment = (segment, fallback, { lowercase = true } = {}) => {
    const fallbackValue = fallback?.toString() || '';
    const value = segment?.toString().trim();
    const base = value || fallbackValue;
    const normalized = lowercase ? base.toLowerCase() : base;

    const cleaned = normalized
      .replace(/[\\/]+/g, '-')
      .replace(/[^\w.-]+/g, '-')
      .replace(/-+/g, '-')
      .replace(/^-+|-+$/g, '');

    return cleaned || (lowercase ? fallbackValue.toLowerCase() : fallbackValue) || 'segment';
  };

  const buildLegacyPipelineFilename = ({ projectName, repositoryName, branchName }) => {
    const projectSegment = sanitizePipelineNameSegment(projectName, 'project');
    const repoSegment = sanitizePipelineNameSegment(repositoryName || projectName, 'repo');
    const branchSegment = sanitizePipelineNameSegment(branchName?.replace(/^refs\/heads\//, ''), 'branch');
    return `${projectSegment}-${repoSegment}-${branchSegment}.yml`;
  };

  const buildLegacyEnvironmentFirstPipelineFilename = ({ projectName, repositoryName, environment, branchName }) => {
    const projectSegment = sanitizePipelineNameSegment(projectName, 'project');
    const repoSegment = sanitizePipelineNameSegment(repositoryName || projectName, 'repo');
    const environmentSegment = sanitizePipelineNameSegment(environment, 'environment');
    const branchSegment = sanitizePipelineNameSegment(branchName?.replace(/^refs\/heads\//, ''), 'branch');
    return `${projectSegment}-${repoSegment}-${environmentSegment}-${branchSegment}.yml`;
  };

  const buildLegacyServiceLessPipelineFilename = ({
    projectName,
    repositoryName,
    environment,
    branchName,
    mode = 'pipeline'
  }) => {
    const projectSegment = sanitizePipelineNameSegment(projectName, 'project');
    const repoSegment = sanitizePipelineNameSegment(repositoryName || projectName, 'repo');
    if (!String(environment || '').trim()) {
      throw new Error('Environment is required to build the Pipeline filename.');
    }
    const environmentSegment = sanitizePipelineNameSegment(environment, 'environment').toUpperCase();
    const branchSegment = sanitizePipelineNameSegment(
      branchName?.replace(/^refs\/heads\//, ''),
      'branch',
      { lowercase: false }
    ).replace(/(^|[-_.])([a-z])/g, (_, separator, character) => `${separator}${character.toUpperCase()}`);
    const modeSegment = normalizeGeneratorMode(mode) === 'monorepo' ? '-MR' : '';
    return `${projectSegment}-${repoSegment}${modeSegment}-${branchSegment}To${environmentSegment}.yml`;
  };

  const buildPipelineFilename = ({
    projectName,
    repositoryName,
    service,
    environment,
    stack = 'default',
    branchName,
    mode = 'pipeline'
  }) => {
    if (!String(service || '').trim()) {
      throw new Error('Service name is required to build the Pipeline filename.');
    }
    const projectSegment = sanitizePipelineNameSegment(projectName, 'project');
    const repoSegment = sanitizePipelineNameSegment(repositoryName || projectName, 'repo');
    const serviceSegment = sanitizePipelineNameSegment(service, 'service');
    if (!String(environment || '').trim()) {
      throw new Error('Environment is required to build the Pipeline filename.');
    }
    const environmentSegment = sanitizePipelineNameSegment(environment, 'environment').toUpperCase();
    const normalizedStack = normalizeStackName(stack);
    const stackSegment = isDefaultStack(normalizedStack) ? '' : `-${sanitizePipelineNameSegment(normalizedStack, 'stack')}`;
    const branchSegment = sanitizePipelineNameSegment(
      branchName?.replace(/^refs\/heads\//, ''),
      'branch',
      { lowercase: false }
    ).replace(/(^|[-_.])([a-z])/g, (_, separator, character) => `${separator}${character.toUpperCase()}`);
    const modeSegment = normalizeGeneratorMode(mode) === 'monorepo' ? '-MR' : '';
    return `${projectSegment}-${repoSegment}${modeSegment}-${serviceSegment}${stackSegment}-${branchSegment}To${environmentSegment}.yml`;
  };

  const buildPipelineName = (pipelineFilename) => pipelineFilename;

  const buildReleaseName = ({ service, environment, stack = 'default', mode = 'pipeline' }) => {
    const normalizePart = (value, label) => {
      const normalized = String(value || '').trim().replace(/\s+/g, ' ').toUpperCase();
      if (!normalized || normalized.length > 100 || /[\0\r\n]/.test(normalized)) {
        throw new Error(`${label} is required to build the classic Release name.`);
      }
      return normalized;
    };
    const normalizedStack = normalizeStackName(stack);
    const parts = [normalizePart(service, 'Service name')];
    if (!isDefaultStack(normalizedStack)) parts.push(normalizePart(normalizedStack, 'Stack'));
    parts.push(normalizePart(environment, 'Environment'));
    if (normalizeGeneratorMode(mode) === 'monorepo') {
      return `MR ${parts.join(' ')}`;
    }
    return parts.join(' ');
  };

  const getProjectRouteSegment = () => {
    const candidate = state.rawProjectName || state.projectName || state.projectId;
    return candidate ? encodeURIComponent(candidate) : '';
  };

  const buildRepositoryFileUrl = ({ repositoryName, filePath }) => {
    const projectRoute = getProjectRouteSegment();
    return `${state.hostUri}${projectRoute}/_git/${encodeURIComponent(repositoryName)}?path=${encodeURIComponent(
      filePath
    )}&version=GB${encodeURIComponent(SCAFFOLD_BRANCH)}&_a=contents`;
  };

  const showCompletionLinks = ({ supportRepositories, pipelineDefinition, generatedRepo, contractPath }) => {
    const nginx = supportRepositories.find((result) => result.kind === 'nginx');
    const docker = supportRepositories.find((result) => result.kind === 'docker');
    if (!docker?.repo?.name || !pipelineDefinition?.id) {
      throw new Error('Provisioning completed but review links could not be constructed.');
    }
    nginxResultItem?.classList?.toggle('hidden', !nginx);
    if (nginxResultItem && !nginxResultItem.classList) {
      nginxResultItem.className = nginx ? '' : 'hidden';
    }
    if (nginxResultLink && nginx) {
      nginxResultLink.href = buildRepositoryFileUrl({
        repositoryName: nginx.repo.name,
        filePath: nginx.filePath
      });
      nginxResultLink.textContent = `Review ${nginx.filePath}`;
    }
    if (composeResultLink) {
      composeResultLink.href = buildRepositoryFileUrl({
        repositoryName: docker.repo.name,
        filePath: docker.filePath
      });
      composeResultLink.textContent = `Review ${docker.filePath}`;
    }
    if (pipelineResultLink) {
      const projectRoute = getProjectRouteSegment();
      pipelineResultLink.href = `${state.hostUri}${projectRoute}/_build?definitionId=${encodeURIComponent(
        pipelineDefinition.id
      )}`;
      pipelineResultLink.textContent = `Open Pipeline ${pipelineDefinition.name || pipelineDefinition.id}`;
    }
    if (contractResultLink && contractPath && generatedRepo?.name) {
      contractResultLink.href = buildRepositoryFileUrl({
        repositoryName: generatedRepo.name,
        filePath: contractPath
      });
      contractResultLink.textContent = `Review ${contractPath}`;
    }
    finishProvisioning();
  };

  const isUnauthorizedError = (error) =>
    error?.status === 401 ||
    error?.status === 403 ||
    /TF400813/i.test(error?.detail || '') ||
    /\b401\b/.test(error?.message || '');

  const buildSignOutUrl = (hostUri) => `${normalizeHostUri(hostUri || state.hostUri || getHostBase())}_signout`;

  const buildExtensionManagementUrl = (hostUri) =>
    `${normalizeHostUri(hostUri || state.hostUri || getHostBase())}_settings/extensions?tab=installed`;

  const navigateHost = async (url) => {
    const sdk = state.sdk || normalizeSdk(window.VSS || window.parent?.VSS);
    const navigationServiceId = sdk?.ServiceIds?.Navigation;
    if (sdk?.getService && navigationServiceId) {
      try {
        const navigationService = await sdk.getService(navigationServiceId);
        if (navigationService?.navigate) {
          navigationService.navigate(url);
          return;
        }
      } catch (error) {
        console.warn('Azure DevOps host navigation service could not open the sign-out page', error);
      }
    }

    try {
      window.top.location.assign(url);
    } catch (error) {
      console.warn('Top-level navigation was unavailable; opening sign-out in the current window', error);
      window.location.assign(url);
    }
  };

  const restartAzureDevOpsSession = async () => {
    if (!state.hostUri) {
      setStatus('Azure DevOps host context is unavailable. Reopen the generator from the target branch.', true);
      return;
    }
    state.accessToken = null;
    state.accessTokenError = null;
    if (reauthenticateButton) {
      reauthenticateButton.disabled = true;
    }
    setStatus('Signing out of Azure DevOps. Complete the login flow, then reopen Generate pipeline from the branch.');
    await navigateHost(buildSignOutUrl(state.hostUri));
  };

  const openExtensionAuthorization = async () => {
    if (!state.hostUri) {
      setStatus('Azure DevOps host context is unavailable. Reopen the generator from the target branch.', true);
      return;
    }
    state.accessToken = null;
    state.accessTokenError = null;
    if (authorizeExtensionButton) {
      authorizeExtensionButton.disabled = true;
    }
    setStatus('Opening Collection Settings → Extensions. Authorize Pipeline Generator, then reopen it from the branch.');
    await navigateHost(buildExtensionManagementUrl(state.hostUri));
  };

  const applyBootstrapPayload = async (payload = {}, source = 'message') => {
    const {
      branch,
      projectId,
      projectName,
      repoId,
      repoName,
      hostUri,
      accessToken,
      accessTokenError,
      mode
    } = payload;

    applyModePresentation(mode || state.mode);

    const normalizedHost = normalizeHostUri(hostUri || state.hostUri || getHostBase());
    state.sourceBranch = branch || state.sourceBranch;
    state.branch = SCAFFOLD_BRANCH;
    state.projectId = projectId || state.projectId;
    state.rawProjectName = projectName || state.rawProjectName;
    state.projectName = projectName || state.projectName;
    state.repoId = repoId || state.repoId;
    state.rawRepositoryName = repoName || state.rawRepositoryName;
    state.repositoryName = repoName || state.repositoryName;
    state.hostUri = normalizedHost;
    if (accessToken) {
      state.accessToken = accessToken;
      state.accessTokenError = null;
      setReauthenticationVisibility(false);
    } else {
      state.accessTokenError = accessTokenError || state.accessTokenError;
    }

    const targetBranch = state.branch;
    const sourceBranch = state.sourceBranch;
    const branchDescriptor =
      sourceBranch && sourceBranch !== targetBranch
        ? `${targetBranch} (source: ${sourceBranch})`
        : targetBranch;
    branchLabel.textContent = branchDescriptor
      ? `Target branch: ${branchDescriptor}`
      : 'Loading branch context...';
    if (branchInput && targetBranch) {
      branchInput.value = targetBranch;
      branchInput.disabled = true;
    }

    targetRepoInput.value = `${state.projectName || 'project'}_Azure_DevOps`;
    setServiceNameFromRepository(state.repositoryName || state.projectName, state.projectName);

    if (!state.projectId || !state.accessToken || !state.hostUri) {
      let authMessage;
      if (state.accessTokenError) {
        const needsHostAuth = isHostAuthorizationError(state.accessTokenError);
        authMessage = needsHostAuth
          ? 'Azure DevOps could not issue an access token because extension authorization is missing. A Collection Administrator must authorize Pipeline Generator in Collection Settings → Extensions; if no authorization action is available, reinstall this same published version.'
          : `Azure DevOps did not provide an access token (${state.accessTokenError}). Refresh the page or sign in again, then relaunch the generator.`;
      } else {
        authMessage = 'Loaded context from branch action but still waiting for an access token from Azure DevOps. Refresh or try again if this persists.';
      }
      setReauthenticationVisibility(
        true,
        buildTokenRecoveryMessage(authMessage)
      );
      setStatus(authMessage, true);
      setSubmitting(false);
      return;
    }

    try {
      await loadDeploymentTargets({
        hostUri: state.hostUri,
        branch: sourceBranch || targetBranch
      });
      const resourceLoads = [
        loadPools({ hostUri: state.hostUri, projectId: state.projectId, accessToken: state.accessToken }),
        loadProjectStacks({
          hostUri: state.hostUri,
          projectId: state.projectId,
          projectName: state.rawProjectName || state.projectName,
          accessToken: state.accessToken
        })
      ];
      if (!isMonorepoMode()) {
        resourceLoads.push(
          loadContainerRegistries({ hostUri: state.hostUri, projectId: state.projectId, accessToken: state.accessToken }),
          refreshDockerfiles({
            hostUri: state.hostUri,
            projectId: state.projectId,
            repoId: state.repoId,
            branch: sourceBranch || targetBranch,
            accessToken: state.accessToken
          })
        );
      }
      await Promise.all(resourceLoads);
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

  const normalizeServiceNameForForm = (value, { trim = true } = {}) => {
    const serviceName = String(value || '');
    const boundedName = trim ? serviceName.trim() : serviceName;
    return boundedName.toLowerCase().replace(/\s+/g, '_');
  };

  // Build templates receive the form value unchanged, so image repositories
  // may intentionally contain underscores. Keep that spelling aligned with
  // the image pushed by the template; resource/path identifiers continue to
  // use normalizeResourceSegment and may use hyphens instead.
  const normalizeImageServiceSegment = (value, label) => {
    const normalized = normalizeServiceNameForForm(value)
      .replace(/[^a-z0-9._-]+/g, '-')
      .replace(/^[._-]+|[._-]+$/g, '');
    if (!normalized) {
      throw new Error(`${label} cannot be converted to a safe image repository segment.`);
    }
    return normalized;
  };

  const extractRepositoryName = (value) => {
    if (!value) return '';
    const segments = value.split('/').filter(Boolean);
    return segments.length ? segments[segments.length - 1] : value;
  };

  const deriveServiceNameFromRepository = (name, projectName) => {
    const targetName = extractRepositoryName(name) || projectName;
    const compactProject = String(projectName || '').replace(/\s+/g, '');
    const projectPrefixes = [String(projectName || '').trim(), compactProject]
      .filter(Boolean)
      .sort((left, right) => right.length - left.length);
    let serviceName = String(targetName || '').trim();
    for (const prefix of projectPrefixes) {
      if (serviceName.toLowerCase().startsWith(prefix.toLowerCase())) {
        const remainder = serviceName.slice(prefix.length);
        if (/^[\s._-]+/.test(remainder)) {
          serviceName = remainder.replace(/^[\s._-]+/, '');
          break;
        }
      }
    }
    return normalizeServiceNameForForm(serviceName || targetName);
  };

  const setServiceNameFromRepository = (name, projectName) => {
    if (!serviceInput) return;
    const normalizedTarget = deriveServiceNameFromRepository(name, projectName);
    if (!normalizedTarget) return;

    const currentValue = normalizeServiceNameForForm(serviceInput.value);
    const projectDefault = normalizeServiceNameForForm(projectName);
    const wasAutoFilled = serviceInput.dataset.autofilled === 'true';
    const shouldUpdate =
      !currentValue ||
      wasAutoFilled ||
      (projectDefault && currentValue === projectDefault);

    if (shouldUpdate && currentValue !== normalizedTarget) {
      serviceInput.value = normalizedTarget;
      serviceInput.dataset.autofilled = 'true';
    }
  };

  const setKomodoServerFromEnvironment = (environment) => {
    if (!environment || !komodoSelect) return;
    const normalizedEnvironment = environment.toLowerCase().replace(/[^a-z0-9]/g, '');
    const aliases = {
      dev: ['dev', 'development'],
      pro: ['pro', 'prod', 'production'],
      prod: ['prod', 'production'],
      qa: ['qa'],
      demo: ['demo'],
      soc: ['soc']
    };
    const candidates = aliases[normalizedEnvironment] || [normalizedEnvironment];
    const match = Array.from(komodoSelect.options).find((option) => {
      const normalizedServer = option.value.toLowerCase().replace(/[^a-z0-9]/g, '');
      return candidates.some((candidate) => normalizedServer.startsWith(candidate));
    });
    komodoSelect.value = match ? match.value : '';
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
      : [];

    return candidates.find((key) => key && lower.includes(key));
  };

  const applyDetectedEnvironment = (branch) => {
    const detected = detectEnvironmentFromBranch(branch);
    if (detected && environmentSelect) {
      const available = Array.from(environmentSelect.options || []).some(
        (option) => option.value.toLowerCase() === detected.toLowerCase()
      );
      if (available) {
        setEditableValue(environmentSelect, environmentCombobox, detected);
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
    const branchName = SCAFFOLD_BRANCH;
    const refUrl = `${hostUri}${encodeURIComponent(projectId)}/_apis/git/repositories/${repoId}/refs?filter=${encodeURIComponent(
      `heads/${branchName}`
    )}&api-version=6.0`;
    const res = await fetch(refUrl, { headers: authHeaders(accessToken) });

    if (res.status === 404) {
      return ZERO_OBJECT_ID;
    }

    if (!res.ok) {
      const detail = await readErrorDetail(res);
      throw buildHttpError('Failed to query branch', res, detail);
    }

    const payload = await res.json();
    return payload.value?.[0]?.objectId || ZERO_OBJECT_ID;
  };

  const getRepositoryFileContent = async ({ hostUri, projectId, repoId, branchName, path, accessToken }) => {
    const normalizedPath = path.startsWith('/') ? path : `/${path}`;
    const url = `${hostUri}${encodeURIComponent(projectId)}/_apis/git/repositories/${repoId}/items?path=${encodeURIComponent(
      normalizedPath
    )}&versionDescriptor.version=${encodeURIComponent(
      branchName
    )}&versionDescriptor.versionType=branch&%24format=text&api-version=6.0`;
    const res = await fetch(url, { headers: authHeaders(accessToken) });

    if (res.status === 404) {
      return null;
    }

    if (!res.ok) {
      const detail = await readErrorDetail(res);
      throw buildHttpError(`Failed to read repository file ${normalizedPath}`, res, detail);
    }

    return res.text();
  };

  const ensureRepositoryByName = async ({ hostUri, projectId, repositoryName, accessToken }) => {
    const url = `${hostUri}${encodeURIComponent(projectId)}/_apis/git/repositories?api-version=6.0`;

    const res = await fetch(url, {
      headers: authHeaders(accessToken)
    });
    if (!res.ok) {
      const detail = await readErrorDetail(res);
      throw buildHttpError('Failed to list repositories', res, detail);
    }
    const payload = await res.json();
    const existing = (payload.value || []).find((repo) => repo.name === repositoryName);
    if (existing) {
      return existing;
    }

    const createRes = await fetch(url, {
      method: 'POST',
      headers: {
        ...authHeaders(accessToken),
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({ name: repositoryName, project: { id: projectId } })
    });
    if (!createRes.ok) {
      const detail = await readErrorDetail(createRes);
      throw buildHttpError('Failed to create repository', createRes, detail);
    }
    const created = await createRes.json();
    if (!created?.id) {
      throw new Error(`Azure DevOps created ${repositoryName} but returned no repository ID.`);
    }
    return created;
  };

  const listProjectRepositories = async ({ hostUri, projectId, accessToken }) => {
    const url = `${hostUri}${encodeURIComponent(projectId)}/_apis/git/repositories?api-version=6.0`;
    const res = await fetch(url, { headers: authHeaders(accessToken) });
    if (!res.ok) {
      const detail = await readErrorDetail(res);
      throw buildHttpError('Failed to list project repositories', res, detail);
    }
    const payload = await res.json();
    return payload.value || [];
  };

  const extractProjectStacks = ({ items, projectName, environments = [] }) => {
    const compactProject = String(projectName || '').replace(/\s+/g, '').toLowerCase();
    if (!compactProject) return ['default'];
    const suffix = `_${compactProject}`;
    const environmentPrefixes = environments
      .map((value) => {
        try {
          return normalizeResourceSegment(value, 'Environment');
        } catch {
          return '';
        }
      })
      .filter(Boolean)
      .sort((left, right) => right.length - left.length);
    const stacks = new Set(['default']);
    for (const item of items || []) {
      if (!item?.isFolder) continue;
      const folder = String(item.path || item.serverItem || '').replace(/^\/+|\/+$/g, '');
      if (!folder || folder.includes('/') || !folder.endsWith(suffix)) continue;
      const prefix = folder.slice(0, -suffix.length);
      let matchedEnvironment = false;
      for (const environment of environmentPrefixes) {
        if (prefix === environment) {
          matchedEnvironment = true;
          stacks.add('default');
          break;
        }
        if (!prefix.startsWith(`${environment}_`)) continue;
        matchedEnvironment = true;
        const candidate = prefix.slice(environment.length + 1);
        try {
          if (candidate && normalizeStackName(candidate) === candidate) stacks.add(candidate);
        } catch {
          // Ignore legacy folders that cannot be represented safely as a Stack.
        }
        break;
      }
      if (!matchedEnvironment) {
        const separatorIndex = prefix.indexOf('_');
        const candidate = separatorIndex > 0 ? prefix.slice(separatorIndex + 1) : '';
        try {
          if (candidate && normalizeStackName(candidate) === candidate) stacks.add(candidate);
        } catch {
          // Unknown custom Environment folders still follow env_stack_project.
        }
      }
    }
    return ['default', ...Array.from(stacks).filter((value) => value !== 'default').sort()];
  };

  const fetchProjectStacks = async ({
    hostUri,
    projectId,
    projectName,
    accessToken,
    environments = []
  }) => {
    const compactProject = String(projectName || '').replace(/\s+/g, '');
    if (!compactProject) return ['default'];
    const repositories = await listProjectRepositories({ hostUri, projectId, accessToken });
    const dockerRepository = repositories.find((repo) => repo.name === `${compactProject}_Docker_DevOps`);
    if (!dockerRepository?.id) return ['default'];
    const url = `${hostUri}${encodeURIComponent(projectId)}/_apis/git/repositories/${encodeURIComponent(
      dockerRepository.id
    )}/items?scopePath=%2F&recursionLevel=OneLevel&versionDescriptor.version=${encodeURIComponent(
      SCAFFOLD_BRANCH
    )}&versionDescriptor.versionType=branch&api-version=6.0`;
    const res = await fetch(url, { headers: authHeaders(accessToken) });
    if (res.status === 404) return ['default'];
    if (!res.ok) {
      const detail = await readErrorDetail(res);
      throw buildHttpError('Failed to discover existing project Stacks', res, detail);
    }
    const payload = await res.json();
    return extractProjectStacks({ items: payload.value || [], projectName: compactProject, environments });
  };

  const loadProjectStacks = async ({ hostUri, projectId, projectName, accessToken }) => {
    const currentValue = String(stackInput?.value || defaultValues.stack).trim() || defaultValues.stack;
    let stacks = ['default'];
    try {
      stacks = await fetchProjectStacks({
        hostUri,
        projectId,
        projectName,
        accessToken,
        environments: state.deploymentTargets?.environments || []
      });
    } catch (error) {
      console.warn('Could not discover existing project Stacks; keeping the default option.', error);
    }
    populateEditableOptions(
      stackInput,
      stackCombobox,
      stacks.map((value) => ({ value, label: value })),
      'Choose or enter a Stack'
    );
    setEditableValue(stackInput, stackCombobox, currentValue);
    return stacks;
  };

  const ensureRepo = async ({ hostUri, projectId, projectName, accessToken }) => {
    const targetName = `${projectName}_Azure_DevOps`;
    targetRepoInput.value = targetName;
    return ensureRepositoryByName({
      hostUri,
      projectId,
      repositoryName: targetName,
      accessToken
    });
  };

  const ensureDefaultBranch = async ({ hostUri, projectId, repoId, branchName, accessToken }) => {
    const defaultBranch = `refs/heads/${branchName}`;
    const url = `${hostUri}${encodeURIComponent(projectId)}/_apis/git/repositories/${repoId}?api-version=6.0`;
    const res = await fetch(url, {
      method: 'PATCH',
      headers: {
        ...authHeaders(accessToken),
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({ defaultBranch })
    });

    if (!res.ok) {
      const detail = await readErrorDetail(res);
      throw buildHttpError('Failed to set default branch', res, detail);
    }

    return res.json();
  };

  const normalizeResourceSegment = (value, label) => {
    const normalized = String(value || '')
      .trim()
      .toLowerCase()
      .replace(/\s+/g, '-')
      .replace(/[^a-z0-9._-]+/g, '-')
      .replace(/[._-]+/g, '-')
      .replace(/^-+|-+$/g, '');
    if (!normalized) {
      throw new Error(`${label} cannot be converted to a safe resource name.`);
    }
    return normalized;
  };

  const normalizeStackName = (value = 'default') =>
    normalizeResourceSegment(String(value || '').trim() || 'default', 'Stack');

  const isDefaultStack = (value) => normalizeStackName(value) === 'default';

  const buildNginxDirectory = ({ environment, stack = 'default' }) => {
    const normalizedEnvironment = normalizeResourceSegment(environment, 'Environment');
    const normalizedStack = normalizeStackName(stack);
    return [normalizedEnvironment, ...(isDefaultStack(normalizedStack) ? [] : [normalizedStack])].join('_');
  };

  const buildComposeDirectory = ({ environment, stack = 'default', projectName }) => {
    const normalizedEnvironment = normalizeResourceSegment(environment, 'Environment');
    const normalizedStack = normalizeStackName(stack);
    const compactProject = String(projectName || '').replace(/\s+/g, '').toLowerCase();
    if (!compactProject || /[\\/\0\r\n]/.test(compactProject)) {
      throw new Error('Project name cannot be converted to a safe Compose directory name.');
    }
    return [normalizedEnvironment, ...(isDefaultStack(normalizedStack) ? [] : [normalizedStack]), compactProject].join('_');
  };

  const frontendRoleAliases = Object.freeze([
    'ui',
    'front',
    'frontend',
    'fe',
    'website',
    'web',
    'client',
    'portal',
    'spa'
  ]);
  const backendRoleAliases = Object.freeze([
    'api',
    'back',
    'backend',
    'be',
    'server',
    'bff',
    'rest',
    'graphql',
    'gateway'
  ]);

  const findServiceRoleSignal = (tokens, aliases) => {
    const aliasSet = new Set(aliases);
    const tokenMatches = tokens
      .map((token, index) => ({ alias: token, index }))
      .filter(({ alias }) => aliasSet.has(alias));
    if (tokenMatches.length) {
      return {
        ...tokenMatches[0],
        matchType: tokens.length === 1 ? 'exact' : 'token'
      };
    }

    const compactName = tokens.join('');
    const affixMatches = aliases
      .filter(
        (alias) =>
          alias.length >= 3 &&
          compactName !== alias &&
          (compactName.startsWith(alias) || compactName.endsWith(alias))
      )
      .sort((left, right) => right.length - left.length || left.localeCompare(right));
    if (!affixMatches.length) return null;
    return {
      alias: affixMatches[0],
      index: compactName.startsWith(affixMatches[0]) ? 0 : 1,
      matchType: 'affix'
    };
  };

  const classifyServiceRouting = (service) => {
    const serviceKey = normalizeResourceSegment(service, 'Service name');
    const rolePattern = '(?:frontend|front|ui|fe|website|web|client|portal|spa|backend|back|api|be|server|bff|rest|graphql|gateway)';
    const variantPattern = '(?:v(?:ersion)?[0-9]+|ver[0-9]+|r[0-9]+|[0-9]+|new|refactor|rewrite|revamp|nextgen|next|modern|latest)';
    const semanticName = String(service || '')
      .trim()
      .replace(/([a-z0-9])([A-Z])/g, '$1-$2')
      .toLowerCase()
      .replace(/front[\s._-]*end/g, 'frontend')
      .replace(/back[\s._-]*end/g, 'backend')
      .replace(/user[\s._-]*interface/g, 'ui')
      .replace(/next[\s._-]*gen(?:eration)?/g, 'nextgen')
      .replace(/refactor(?:ed|ing)?/g, 'refactor')
      .replace(/re[\s._-]*write/g, 'rewrite')
      .replace(/(?:web|rest)[\s._-]*api/g, 'api')
      .replace(new RegExp(`(${rolePattern})(${variantPattern})`, 'g'), '$1-$2')
      .replace(new RegExp(`(${variantPattern})(${rolePattern})`, 'g'), '$1-$2')
      .replace(/[^a-z0-9]+/g, '-')
      .replace(/^-+|-+$/g, '');
    const tokens = semanticName ? semanticName.split('-').filter(Boolean) : [];
    const frontendSignal = findServiceRoleSignal(tokens, frontendRoleAliases);
    const backendSignal = findServiceRoleSignal(tokens, backendRoleAliases);
    let kind = 'service';
    if (frontendSignal || backendSignal) {
      if (!backendSignal) {
        kind = 'frontend';
      } else if (!frontendSignal) {
        kind = 'backend';
      } else {
        const matchRank = { exact: 0, token: 1, affix: 2 };
        const frontendRank = [matchRank[frontendSignal.matchType], frontendSignal.index];
        const backendRank = [matchRank[backendSignal.matchType], backendSignal.index];
        kind =
          frontendRank[0] < backendRank[0] ||
          (frontendRank[0] === backendRank[0] && frontendRank[1] <= backendRank[1])
            ? 'frontend'
            : 'backend';
      }
    }

    const numberedVariants = [];
    let namedVariant = '';
    const namedVariants = new Map([
      ['new', 'new'],
      ['refactor', 'refactor'],
      ['rewrite', 'rewrite'],
      ['revamp', 'revamp'],
      ['next', 'next'],
      ['nextgen', 'nextgen'],
      ['modern', 'modern'],
      ['latest', 'latest']
    ]);
    tokens.forEach((token, index) => {
      let versionMatch = /^(?:v(?:ersion)?|ver|r)([0-9]+)$/.exec(token);
      if (!versionMatch && /^(?:v|version|ver)$/.test(token) && /^[0-9]+$/.test(tokens[index + 1] || '')) {
        versionMatch = ['', tokens[index + 1]];
      }
      if (!versionMatch && /^[0-9]+$/.test(token)) {
        versionMatch = ['', token];
      }
      if (versionMatch) {
        const version = Number(versionMatch[1]);
        if (Number.isSafeInteger(version) && version >= 2) numberedVariants.push(version);
      } else if (!namedVariant && namedVariants.has(token)) {
        namedVariant = namedVariants.get(token);
      }
    });
    const variant = numberedVariants.length ? `v${Math.max(...numberedVariants)}` : namedVariant;
    let location = `/${serviceKey}/`;
    if (kind === 'frontend') location = variant ? `/${variant}/` : '/';
    if (kind === 'backend') location = variant ? `/api/${variant}/` : '/api/';
    return {
      kind,
      variant,
      location,
      internalPort: kind === 'frontend' ? 80 : 8080
    };
  };

  const buildProjectServiceRoutingPlan = ({ repositories, projectName, repositoryName, service }) => {
    const normalizeIdentity = (value) => String(value || '').trim().toLowerCase();
    const normalizedProjectName = normalizeIdentity(projectName);
    const buildCandidate = (repoName, candidateService) => {
      const repositoryMatchesProject =
        Boolean(normalizedProjectName) && normalizeIdentity(repoName) === normalizedProjectName;
      const serviceKey = normalizeResourceSegment(candidateService || repoName, 'Service name');
      const classified = classifyServiceRouting(serviceKey);
      const routing = repositoryMatchesProject
        ? { kind: 'frontend', variant: '', location: '/', internalPort: 80 }
        : classified;
      return {
        repositoryName: String(repoName || ''),
        repositoryIdentity: normalizeIdentity(repoName),
        repositoryMatchesProject,
        serviceKey,
        routing
      };
    };

    const candidatesByRepository = new Map();
    for (const repository of repositories || []) {
      const repoName = String(repository?.name || '').trim();
      if (!repoName) continue;
      candidatesByRepository.set(
        normalizeIdentity(repoName),
        buildCandidate(repoName, deriveServiceNameFromRepository(repoName, projectName))
      );
    }
    const currentRepositoryName = String(repositoryName || projectName || '').trim();
    const currentCandidate = buildCandidate(
      currentRepositoryName,
      service || deriveServiceNameFromRepository(currentRepositoryName, projectName)
    );
    candidatesByRepository.set(currentCandidate.repositoryIdentity, currentCandidate);
    const candidates = Array.from(candidatesByRepository.values());

    const baseRouteAliases = new Map([
      ['/', frontendRoleAliases],
      ['/api/', backendRoleAliases]
    ]);
    const compareCandidates = (baseLocation, left, right) => {
      const aliases = baseRouteAliases.get(baseLocation);
      const aliasSet = new Set(aliases);
      const score = (candidate) => {
        const tokens = candidate.serviceKey.split('-').filter(Boolean);
        const roleSignal = findServiceRoleSignal(tokens, aliases);
        const exactRoleName = tokens.length === 1 && aliasSet.has(tokens[0]);
        const roleTokenCount = tokens.filter((token) => aliasSet.has(token)).length;
        const qualifierCount = roleTokenCount
          ? Math.max(0, tokens.length - roleTokenCount)
          : roleSignal
            ? 1
            : tokens.length;
        const qualifierLength = roleSignal
          ? Math.max(0, candidate.serviceKey.replace(/-/g, '').length - roleSignal.alias.length)
          : candidate.serviceKey.length;
        const matchTypeRank = roleSignal
          ? { exact: 0, token: 1, affix: 2 }[roleSignal.matchType]
          : 3;
        return [
          candidate.repositoryMatchesProject ? 0 : 1,
          exactRoleName ? 0 : 1,
          qualifierCount,
          matchTypeRank,
          qualifierLength,
          tokens.length
        ];
      };
      const leftScore = score(left);
      const rightScore = score(right);
      for (let index = 0; index < leftScore.length; index += 1) {
        if (leftScore[index] !== rightScore[index]) return leftScore[index] - rightScore[index];
      }
      return left.repositoryName.localeCompare(right.repositoryName, 'en', { sensitivity: 'base' });
    };
    const owners = new Map();
    for (const baseLocation of baseRouteAliases.keys()) {
      const eligible = candidates.filter((candidate) => candidate.routing.location === baseLocation);
      if (eligible.length) {
        eligible.sort((left, right) => compareCandidates(baseLocation, left, right));
        owners.set(baseLocation, eligible[0]);
      }
    }

    const routingOverrides = new Map();
    for (const candidate of candidates) {
      const owner = owners.get(candidate.routing.location);
      const ownsBaseRoute = !owner || owner.repositoryIdentity === candidate.repositoryIdentity;
      const routing = ownsBaseRoute
        ? candidate.routing
        : { ...candidate.routing, location: `/${candidate.serviceKey}/` };
      routingOverrides.set(candidate.serviceKey, routing);
      if (candidate.repositoryIdentity === currentCandidate.repositoryIdentity) {
        currentCandidate.routing = routing;
        currentCandidate.ownsBaseRoute = ownsBaseRoute;
        currentCandidate.baseRouteOwner = owner?.repositoryName || '';
      }
    }
    return {
      current: currentCandidate,
      routingOverrides,
      rootOwner: owners.get('/')?.repositoryName || '',
      apiOwner: owners.get('/api/')?.repositoryName || ''
    };
  };

  const buildComposeSample = ({
    projectKey,
    serviceKey,
    imageServiceKey = serviceKey,
    environment,
    stack = 'default',
    repositoryAddress,
    nginxNetworkName = 'nginx-network',
    routing: requestedRouting
  }) => {
    const normalizedStack = normalizeStackName(stack);
    const containerStackSegment = isDefaultStack(normalizedStack) ? '' : `_${normalizedStack.replace(/-/g, '_')}`;
    const imageStackSegment = isDefaultStack(normalizedStack) ? '' : `-${normalizedStack}`;
    const containerName = `${projectKey}_${serviceKey.replace(/-/g, '_')}${containerStackSegment}_${environment}`;
    const internalPort = (requestedRouting || classifyServiceRouting(serviceKey)).internalPort;
    const registry = String(repositoryAddress || defaultValues.repositoryAddress).trim().replace(/\/+$/, '');
    const networkName = selectNginxNetworkName([nginxNetworkName]);
    return [
      'services:',
      `  ${containerName}:`,
      `    container_name: ${containerName}`,
      `    image: ${registry}/${projectKey}/${imageServiceKey}${imageStackSegment}-${environment}:\${IMAGE_TAG:-CHANGE_ME}`,
      '    restart: unless-stopped',
      '    expose:',
      `      - "${internalPort}"`,
      '    networks:',
      `      - ${networkName}`,
      '',
      'networks:',
      `  ${networkName}:`,
      `    name: ${networkName}`,
      '    external: true',
      ''
    ].join('\n');
  };

  const NGINX_MANAGED_ROUTES_START = '    # BEGIN PIPELINE-GENERATOR MANAGED ROUTES';
  const NGINX_MANAGED_ROUTES_END = '    # END PIPELINE-GENERATOR MANAGED ROUTES';

  const buildNginxRouteBlock = ({
    projectKey,
    serviceKey,
    environment,
    stack = 'default',
    containerName: requestedContainerName,
    location: requestedLocation,
    internalPort: requestedInternalPort,
    frontend: requestedFrontend,
    routing: requestedRouting
  }) => {
    const normalizedStack = normalizeStackName(stack);
    const stackSegment = isDefaultStack(normalizedStack) ? '' : `_${normalizedStack.replace(/-/g, '_')}`;
    const containerName = requestedContainerName ||
      `${projectKey}_${serviceKey.replace(/-/g, '_')}${stackSegment}_${environment}`;
    const routing = requestedRouting || classifyServiceRouting(serviceKey);
    const frontend = typeof requestedFrontend === 'boolean' ? requestedFrontend : routing.kind === 'frontend';
    const location = requestedLocation || routing.location;
    const internalPort = requestedInternalPort || (frontend ? 80 : routing.internalPort);
    const upstreamDirectives = [
      '        resolver         127.0.0.11         ipv6=off;',
      `        set              $target            ${containerName};`
    ];
    upstreamDirectives.push(
      `        proxy_pass                          http://$target:${internalPort};`
    );
    return {
      location,
      content: [
        `    # BEGIN PIPELINE-GENERATOR ROUTE ${serviceKey}`,
        `    location ${location} {`,
        ...upstreamDirectives,
        '        proxy_http_version 1.1;',
        '        proxy_set_header Upgrade $http_upgrade;',
        '        proxy_set_header Connection "upgrade";',
        '        proxy_set_header Host $host;',
        '        proxy_set_header X-Real-IP $remote_addr;',
        '        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;',
        '        proxy_set_header X-Forwarded-Proto $scheme;',
        '        proxy_read_timeout 3600s;',
        '        proxy_send_timeout 3600s;',
        '    }',
        `    # END PIPELINE-GENERATOR ROUTE ${serviceKey}`
      ].join('\n')
    };
  };

  const tokenizeNginx = (content) => {
    const tokens = [];
    let index = 0;
    while (index < content.length) {
      const character = content[index];
      if (/\s/.test(character)) {
        index += 1;
        continue;
      }
      if (character === '#') {
        const newline = content.indexOf('\n', index);
        index = newline === -1 ? content.length : newline + 1;
        continue;
      }
      if (character === '"' || character === "'") {
        const quote = character;
        const start = index;
        let value = '';
        index += 1;
        let closed = false;
        while (index < content.length) {
          if (content[index] === '\\' && index + 1 < content.length) {
            value += content[index + 1];
            index += 2;
            continue;
          }
          if (content[index] === quote) {
            index += 1;
            closed = true;
            break;
          }
          value += content[index];
          index += 1;
        }
        if (!closed) {
          throw new Error('Nginx configuration contains an unterminated quoted value.');
        }
        tokens.push({ value, start, end: index });
        continue;
      }
      if ('{};'.includes(character)) {
        tokens.push({ value: character, start: index, end: index + 1 });
        index += 1;
        continue;
      }
      const start = index;
      while (index < content.length && !/[\s{};#]/.test(content[index])) {
        index += 1;
      }
      tokens.push({ value: content.slice(start, index), start, end: index });
    }
    let depth = 0;
    tokens.forEach((token) => {
      if (token.value === '}') depth -= 1;
      if (depth < 0) {
        throw new Error('Nginx configuration has an unmatched closing brace.');
      }
      token.depth = depth;
      if (token.value === '{') depth += 1;
    });
    if (depth !== 0) {
      throw new Error('Nginx configuration has unmatched braces.');
    }
    return tokens;
  };

  const findMatchingBraceToken = (tokens, openIndex) => {
    let nested = 0;
    for (let index = openIndex; index < tokens.length; index += 1) {
      if (tokens[index].value === '{') nested += 1;
      if (tokens[index].value === '}') nested -= 1;
      if (nested === 0) return index;
    }
    return -1;
  };

  const readNginxDirectiveValues = (tokens, directiveIndex, blockDepth) => {
    const values = [];
    for (let index = directiveIndex + 1; index < tokens.length; index += 1) {
      const token = tokens[index];
      if (token.depth !== blockDepth) continue;
      if (token.value === ';') return values;
      if (token.value === '{' || token.value === '}') return [];
      values.push(token.value);
    }
    return [];
  };

  const findNginxHttpsServer = (content, serverName) => {
    const tokens = tokenizeNginx(content);
    const matches = [];
    for (let index = 0; index < tokens.length - 1; index += 1) {
      const token = tokens[index];
      if (token.value.toLowerCase() !== 'server' || tokens[index + 1]?.value !== '{') continue;
      const openIndex = index + 1;
      const closeIndex = findMatchingBraceToken(tokens, openIndex);
      if (closeIndex === -1) {
        throw new Error('Nginx server block has no matching closing brace.');
      }
      const blockDepth = tokens[openIndex].depth + 1;
      const names = [];
      const listens = [];
      const locations = [];
      for (let cursor = openIndex + 1; cursor < closeIndex; cursor += 1) {
        const current = tokens[cursor];
        if (current.depth !== blockDepth) continue;
        const keyword = current.value.toLowerCase();
        if (keyword === 'server_name') {
          names.push(...readNginxDirectiveValues(tokens, cursor, blockDepth));
        } else if (keyword === 'listen') {
          listens.push(...readNginxDirectiveValues(tokens, cursor, blockDepth));
        } else if (keyword === 'location') {
          const values = [];
          for (let valueIndex = cursor + 1; valueIndex < closeIndex; valueIndex += 1) {
            const valueToken = tokens[valueIndex];
            if (valueToken.depth !== blockDepth) continue;
            if (valueToken.value === '{') break;
            if (valueToken.value === ';' || valueToken.value === '}') break;
            values.push(valueToken.value);
          }
          if (values.length && !values[0].startsWith('~')) {
            const location = ['=', '^~'].includes(values[0]) ? values[1] : values[0];
            if (location) locations.push(location);
          }
        }
      }
      const hasServerName = names.some((name) => name.toLowerCase() === serverName.toLowerCase());
      const listensOnHttps = listens.some((listen) => /(?:^|:)443$/.test(listen));
      if (hasServerName && listensOnHttps) {
        matches.push({
          open: tokens[openIndex],
          close: tokens[closeIndex],
          locations
        });
      }
      index = closeIndex;
    }
    if (!matches.length) {
      throw new Error(`Nginx file has no HTTPS server block for ${serverName}.`);
    }
    if (matches.length > 1) {
      throw new Error(`Nginx file has multiple HTTPS server blocks for ${serverName}; merge them manually first.`);
    }
    return matches[0];
  };

  const migrateNginxCertificatePaths = ({ content, serverName, domain }) => {
    const normalizedDomain = String(domain || '').trim().toLowerCase();
    if (!normalizedDomain) return content;
    const legacyCertificateName = normalizedDomain.split('.')[0];
    if (!legacyCertificateName || legacyCertificateName === normalizedDomain) return content;

    const server = findNginxHttpsServer(content, serverName);
    const tokens = tokenizeNginx(content);
    const blockDepth = server.open.depth + 1;
    const replacements = [];
    const legacyPaths = {
      ssl_certificate: `/etc/nginx/conf.d/${legacyCertificateName}.pem`,
      ssl_certificate_key: `/etc/nginx/conf.d/${legacyCertificateName}.key`
    };
    const currentPaths = {
      ssl_certificate: `/etc/nginx/conf.d/${normalizedDomain}.pem`,
      ssl_certificate_key: `/etc/nginx/conf.d/${normalizedDomain}.key`
    };

    for (let index = 0; index < tokens.length; index += 1) {
      const directive = tokens[index];
      if (
        directive.start <= server.open.end ||
        directive.start >= server.close.start ||
        directive.depth !== blockDepth
      ) {
        continue;
      }
      const keyword = directive.value.toLowerCase();
      if (!Object.prototype.hasOwnProperty.call(legacyPaths, keyword)) continue;

      const values = [];
      for (let cursor = index + 1; cursor < tokens.length; cursor += 1) {
        const token = tokens[cursor];
        if (token.start >= server.close.start || token.depth !== blockDepth) continue;
        if (token.value === ';') break;
        if (token.value === '{' || token.value === '}') {
          values.length = 0;
          break;
        }
        values.push(token);
      }
      if (values.length !== 1 || values[0].value !== legacyPaths[keyword]) continue;

      const valueToken = values[0];
      const original = content.slice(valueToken.start, valueToken.end);
      const quote = original[0] === '"' || original[0] === "'" ? original[0] : '';
      replacements.push({
        start: valueToken.start,
        end: valueToken.end,
        value: quote ? `${quote}${currentPaths[keyword]}${quote}` : currentPaths[keyword]
      });
    }

    return replacements
      .sort((left, right) => right.start - left.start)
      .reduce(
        (migrated, replacement) =>
          `${migrated.slice(0, replacement.start)}${replacement.value}${migrated.slice(replacement.end)}`,
        content
      );
  };

  const normalizeNginxManagedRoutes = ({ content, startIndex, endIndex, routingOverrides }) => {
    const managedRoutes = content.slice(startIndex, endIndex);
    const routeBlockPattern = () =>
      /^([ \t]*# BEGIN PIPELINE-GENERATOR ROUTE ([^\r\n]+)[ \t]*\r?\n)([\s\S]*?)(^[ \t]*# END PIPELINE-GENERATOR ROUTE \2[ \t]*)(?:\r?\n)?/gm;
    const escapeRegex = (value) => String(value).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    let migratedRoutes = managedRoutes.replace(
      routeBlockPattern(),
      (routeBlock, startMarker, serviceKey, routeBody, endMarker) => {
        const routingOverride = routingOverrides?.get(serviceKey);
        const routing = routingOverride || classifyServiceRouting(serviceKey);
        const canonicalLocation = routing.location;
        const frontend = routing.kind === 'frontend';
        const legacyFrontend =
          ['ui', 'front', 'frontend', 'newui'].includes(serviceKey) || serviceKey.endsWith('-ui');
        const legacyGeneratedLocations = new Set([
          `/${serviceKey}`,
          `/${serviceKey}/`,
          ...(legacyFrontend ? ['/'] : []),
          ...(routingOverride ? ['/', '/api/'] : [])
        ]);
        let inferredRoute = false;
        let migratedBody = routeBody.replace(
          /^([ \t]*location[ \t]+)([^\s{]+)([ \t]*\{[ \t]*$)/m,
          (line, prefix, existingLocation, suffix) => {
            inferredRoute = existingLocation === canonicalLocation || legacyGeneratedLocations.has(existingLocation);
            return inferredRoute ? `${prefix}${canonicalLocation}${suffix}` : line;
          }
        );

        migratedBody = migratedBody.replace(
          /^([ \t]*)proxy_pass[ \t]+http:\/\/([A-Za-z0-9][A-Za-z0-9._-]*):([0-9]+)\/?;[ \t]*$/gm,
          (_, indentation, containerName, port) => {
            return [
              `${indentation}resolver         127.0.0.11         ipv6=off;`,
              `${indentation}set              $target            ${containerName};`,
              `${indentation}proxy_pass                          http://$target:${port};`
            ].join('\n');
          }
        );

        migratedBody = migratedBody.replace(
          /^([ \t]*)proxy_pass[ \t]+http:\/\/\$target:([0-9]+)\/?;[ \t]*$/gm,
          (_, indentation, port) => {
            return `${indentation}proxy_pass                          http://$target:${port};`;
          }
        );

        if (inferredRoute) {
          migratedBody = migratedBody.replace(
            /^([ \t]*proxy_pass[ \t]+http:\/\/\$target:)[0-9]+(;[ \t]*$)/gm,
            `$1${routing.internalPort}$2`
          );
        }

        if (!frontend || canonicalLocation !== '/') {
          const rewriteLocations = Array.from(new Set([`/${serviceKey}/`, canonicalLocation]))
            .filter((location) => location !== '/');
          for (const location of rewriteLocations) {
            const rewriteExpression = `^${location}(.*)$`;
            const generatedRewritePattern = new RegExp(
              `^[ \\t]*rewrite[ \\t]+${escapeRegex(rewriteExpression)}[ \\t]+/\\$1[ \\t]+break;[ \\t]*(?:\\r?\\n)?`,
              'm'
            );
            migratedBody = migratedBody.replace(generatedRewritePattern, '');
          }
        }

        return `${startMarker}${migratedBody}${endMarker}\n`;
      }
    );

    const routeBlocks = [];
    const blockScanner = routeBlockPattern();
    let blockMatch;
    while ((blockMatch = blockScanner.exec(migratedRoutes))) {
      routeBlocks.push({
        start: blockMatch.index,
        end: blockScanner.lastIndex,
        content: blockMatch[0],
        root: /^[ \t]*location[ \t]+\/[ \t]*\{/m.test(blockMatch[0])
      });
    }
    const rootBlock = routeBlocks.find((block) => block.root);
    if (rootBlock && routeBlocks.some((block) => !block.root && block.start > rootBlock.start)) {
      const withoutRoot = `${migratedRoutes.slice(0, rootBlock.start)}${migratedRoutes.slice(rootBlock.end)}`;
      const separator = withoutRoot.endsWith('\n\n') ? '' : withoutRoot.endsWith('\n') ? '\n' : '\n\n';
      migratedRoutes = `${withoutRoot}${separator}${rootBlock.content}`;
    }
    if (migratedRoutes === managedRoutes) return content;
    return `${content.slice(0, startIndex)}${migratedRoutes}${content.slice(endIndex)}`;
  };

  const findManagedRootRouteIndex = ({ content, startIndex, endIndex }) => {
    const managedRoutes = content.slice(startIndex, endIndex);
    const routePattern =
      /^([ \t]*# BEGIN PIPELINE-GENERATOR ROUTE ([^\r\n]+)[ \t]*\r?\n)([\s\S]*?)(^[ \t]*# END PIPELINE-GENERATOR ROUTE \2[ \t]*)(?:\r?\n)?/gm;
    let match;
    while ((match = routePattern.exec(managedRoutes))) {
      if (/^[ \t]*location[ \t]+\/[ \t]*\{/m.test(match[0])) {
        return startIndex + match.index;
      }
    }
    return -1;
  };

  const replaceManagedNginxRouteAtLocation = ({
    content,
    startIndex,
    endIndex,
    location,
    legacyLocations = [],
    managedServiceKeys = [],
    managedContainerNames = [],
    routeContent
  }) => {
    const managedRoutes = content.slice(startIndex, endIndex);
    const routePattern =
      /^([ \t]*# BEGIN PIPELINE-GENERATOR ROUTE ([^\r\n]+)[ \t]*\r?\n)([\s\S]*?)(^[ \t]*# END PIPELINE-GENERATOR ROUTE \2[ \t]*)(?:\r?\n)?/gm;
    const locations = [location, ...legacyLocations];
    const locationPatterns = locations.map((candidate) => {
      const escapedLocation = String(candidate).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      return {
        location: candidate,
        pattern: new RegExp(`^[ \\t]*location[ \\t]+${escapedLocation}[ \\t]*\\{`, 'm')
      };
    });
    const identityRequired = managedServiceKeys.length > 0 || managedContainerNames.length > 0;
    const escapedContainers = managedContainerNames.map((containerName) =>
      String(containerName).replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
    );
    const hasManagedContainer = (routeBlock) => escapedContainers.some((containerName) =>
      new RegExp(
        `(?:set[ \\t]+\\$target[ \\t]+${containerName}[ \\t]*;|proxy_pass[ \\t]+http:\\/\\/${containerName}:)`,
        'm'
      ).test(routeBlock)
    );
    let replacements = 0;
    let desiredLocationConflict = false;
    const replacedRoutes = managedRoutes.replace(routePattern, (routeBlock, _start, serviceKey) => {
      const matchedLocation = locationPatterns.find(({ pattern }) => pattern.test(routeBlock));
      if (!matchedLocation) return routeBlock;
      const managedIdentity =
        !identityRequired ||
        managedServiceKeys.includes(serviceKey) ||
        hasManagedContainer(routeBlock);
      if (!managedIdentity) {
        if (matchedLocation.location === location) desiredLocationConflict = true;
        return routeBlock;
      }
      replacements += 1;
      return `${routeContent}\n`;
    });
    if (desiredLocationConflict) {
      throw new Error(
        `Nginx managed location ${location} belongs to another service; move it manually before generating the Monorepo route.`
      );
    }
    if (replacements > 1) {
      throw new Error(`Nginx managed routes contain multiple identities for location ${location}.`);
    }
    if (replacements !== 1) return { content, replaced: false };
    return {
      content: `${content.slice(0, startIndex)}${replacedRoutes}${content.slice(endIndex)}`,
      replaced: true
    };
  };

  const mergeNginxServiceRoute = ({
    content,
    serverName,
    domain,
    projectKey,
    serviceKey,
    environment,
    stack = 'default',
    routeOptions = {}
  }) => {
    let mergedContent = migrateNginxCertificatePaths({ content, serverName, domain });
    let server = findNginxHttpsServer(mergedContent, serverName);
    const route = buildNginxRouteBlock({ projectKey, serviceKey, environment, stack, ...routeOptions });

    let startIndex = mergedContent.indexOf(NGINX_MANAGED_ROUTES_START, server.open.end);
    let endIndex = mergedContent.indexOf(NGINX_MANAGED_ROUTES_END, server.open.end);
    const startInsideServer = startIndex !== -1 && startIndex < server.close.start;
    const endInsideServer = endIndex !== -1 && endIndex < server.close.start;
    if (startInsideServer !== endInsideServer || (startInsideServer && startIndex > endIndex)) {
      throw new Error('Nginx managed-route markers are incomplete or out of order.');
    }

    if (startInsideServer && endInsideServer) {
      mergedContent = normalizeNginxManagedRoutes({
        content: mergedContent,
        startIndex,
        endIndex,
        routingOverrides: routeOptions.routingOverrides
      });
      if (mergedContent !== content) {
        server = findNginxHttpsServer(mergedContent, serverName);
        startIndex = mergedContent.indexOf(NGINX_MANAGED_ROUTES_START, server.open.end);
        endIndex = mergedContent.indexOf(NGINX_MANAGED_ROUTES_END, server.open.end);
      }
    }

    const legacyManagedLocations = routeOptions.legacyManagedLocations || [];
    const hasDesiredOrLegacyLocation = [route.location, ...legacyManagedLocations]
      .some((location) => server.locations.includes(location));
    if (hasDesiredOrLegacyLocation) {
      if (routeOptions.replaceManagedLocation && startInsideServer && endInsideServer) {
        const replacement = replaceManagedNginxRouteAtLocation({
          content: mergedContent,
          startIndex,
          endIndex,
          location: route.location,
          legacyLocations: legacyManagedLocations,
          managedServiceKeys: routeOptions.managedServiceKeys || [],
          managedContainerNames: routeOptions.managedContainerNames || [],
          routeContent: route.content
        });
        if (replacement.replaced) return replacement.content;
      }
      if (server.locations.includes(route.location)) return mergedContent;
    }

    if (startInsideServer && endInsideServer) {
      const rootRouteIndex =
        route.location === '/' ? -1 : findManagedRootRouteIndex({ content: mergedContent, startIndex, endIndex });
      const insertionIndex =
        rootRouteIndex === -1 ? mergedContent.lastIndexOf('\n', endIndex) + 1 : rootRouteIndex;
      return `${mergedContent.slice(0, insertionIndex)}${route.content}\n\n${mergedContent.slice(insertionIndex)}`;
    }

    const beforeClose = mergedContent.slice(0, server.close.start);
    const separator = beforeClose.endsWith('\n') ? '\n' : '\n\n';
    const managedBlock = [
      NGINX_MANAGED_ROUTES_START,
      route.content,
      NGINX_MANAGED_ROUTES_END,
      ''
    ].join('\n');
    return `${beforeClose}${separator}${managedBlock}${mergedContent.slice(server.close.start)}`;
  };

  const buildNginxSample = ({ projectHost, projectKey, serviceKey, environment, stack = 'default', domain, routing }) => {
    const route = buildNginxRouteBlock({ projectKey, serviceKey, environment, stack, routing });
    const certificateName = domain;
    const serverName = `${projectHost}.${domain}`;
    return [
      'server {',
      '    listen 80;',
      `    server_name ${serverName};`,
      '    return 301 https://$host$request_uri;',
      '}',
      '',
      'server {',
      '    listen 443 ssl;',
      `    server_name ${serverName};`,
      '',
      '    client_max_body_size 0;',
      `    ssl_certificate /etc/nginx/conf.d/${certificateName}.pem;`,
      `    ssl_certificate_key /etc/nginx/conf.d/${certificateName}.key;`,
      '',
      NGINX_MANAGED_ROUTES_START,
      route.content,
      NGINX_MANAGED_ROUTES_END,
      '}',
      ''
    ].join('\n');
  };

  const buildMonorepoKomodoResourceNames = ({ compactProject, environment, stack = 'default' }) => {
    const normalizedStack = normalizeStackName(stack);
    const stackSuffix = isDefaultStack(normalizedStack) ? '' : `-${normalizedStack}`;
    return {
      repository: `${compactProject}_Docker_DevOps-${environment}${stackSuffix}`,
      stack: `${compactProject}_Docker_DevOps-${environment}${stackSuffix}`
    };
  };

  const buildMonorepoTagKeys = ({ serviceKey }) => {
    const staticTagKey = normalizeResourceSegment(serviceKey, 'Service name').replace(/-/g, '_');
    return {
      staticTagKey,
      bffTagKey: `${staticTagKey}_bff`
    };
  };

  const buildMonorepoEnvSample = ({ serviceKey }) => {
    const { staticTagKey, bffTagKey } = buildMonorepoTagKeys({ serviceKey });
    return [
      '# Managed image tags for the Monorepo services. Releases update only these values.',
      `${staticTagKey}:CHANGE_ME`,
      `${bffTagKey}:CHANGE_ME`,
      ''
    ].join('\n');
  };

  const mergeMonorepoEnvTags = ({ content, serviceKey }) => {
    const newline = String(content || '').includes('\r\n') ? '\r\n' : '\n';
    const hadTrailingNewline = String(content || '').endsWith('\n');
    const lines = String(content || '').replace(/\r\n/g, '\n').split('\n');
    while (lines.length && !lines[lines.length - 1].trim()) lines.pop();
    const { staticTagKey, bffTagKey } = buildMonorepoTagKeys({ serviceKey });
    for (const key of [staticTagKey, bffTagKey]) {
      const pattern = new RegExp(`^\\s*${key}\\s*[:=]`);
      if (!lines.some((line) => pattern.test(line))) {
        lines.push(`${key}:CHANGE_ME`);
      }
    }
    const merged = lines.join('\n');
    return `${merged}${hadTrailingNewline ? '\n' : ''}`.replace(/\n/g, newline);
  };

  const buildMonorepoComposeServices = ({
    projectKey,
    serviceKey,
    imageServiceKey = serviceKey,
    environment,
    stack = 'default',
    repositoryAddress,
    nginxNetworkKey = 'nginx-network'
  }) => {
    const normalizedService = serviceKey.replace(/-/g, '_');
    const normalizedStack = normalizeStackName(stack);
    const containerStackSegment = isDefaultStack(normalizedStack) ? '' : `_${normalizedStack.replace(/-/g, '_')}`;
    const imageStackSegment = isDefaultStack(normalizedStack) ? '' : `-${normalizedStack}`;
    const staticContainer = `${projectKey}_${normalizedService}${containerStackSegment}_${environment}`;
    const bffContainer = `${projectKey}_${normalizedService}_bff${containerStackSegment}_${environment}`;
    const registry = String(repositoryAddress || defaultValues.repositoryAddress)
      .trim()
      .replace(/^https?:\/\//i, '')
      .replace(/\/+$/, '')
      .toLowerCase();
    const bffProfile = `mr-${serviceKey}${imageStackSegment}-bff`;
    const staticRepository = `${registry}/${projectKey}/${imageServiceKey}${imageStackSegment}-${environment}`;
    const bffRepository = `${registry}/${projectKey}/${imageServiceKey}-bff${imageStackSegment}-${environment}`;
    const { staticTagKey, bffTagKey } = buildMonorepoTagKeys({ serviceKey });
    return [
      {
        name: staticContainer,
        content: [
      `  ${staticContainer}:`,
      `    container_name: ${staticContainer}`,
      `    image: ${staticRepository}:\${${staticTagKey}}`,
      '    restart: unless-stopped',
      '    expose:',
      '      - "80"',
      '    networks:',
      `      - ${nginxNetworkKey}`
        ].join('\n')
      },
      {
        name: bffContainer,
        content: [
      `  ${bffContainer}:`,
      `    container_name: ${bffContainer}`,
      `    image: ${bffRepository}:\${${bffTagKey}}`,
      `    profiles: ["${bffProfile}"]`,
      '    restart: unless-stopped',
      '    expose:',
      '      - "3000"',
      '    networks:',
      `      - ${nginxNetworkKey}`
        ].join('\n')
      }
    ];
  };

  const buildMonorepoComposeSample = ({
    projectKey,
    serviceKey,
    imageServiceKey = serviceKey,
    environment,
    repositoryAddress,
    stack = 'default',
    nginxNetworkName = 'nginx-network'
  }) => {
    const networkName = selectNginxNetworkName([nginxNetworkName]);
    const services = buildMonorepoComposeServices({
      projectKey,
      serviceKey,
      imageServiceKey,
      environment,
      repositoryAddress,
      stack,
      nginxNetworkKey: networkName
    });
    return [
      'services:',
      ...services.flatMap(({ content }, index) => (index ? ['', content] : [content])),
      '',
      'networks:',
      `  ${networkName}:`,
      `    name: ${networkName}`,
      '    external: true',
      ''
    ].join('\n');
  };

  const mergeMonorepoComposeServices = ({
    content,
    projectKey,
    serviceKey,
    imageServiceKey = serviceKey,
    environment,
    repositoryAddress,
    stack = 'default',
    nginxNetworkName = 'nginx-network'
  }) => {
    if (!String(content || '').trim()) {
      return buildMonorepoComposeSample({
        projectKey,
        serviceKey,
        imageServiceKey,
        environment,
        repositoryAddress,
        stack,
        nginxNetworkName
      });
    }
    const networkName = selectNginxNetworkName([nginxNetworkName]);
    const newline = content.includes('\r\n') ? '\r\n' : '\n';
    const hadTrailingNewline = content.endsWith('\n');
    const lines = content.replace(/\r\n/g, '\n').split('\n');
    let networksIndex = lines.findIndex((line) => /^networks:\s*(?:#.*)?$/.test(line));
    let networksEnd = networksIndex === -1
      ? -1
      : lines.findIndex((line, index) => index > networksIndex && /^[A-Za-z0-9_.-]+:\s*(?:.*)?$/.test(line));
    if (networksIndex !== -1 && networksEnd === -1) networksEnd = lines.length;
    const networkEntries = networksIndex === -1
      ? []
      : lines
          .slice(networksIndex + 1, networksEnd)
          .map((line, offset) => ({
            name: /^  ([A-Za-z0-9_.-]+):\s*(?:#.*)?$/.exec(line)?.[1] || '',
            index: networksIndex + 1 + offset
          }))
          .filter(({ name }) => name);
    const entryActualName = (entry) => {
      const nextEntry = networkEntries.find(({ index }) => index > entry.index)?.index ?? networksEnd;
      for (let index = entry.index + 1; index < nextEntry; index += 1) {
        const match = /^    name:\s*["']?([^\s"'#]+)["']?\s*(?:#.*)?$/.exec(lines[index]);
        if (match) return match[1];
      }
      return '';
    };
    const logicalNetworkKey =
      networkEntries.find((entry) => entryActualName(entry) === networkName)?.name ||
      networkEntries.find((entry) => entry.name === networkName)?.name ||
      networkEntries.find((entry) => ['nginx-network', 'nginx-net'].includes(entry.name))?.name ||
      networkName;
    const desiredServices = buildMonorepoComposeServices({
      projectKey,
      serviceKey,
      imageServiceKey,
      environment,
      repositoryAddress,
      stack,
      nginxNetworkKey: logicalNetworkKey
    });
    let servicesIndex = lines.findIndex((line) => /^services:\s*(?:#.*)?$/.test(line));
    const inlineEmptyIndex = lines.findIndex((line) => /^services:\s*\{\s*\}\s*(?:#.*)?$/.test(line));
    if (servicesIndex === -1 && inlineEmptyIndex !== -1) {
      lines[inlineEmptyIndex] = 'services:';
      servicesIndex = inlineEmptyIndex;
    }
    if (servicesIndex === -1) {
      while (lines.length && !lines[lines.length - 1].trim()) lines.pop();
      if (lines.length) lines.push('');
      lines.push('services:');
      servicesIndex = lines.length - 1;
    }
    let servicesEnd = lines.findIndex(
      (line, index) => index > servicesIndex && /^[A-Za-z0-9_.-]+:\s*(?:.*)?$/.test(line)
    );
    if (servicesEnd === -1) servicesEnd = lines.length;
    const existingNames = new Set(
      lines
        .slice(servicesIndex + 1, servicesEnd)
        .map((line) => /^  ([A-Za-z0-9_.-]+):\s*(?:#.*)?$/.exec(line)?.[1])
        .filter(Boolean)
    );
    for (const { name, content: serviceContent } of desiredServices) {
      if (!existingNames.has(name)) continue;
      const serviceStart = lines.findIndex(
        (line, index) => index > servicesIndex && index < servicesEnd && line === `  ${name}:`
      );
      if (serviceStart === -1) continue;
      let serviceEnd = servicesEnd;
      for (let index = serviceStart + 1; index < servicesEnd; index += 1) {
        if (/^  [A-Za-z0-9_.-]+:\s*(?:#.*)?$/.test(lines[index])) {
          serviceEnd = index;
          break;
        }
      }
      const desiredImage = /^    image:\s*([^\s#]+)\s*$/m.exec(serviceContent)?.[1] || '';
      const runtimeSuffix = name.includes('_bff_') ? 'node:20-alpine' : 'nginx:1.27-alpine';
      const legacyImages = new Set([runtimeSuffix, `registry.buluttakin.com/${runtimeSuffix}`]);
      for (let index = serviceStart + 1; index < serviceEnd; index += 1) {
        const image = /^(\s{4}image:\s*)(["']?)([^\s"'#]+)\2(\s*(?:#.*)?)$/.exec(lines[index]);
        if (!image || !legacyImages.has(image[3]) || !desiredImage) continue;
        lines[index] = `${image[1]}${image[2]}${desiredImage}${image[2]}${image[4]}`;
      }

      const migratedBlock = [lines[serviceStart]];
      for (let index = serviceStart + 1; index < serviceEnd;) {
        const line = lines[index];
        if (/^    working_dir:\s*\/srv\/monorepo\/current\/modules\//.test(line) ||
            /^    command:.*MR_[A-Z0-9_]+_BFF_ENTRY/.test(line)) {
          index += 1;
          continue;
        }
        if (/^    volumes:\s*(?:#.*)?$/.test(line)) {
          const retained = [];
          index += 1;
          while (index < serviceEnd && /^(?:      |\s*$)/.test(lines[index])) {
            const volumeLine = lines[index];
            const managedMount = /:\/srv\/monorepo(?::ro)?\s*(?:#.*)?$/.test(volumeLine) ||
              /:\/etc\/nginx\/conf\.d\/default\.conf:ro\s*(?:#.*)?$/.test(volumeLine);
            if (!managedMount) retained.push(volumeLine);
            index += 1;
          }
          if (retained.some((candidate) => candidate.trim())) {
            migratedBlock.push(line, ...retained);
          }
          continue;
        }
        migratedBlock.push(line);
        index += 1;
      }
      lines.splice(serviceStart, serviceEnd - serviceStart, ...migratedBlock);
      servicesEnd += migratedBlock.length - (serviceEnd - serviceStart);
    }
    const missing = desiredServices.filter(({ name }) => !existingNames.has(name));
    if (missing.length) {
      const insertion = missing.flatMap(({ content: serviceContent }, index) => [
        ...(index || (servicesEnd > servicesIndex + 1 && lines[servicesEnd - 1]?.trim()) ? [''] : []),
        ...serviceContent.split('\n')
      ]);
      lines.splice(servicesEnd, 0, ...insertion);
    }
    networksIndex = lines.findIndex((line) => /^networks:\s*(?:#.*)?$/.test(line));
    if (networksIndex === -1) {
      while (lines.length && !lines[lines.length - 1].trim()) lines.pop();
      lines.push('', 'networks:', `  ${logicalNetworkKey}:`, `    name: ${networkName}`, '    external: true');
    } else {
      networksEnd = lines.findIndex(
        (line, index) => index > networksIndex && /^[A-Za-z0-9_.-]+:\s*(?:.*)?$/.test(line)
      );
      if (networksEnd === -1) networksEnd = lines.length;
      let networkStart = lines.findIndex(
        (line, index) => index > networksIndex && index < networksEnd && line === `  ${logicalNetworkKey}:`
      );
      if (networkStart === -1) {
        lines.splice(networksEnd, 0, `  ${logicalNetworkKey}:`, `    name: ${networkName}`, '    external: true');
      } else {
        let networkEnd = networksEnd;
        for (let index = networkStart + 1; index < networksEnd; index += 1) {
          if (/^  [A-Za-z0-9_.-]+:\s*(?:#.*)?$/.test(lines[index])) {
            networkEnd = index;
            break;
          }
        }
        const nameIndex = lines.findIndex(
          (line, index) => index > networkStart && index < networkEnd && /^    name:\s*/.test(line)
        );
        const externalIndex = lines.findIndex(
          (line, index) => index > networkStart && index < networkEnd && /^    external:\s*/.test(line)
        );
        if (nameIndex === -1) {
          lines.splice(networkStart + 1, 0, `    name: ${networkName}`);
          networkEnd += 1;
        } else {
          lines[nameIndex] = `    name: ${networkName}`;
        }
        const adjustedExternalIndex = externalIndex === -1
          ? -1
          : externalIndex + (nameIndex === -1 && externalIndex > networkStart ? 1 : 0);
        if (adjustedExternalIndex === -1) {
          lines.splice(networkEnd, 0, '    external: true');
        } else {
          lines[adjustedExternalIndex] = '    external: true';
        }
      }
    }
    const merged = lines.join('\n').replace(/\n+$/, '');
    return `${merged}${hadTrailingNewline ? '\n' : ''}`.replace(/\n/g, newline);
  };

  const buildMonorepoNginxRoutes = ({
    projectKey,
    serviceKey,
    environment,
    stack = 'default',
    routing,
    routingOverrides
  }) => {
    const normalizedService = serviceKey.replace(/-/g, '_');
    const normalizedStack = normalizeStackName(stack);
    const stackSegment = isDefaultStack(normalizedStack) ? '' : `_${normalizedStack.replace(/-/g, '_')}`;
    const bffContainer = `${projectKey}_${normalizedService}_bff${stackSegment}_${environment}`;
    return [
    {
      serviceKey: `${serviceKey}-bff`,
      routeOptions: {
        containerName: bffContainer,
        location: '/bff/',
        internalPort: 3000,
        frontend: false,
        replaceManagedLocation: true,
        legacyManagedLocations: ['/api/'],
        managedServiceKeys: [`${serviceKey}-bff`],
        managedContainerNames: [
          bffContainer,
          `${projectKey}_mr_bff_${environment}`
        ],
        routingOverrides
      }
    },
    {
      serviceKey,
      routeOptions: {
        containerName: `${projectKey}_${normalizedService}${stackSegment}_${environment}`,
        location: routing?.location || '/',
        internalPort: 80,
        frontend: true,
        replaceManagedLocation: true,
        routing,
        routingOverrides
      }
    }
    ];
  };

  const buildMonorepoNginxSample = ({
    projectHost,
    projectKey,
    serviceKey,
    environment,
    stack = 'default',
    domain,
    routing,
    routingOverrides
  }) => {
    const certificateName = domain;
    const serverName = `${projectHost}.${domain}`;
    const routes = buildMonorepoNginxRoutes({ projectKey, serviceKey, environment, stack, routing, routingOverrides })
      .map(({ serviceKey, routeOptions }) =>
        buildNginxRouteBlock({ projectKey, serviceKey, environment, ...routeOptions }).content
      );
    return [
      'server {',
      '    listen 80;',
      `    server_name ${serverName};`,
      '    return 301 https://$host$request_uri;',
      '}',
      '',
      'server {',
      '    listen 443 ssl;',
      `    server_name ${serverName};`,
      '',
      '    client_max_body_size 0;',
      `    ssl_certificate /etc/nginx/conf.d/${certificateName}.pem;`,
      `    ssl_certificate_key /etc/nginx/conf.d/${certificateName}.key;`,
      '',
      NGINX_MANAGED_ROUTES_START,
      ...routes.flatMap((route, index) => (index ? ['', route] : [route])),
      NGINX_MANAGED_ROUTES_END,
      '}',
      ''
    ].join('\n');
  };

  const mergeMonorepoNginxRoutes = ({
    content,
    serverName,
    domain,
    projectKey,
    serviceKey,
    environment,
    stack = 'default',
    routing,
    routingOverrides
  }) =>
    buildMonorepoNginxRoutes({ projectKey, serviceKey, environment, stack, routing, routingOverrides }).reduce(
      (merged, { serviceKey, routeOptions }) => mergeNginxServiceRoute({
        content: merged,
        serverName,
        domain,
        projectKey,
        serviceKey,
        environment,
        stack,
        routeOptions
      }),
      content
    );

  const buildMonorepoSupportRepositorySpecs = ({
    projectName,
    environment,
    domain,
    stack = 'default',
    service,
    repositoryAddress,
    nginxNetworkName = 'nginx-network',
    includeNginx = true,
    serviceRouting,
    routingOverrides
  }) => {
    const compactProject = String(projectName || '').replace(/\s+/g, '');
    const normalizedEnvironment = normalizeResourceSegment(environment, 'Environment');
    const shouldIncludeNginx = includeNginx;
    if (!compactProject || /[\\/\0\r\n]/.test(compactProject)) {
      throw new Error('Project name cannot be converted to a safe DevOps repository name.');
    }
    if (shouldIncludeNginx && !/^(?=.{1,253}$)(?:[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?\.)+[a-z]{2,63}$/i.test(domain || '')) {
      throw new Error(`Environment ${environment || '(empty)'} has no valid domain.`);
    }
    const projectKey = normalizeResourceSegment(compactProject.toLowerCase(), 'Project key');
    const serviceKey = normalizeResourceSegment(service, 'Service name');
    const imageServiceKey = normalizeImageServiceSegment(service, 'Service name');
    const projectHost = projectKey;
    const composeDirectory = buildComposeDirectory({ environment: normalizedEnvironment, stack, projectName: compactProject });
    const nginxDirectory = buildNginxDirectory({ environment: normalizedEnvironment, stack });
    return [
      {
        kind: 'docker',
        name: `${compactProject}_Docker_DevOps`,
        directory: composeDirectory,
        filePath: `/${composeDirectory}/compose.yml`,
        content: buildMonorepoComposeSample({
          projectKey,
          serviceKey,
          imageServiceKey,
          environment: normalizedEnvironment,
          stack,
          repositoryAddress,
          nginxNetworkName
        }),
        mergeExisting: (content) => mergeMonorepoComposeServices({
          content,
          projectKey,
          serviceKey,
          imageServiceKey,
          environment: normalizedEnvironment,
          stack,
          repositoryAddress,
          nginxNetworkName
        }),
        additionalFiles: [
          {
            path: `/${composeDirectory}/.env`,
            content: buildMonorepoEnvSample({ serviceKey }),
            mergeExisting: (content) => mergeMonorepoEnvTags({ content, serviceKey })
          }
        ]
      },
      ...(shouldIncludeNginx ? [{
        kind: 'nginx',
        name: `${compactProject}_Nginx_DevOps`,
        directory: nginxDirectory,
        filePath: `/${nginxDirectory}/${projectHost}-${normalizedEnvironment}.conf`,
        content: buildMonorepoNginxSample({
          projectHost,
          projectKey,
          serviceKey,
          environment: normalizedEnvironment,
          stack,
          domain: String(domain).toLowerCase(),
          routing: serviceRouting,
          routingOverrides
        }),
        mergeExisting: (content) => mergeMonorepoNginxRoutes({
          content,
          serverName: `${projectHost}.${String(domain).toLowerCase()}`,
          domain: String(domain).toLowerCase(),
          projectKey,
          serviceKey,
          environment: normalizedEnvironment,
          stack,
          routing: serviceRouting,
          routingOverrides
        })
      }] : [])
    ];
  };

  const buildSupportRepositorySpecs = ({
    projectName,
    environment,
    domain,
    stack = 'default',
    service,
    projectsRoot,
    repositoryAddress,
    nginxNetworkName = 'nginx-network',
    serviceRouting,
    includeNginx = true,
    routingOverrides
  }) => {
    const compactProject = String(projectName || '').replace(/\s+/g, '');
    const normalizedEnvironment = normalizeResourceSegment(environment, 'Environment');
    const shouldIncludeNginx = includeNginx;
    const serviceKey = normalizeResourceSegment(service, 'Service name');
    if (!compactProject || /[\\/\0\r\n]/.test(compactProject)) {
      throw new Error('Project name cannot be converted to a safe DevOps repository name.');
    }
    if (shouldIncludeNginx && !/^(?=.{1,253}$)(?:[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?\.)+[a-z]{2,63}$/i.test(domain || '')) {
      throw new Error(`Environment ${environment || '(empty)'} has no valid domain.`);
    }
    const compactProjectLower = compactProject.toLowerCase();
    const projectHost = normalizeResourceSegment(compactProjectLower, 'Project hostname');
    const imageServiceKey = normalizeImageServiceSegment(service, 'Service name');
    const composeDirectory = buildComposeDirectory({ environment: normalizedEnvironment, stack, projectName: compactProject });
    const nginxDirectory = buildNginxDirectory({ environment: normalizedEnvironment, stack });
    return [
      {
        kind: 'docker',
        name: `${compactProject}_Docker_DevOps`,
        directory: composeDirectory,
        filePath: `/${composeDirectory}/compose.yml`,
        content: buildComposeSample({
          projectKey: compactProjectLower,
          serviceKey,
          imageServiceKey,
          environment: normalizedEnvironment,
          stack,
          repositoryAddress,
          nginxNetworkName,
          routing: serviceRouting
        })
      },
      ...(shouldIncludeNginx ? [{
        kind: 'nginx',
        name: `${compactProject}_Nginx_DevOps`,
        directory: nginxDirectory,
        filePath: `/${nginxDirectory}/${projectHost}-${normalizedEnvironment}.conf`,
        content: buildNginxSample({
          projectHost,
          projectKey: compactProjectLower,
          serviceKey,
          environment: normalizedEnvironment,
          stack,
          domain: String(domain).toLowerCase(),
          routing: serviceRouting
        }),
        mergeExisting: (content) => mergeNginxServiceRoute({
          content,
          serverName: `${projectHost}.${String(domain).toLowerCase()}`,
          domain: String(domain).toLowerCase(),
          projectKey: compactProjectLower,
          serviceKey,
          environment: normalizedEnvironment,
          stack,
          routeOptions: {
            routing: serviceRouting,
            routingOverrides
          }
        })
      }] : [])
    ];
  };

  const ensureRepositoryBootstrapFiles = async ({
    hostUri,
    projectId,
    repo,
    directory,
    sampleFile,
    additionalFiles = [],
    accessToken
  }) => {
    const branchName = SCAFFOLD_BRANCH;
    const branchRef = `refs/heads/${branchName}`;
    const oldObjectId = await getBranchObjectId({
      hostUri,
      projectId,
      repoId: repo.id,
      branch: branchName,
      accessToken
    });
    const desiredFiles = [sampleFile, ...additionalFiles];
    let existingFiles = desiredFiles.map(() => null);
    if (oldObjectId !== ZERO_OBJECT_ID) {
      existingFiles = await Promise.all(
        desiredFiles.map((file) => getRepositoryFileContent({
          hostUri,
          projectId,
          repoId: repo.id,
          branchName,
          path: file.path,
          accessToken
        }))
      );
    }
    const changes = [];
    desiredFiles.forEach((file, index) => {
      const existing = existingFiles[index];
      if (existing === null) {
        changes.push({ ...file, changeType: 'add' });
        return;
      }
      if (typeof file.mergeExisting === 'function') {
        const merged = file.mergeExisting(existing);
        if (merged !== existing) {
          changes.push({ ...file, content: merged, changeType: 'edit' });
        }
      }
    });
    if (!changes.length) {
      return { skipped: true, unchanged: true };
    }

    const url = `${hostUri}${encodeURIComponent(projectId)}/_apis/git/repositories/${repo.id}/pushes?api-version=6.0`;
    const body = {
      refUpdates: [{ name: branchRef, oldObjectId }],
      commits: [
        {
          comment: `Initialize ${repo.name} for ${directory}`,
          changes: changes.map((file) => ({
            changeType: file.changeType,
            item: { path: file.path },
            newContent: { content: file.content, contentType: 'rawtext' }
          }))
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
      const detail = await readErrorDetail(res);
      throw buildHttpError(`Failed to initialize repository ${repo.name}`, res, detail);
    }
    return { skipped: false, paths: changes.map((file) => file.path) };
  };

  const ensureSupportRepositories = async ({
    hostUri,
    projectId,
    projectName,
    environment,
    domain,
    stack = 'default',
    service,
    projectsRoot,
    repositoryAddress,
    nginxNetworkName,
    accessToken,
    includeNginx = true,
    mode = 'pipeline',
    serviceRouting,
    routingOverrides
  }) => {
    const specs = normalizeGeneratorMode(mode) === 'monorepo'
      ? buildMonorepoSupportRepositorySpecs({
          projectName,
          environment,
          stack,
          domain,
          service,
          repositoryAddress,
          nginxNetworkName,
          includeNginx,
          serviceRouting,
          routingOverrides
        })
      : buildSupportRepositorySpecs({
          projectName,
          environment,
          stack,
          domain,
          service,
          repositoryAddress,
          nginxNetworkName,
          includeNginx,
          serviceRouting,
          routingOverrides
        });
    const results = [];
    for (const spec of specs) {
      const repo = await ensureRepositoryByName({
        hostUri,
        projectId,
        repositoryName: spec.name,
        accessToken
      });
      const bootstrap = await ensureRepositoryBootstrapFiles({
        hostUri,
        projectId,
        repo,
        directory: spec.directory,
        sampleFile: {
          path: spec.filePath,
          content: spec.content,
          mergeExisting: spec.mergeExisting
        },
        additionalFiles: spec.additionalFiles || [],
        accessToken
      });
      await ensureDefaultBranch({
        hostUri,
        projectId,
        repoId: repo.id,
        branchName: SCAFFOLD_BRANCH,
        accessToken
      });
      results.push({
        kind: spec.kind,
        repo,
        directory: spec.directory,
        filePath: spec.filePath,
        bootstrap
      });
    }
    return results;
  };

  const postScaffold = async ({
    hostUri,
    projectId,
    repoId,
    branch,
    accessToken,
    content,
    pipelineFilename = 'project-repo-environment-branch.yml'
  }) => {
    const branchName = SCAFFOLD_BRANCH;
    const branchRef = `refs/heads/${branchName}`;
    const url = `${hostUri}${encodeURIComponent(projectId)}/_apis/git/repositories/${repoId}/pushes?api-version=6.0`;
    const pipelineContent = content || '';
    const oldObjectId = await getBranchObjectId({ hostUri, projectId, repoId, branch: branchName, accessToken });
    const filePath = `/${pipelineFilename}`;
    const existingContent =
      oldObjectId === ZERO_OBJECT_ID
        ? null
        : await getRepositoryFileContent({
          hostUri,
          projectId,
          repoId,
        branchName,
        path: filePath,
          accessToken
        });
    const fileExists = existingContent !== null;
    if (fileExists && existingContent === pipelineContent) {
      return { skipped: true, unchanged: true, path: filePath, branch: branchRef };
    }
    const body = {
      refUpdates: [
        {
          name: branchRef,
          oldObjectId
        }
      ],
      commits: [
        {
          comment: `${fileExists ? 'Update' : 'Add'} pipeline generator defaults`,
          changes: [
            {
              changeType: fileExists ? 'edit' : 'add',
              item: { path: filePath },
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
      const detail = await readErrorDetail(res);
      const error = new Error(`Failed to push scaffold (${res.status})${detail ? `: ${detail}` : ''}`);
      error.status = res.status;
      error.detail = detail;
      throw error;
    }
  };

  const postGeneratedFiles = async ({ hostUri, projectId, repoId, accessToken, files, comment }) => {
    const branchName = SCAFFOLD_BRANCH;
    const branchRef = `refs/heads/${branchName}`;
    const oldObjectId = await getBranchObjectId({ hostUri, projectId, repoId, branch: branchName, accessToken });
    const normalizedFiles = files.map((file) => ({
      ...file,
      path: file.path.startsWith('/') ? file.path : `/${file.path}`,
      content: String(file.content || '')
    }));
    const existing = oldObjectId === ZERO_OBJECT_ID
      ? normalizedFiles.map(() => null)
      : await Promise.all(normalizedFiles.map((file) => getRepositoryFileContent({
          hostUri,
          projectId,
          repoId,
          branchName,
          path: file.path,
          accessToken
        })));
    const changes = [];
    normalizedFiles.forEach((file, index) => {
      const previous = existing[index];
      if (previous === null) {
        changes.push({ ...file, changeType: 'add' });
      } else if (file.overwrite !== false && previous !== file.content) {
        changes.push({ ...file, changeType: 'edit' });
      }
    });
    if (!changes.length) return { skipped: true, unchanged: true, paths: [] };
    const url = `${hostUri}${encodeURIComponent(projectId)}/_apis/git/repositories/${repoId}/pushes?api-version=6.0`;
    const res = await fetch(url, {
      method: 'POST',
      headers: { ...authHeaders(accessToken), 'Content-Type': 'application/json' },
      body: JSON.stringify({
        refUpdates: [{ name: branchRef, oldObjectId }],
        commits: [{
          comment: comment || 'Generate Monorepo pipeline files',
          changes: changes.map((file) => ({
            changeType: file.changeType,
            item: { path: file.path },
            newContent: { content: file.content, contentType: 'rawtext' }
          }))
        }]
      })
    });
    if (!res.ok) {
      const detail = await readErrorDetail(res);
      throw buildHttpError('Failed to save generated Monorepo files', res, detail);
    }
    return { skipped: false, paths: changes.map((file) => file.path) };
  };

  const buildPipelineConfiguration = ({ repoId, repositoryName, pipelinePath, branch }) => ({
    type: 'yaml',
    path: pipelinePath.startsWith('/') ? pipelinePath : `/${pipelinePath}`,
    repository: {
      id: repoId,
      name: repositoryName,
      type: 'azureReposGit',
      defaultBranch: `refs/heads/${branch}`
    }
  });

  const buildPipelinesApiUrl = ({ hostUri, projectId, pipelineId, repositoryId }) => {
    const pipelineSegment = pipelineId ? `/${encodeURIComponent(pipelineId)}` : '';
    const searchParams = new URLSearchParams();
    if (repositoryId) {
      searchParams.set('repositoryId', repositoryId);
    }
    searchParams.set('api-version', PIPELINE_API_VERSION);
    return `${hostUri}${encodeURIComponent(projectId)}/_apis/pipelines${pipelineSegment}?${searchParams.toString()}`;
  };

  const readPipelineResponse = async (res, failureMessage) => {
    if (!res.ok) {
      const detail = await readErrorDetail(res);
      throw markErrorDomain(buildHttpError(failureMessage, res, detail), 'pipeline');
    }

    const pipeline = await res.json();
    if (!pipeline?.id) {
      throw markErrorDomain(
        new Error(`${failureMessage.replace(/^Failed to /, 'Azure DevOps reported success for ')} but returned no pipeline ID.`),
        'pipeline'
      );
    }
    return pipeline;
  };

  const getPipelineByName = async ({ hostUri, projectId, pipelineName, legacyPipelineNames = [], accessToken }) => {
    const url = buildPipelinesApiUrl({ hostUri, projectId });
    const res = await fetch(url, { headers: authHeaders(accessToken) });
    if (!res.ok) {
      const detail = await readErrorDetail(res);
      throw markErrorDomain(buildHttpError('Failed to list Azure Pipelines', res, detail), 'pipeline');
    }
    const payload = await res.json();
    const pipelines = payload.value || [];
    const desired = pipelines.find((pipeline) => pipeline.name === pipelineName);
    if (desired) return desired;
    for (const legacyPipelineName of legacyPipelineNames) {
      if (legacyPipelineName && legacyPipelineName !== pipelineName) {
        const legacy = pipelines.find((pipeline) => pipeline.name === legacyPipelineName);
        if (legacy) return legacy;
      }
    }
    return undefined;
  };

  const normalizeComparableFolder = (folder) => normalizePipelineFolder(folder, '\\').toLowerCase();

  const normalizeComparableYamlPath = (path = '') => {
    const normalized = String(path || '').trim().replace(/\\/g, '/');
    return normalized.startsWith('/') ? normalized : `/${normalized}`;
  };

  const buildDefinitionsApiUrl = ({ hostUri, projectId, definitionId, query = {} }) => {
    const definitionSegment = definitionId ? `/${encodeURIComponent(definitionId)}` : '';
    const searchParams = new URLSearchParams(query);
    searchParams.set('api-version', BUILD_API_VERSION);
    return `${hostUri}${encodeURIComponent(projectId)}/_apis/build/definitions${definitionSegment}?${searchParams.toString()}`;
  };

  const readBuildDefinitionResponse = async (res, failureMessage) => {
    if (!res.ok) {
      const detail = await readErrorDetail(res);
      throw markErrorDomain(buildHttpError(failureMessage, res, detail), 'pipeline');
    }

    const definition = await res.json();
    if (!definition?.id) {
      throw markErrorDomain(new Error(`${failureMessage} but Azure DevOps returned no Build Definition ID.`), 'pipeline');
    }
    return definition;
  };

  const getBuildDefinitionById = async ({ hostUri, projectId, definitionId, accessToken }) => {
    const url = buildDefinitionsApiUrl({ hostUri, projectId, definitionId });
    const res = await fetch(url, { headers: authHeaders(accessToken) });
    return readBuildDefinitionResponse(res, `Failed to load Build Definition ${definitionId}`);
  };

  const getBuildDefinitionByYamlPath = async ({
    hostUri,
    projectId,
    repoId,
    pipelinePath,
    accessToken
  }) => {
    const desiredPath = normalizeComparableYamlPath(pipelinePath);
    const url = buildDefinitionsApiUrl({
      hostUri,
      projectId,
      query: {
        repositoryId: repoId,
        repositoryType: 'TfsGit',
        yamlFilename: desiredPath,
        includeAllProperties: 'true'
      }
    });
    const res = await fetch(url, { headers: authHeaders(accessToken) });
    if (!res.ok) {
      const detail = await readErrorDetail(res);
      throw markErrorDomain(buildHttpError('Failed to find a Pipeline by its YAML path', res, detail), 'pipeline');
    }

    const payload = await res.json();
    const references = [...(payload.value || [])].sort((left, right) => Number(left.id) - Number(right.id));
    for (const reference of references) {
      const definition =
        reference?.process?.yamlFilename && reference?.repository?.id
          ? reference
          : await getBuildDefinitionById({
              hostUri,
              projectId,
              definitionId: reference.id,
              accessToken
            });
      const samePath = normalizeComparableYamlPath(definition?.process?.yamlFilename) === desiredPath;
      const sameRepository = String(definition?.repository?.id || '') === String(repoId);
      if (samePath && sameRepository) {
        return definition;
      }
    }
    return undefined;
  };

  const updateBuildDefinition = async ({
    hostUri,
    projectId,
    definition,
    repo,
    pipelineName,
    desiredConfig,
    pipelineFolder = PIPELINE_FOLDER,
    accessToken
  }) => {
    // Azure DevOps requires the current revision and recommends GET-modify-PUT
    // with the complete Build Definition document. A list response can include
    // a revision without containing every field, so always fetch the full
    // definition immediately before the update.
    const current = await getBuildDefinitionById({
      hostUri,
      projectId,
      definitionId: definition.id,
      accessToken
    });
    const updated = {
      ...current,
      name: pipelineName,
      path: pipelineFolder,
      comment: 'Updated by Pipeline Generator.',
      process: {
        ...(current.process || {}),
        type: current.process?.type ?? 2,
        yamlFilename: desiredConfig.path
      },
      repository: {
        ...(current.repository || {}),
        id: desiredConfig.repository.id,
        name: repo.name,
        type: current.repository?.type || 'TfsGit',
        defaultBranch: desiredConfig.repository.defaultBranch
      }
    };
    const url = buildDefinitionsApiUrl({ hostUri, projectId, definitionId: current.id });
    const res = await fetch(url, {
      method: 'PUT',
      headers: {
        ...authHeaders(accessToken),
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(updated)
    });
    return readBuildDefinitionResponse(res, `Failed to update Build Definition ${current.id}`);
  };

  const pipelineBindingMatches = ({ pipeline, pipelineName, desiredConfig, pipelineFolder = PIPELINE_FOLDER }) => {
    const configuration = pipeline?.configuration;
    const repository = configuration?.repository || pipeline?.repository;
    const yamlPath = configuration?.path || pipeline?.process?.yamlFilename;
    const folder = pipeline?.folder || pipeline?.path;
    return (
      pipeline?.name === pipelineName &&
      normalizeComparableFolder(folder) === normalizeComparableFolder(pipelineFolder) &&
      normalizeComparableYamlPath(yamlPath) === normalizeComparableYamlPath(desiredConfig.path) &&
      String(repository?.id || '') === String(desiredConfig.repository.id) &&
      repository?.defaultBranch === desiredConfig.repository.defaultBranch
    );
  };

  const upsertPipelineDefinition = async ({
    hostUri,
    projectId,
    repo,
    pipelineName,
    pipelinePath,
    legacyPipelineNames = [],
    legacyPipelinePaths = [],
    branch,
    pipelineFolder = PIPELINE_FOLDER,
    accessToken
  }) => {
    const repositoryName = `${state.projectName || projectId}/${repo.name}`;
    const desiredConfig = buildPipelineConfiguration({
      repoId: repo.id,
      repositoryName,
      pipelinePath,
      branch
    });

    const existing = await getPipelineByName({
      hostUri,
      projectId,
      pipelineName,
      legacyPipelineNames,
      accessToken
    });
    if (existing?.id) {
      // The Pipelines API's by-ID response can omit repository.defaultBranch
      // on Azure DevOps Server. Treating that sparse response as a mismatch
      // caused a no-op rerun to PUT the Build Definition and increment its
      // revision. Read the canonical full Build Definition before deciding
      // whether migration is necessary.
      const current = await getBuildDefinitionById({
        hostUri,
        projectId,
        definitionId: existing.id,
        accessToken
      });
      if (pipelineBindingMatches({ pipeline: current, pipelineName, desiredConfig, pipelineFolder })) {
        return current || existing;
      }
      return updateBuildDefinition({
        hostUri,
        projectId,
        definition: existing,
        repo,
        pipelineName,
        desiredConfig,
        pipelineFolder,
        accessToken
      });
    }

    const existingForYaml = await getBuildDefinitionByYamlPath({
      hostUri,
      projectId,
      repoId: repo.id,
      pipelinePath: desiredConfig.path,
      accessToken
    });
    if (existingForYaml?.id) {
      const alreadyDesired = pipelineBindingMatches({
        pipeline: existingForYaml,
        pipelineName,
        desiredConfig,
        pipelineFolder
      });
      if (alreadyDesired) {
        return existingForYaml;
      }
      return updateBuildDefinition({
        hostUri,
        projectId,
        definition: existingForYaml,
        repo,
        pipelineName,
        desiredConfig,
        pipelineFolder,
        accessToken
      });
    }

    for (const legacyPipelinePath of legacyPipelinePaths) {
      if (!legacyPipelinePath || normalizeComparableYamlPath(legacyPipelinePath) === desiredConfig.path) continue;
      const legacyForYaml = await getBuildDefinitionByYamlPath({
        hostUri,
        projectId,
        repoId: repo.id,
        pipelinePath: legacyPipelinePath,
        accessToken
      });
      if (legacyForYaml?.id) {
        return updateBuildDefinition({
          hostUri,
          projectId,
          definition: legacyForYaml,
          repo,
          pipelineName,
          desiredConfig,
          pipelineFolder,
          accessToken
        });
      }
    }

    const createUrl = buildPipelinesApiUrl({
      hostUri,
      projectId,
      repositoryId: repo.id
    });
    const res = await fetch(createUrl, {
      method: 'POST',
      headers: {
        ...authHeaders(accessToken),
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({ name: pipelineName, folder: pipelineFolder, configuration: desiredConfig })
    });

    return readPipelineResponse(res, 'Failed to create pipeline');
  };

  const buildReleaseDefinitionsApiUrl = ({ hostUri, projectId, definitionId, query = {} }) => {
    const definitionSegment = definitionId ? `/${encodeURIComponent(definitionId)}` : '';
    const searchParams = new URLSearchParams(query);
    searchParams.set('api-version', RELEASE_API_VERSION);
    return `${hostUri}${encodeURIComponent(projectId)}/_apis/release/definitions${definitionSegment}?${searchParams.toString()}`;
  };

  const readReleaseDefinitionResponse = async (res, failureMessage) => {
    if (!res.ok) {
      const detail = await readErrorDetail(res);
      const error = markErrorDomain(buildHttpError(failureMessage, res, detail), 'release');
      error.responseDetail = detail;
      throw error;
    }
    const definition = await res.json();
    if (!definition?.id) {
      throw markErrorDomain(new Error(`${failureMessage} but Azure DevOps returned no Release Definition ID.`), 'release');
    }
    return definition;
  };

  const getReleaseDefinitionByName = async ({ hostUri, projectId, releaseName, accessToken }) => {
    const searchParams = new URLSearchParams({
      searchText: releaseName,
      isExactNameMatch: 'true',
      searchTextContainsFolderName: 'false',
      '$top': '100'
    });
    const url = buildReleaseDefinitionsApiUrl({
      hostUri,
      projectId,
      query: Object.fromEntries(searchParams.entries())
    });
    const res = await fetch(url, { headers: authHeaders(accessToken) });
    if (!res.ok) {
      const detail = await readErrorDetail(res);
      throw markErrorDomain(buildHttpError('Failed to list classic Release definitions', res, detail), 'release');
    }
    const payload = await res.json();
    return (payload.value || []).find((definition) => definition.name === releaseName);
  };

  const getReleaseDefinitionById = async ({ hostUri, projectId, definitionId, accessToken }) => {
    const url = buildReleaseDefinitionsApiUrl({ hostUri, projectId, definitionId });
    const res = await fetch(url, { headers: authHeaders(accessToken) });
    return readReleaseDefinitionResponse(res, `Failed to load classic Release definition ${definitionId}`);
  };

  const getReleaseDefinitionByPipelineId = async ({
    hostUri,
    projectId,
    pipelineId,
    accessToken
  }) => {
    const url = buildReleaseDefinitionsApiUrl({
      hostUri,
      projectId,
      query: {
        '$expand': 'Artifacts',
        artifactType: 'Build',
        artifactSourceId: `${projectId}:${pipelineId}`,
        '$top': '100'
      }
    });
    const res = await fetch(url, { headers: authHeaders(accessToken) });
    if (!res.ok) {
      const detail = await readErrorDetail(res);
      throw markErrorDomain(buildHttpError('Failed to find a Release definition by Pipeline artifact', res, detail), 'release');
    }
    const payload = await res.json();
    return (payload.value || []).find((definition) =>
      (definition.artifacts || []).some(
        (artifact) => String(artifact?.definitionReference?.definition?.id || '') === String(pipelineId)
      )
    );
  };

  const resolveReleaseAgentQueue = async ({ hostUri, projectId, queueName, accessToken }) => {
    if (!queueName) {
      throw markErrorDomain(new Error('No agent queue was selected for the classic Release job.'), 'release');
    }

    const url = `${hostUri}${encodeURIComponent(projectId)}/_apis/distributedtask/queues?api-version=6.0`;
    const res = await fetch(url, { headers: authHeaders(accessToken) });
    if (!res.ok) {
      const detail = await readErrorDetail(res);
      throw markRequiredExtensionScope(
        markErrorDomain(buildHttpError(`Failed to load agent queue ${queueName}`, res, detail), 'release'),
        'vso.agentpools'
      );
    }

    const payload = await res.json();
    const queue = (payload.value || []).find((item) => item.name === queueName);
    if (!queue?.id) {
      throw markErrorDomain(
        new Error(`Release agent queue was not found: ${queueName}. Check Project settings → Agent pools/queues.`),
        'release'
      );
    }
    return queue;
  };

  const resolveReleaseScriptRepository = async ({ hostUri, source, accessToken }) => {
    const scriptProject = source.project || state.projectId;
    const repositoryName = source.repository;
    const rawPath = source.path;
    const branch = source.branch || SCAFFOLD_BRANCH;
    if (!scriptProject || !repositoryName || !rawPath) {
      throw markErrorDomain(
        new Error('Release script source is incomplete. azureReposFile requires project, repository, and path.'),
        'release'
      );
    }

    const repositoriesUrl = `${hostUri}${encodeURIComponent(scriptProject)}/_apis/git/repositories?api-version=6.0`;
    const repositoriesResponse = await fetch(repositoriesUrl, { headers: authHeaders(accessToken) });
    if (!repositoriesResponse.ok) {
      const detail = await readErrorDetail(repositoriesResponse);
      throw markErrorDomain(buildHttpError('Failed to list the configured Release script repository', repositoriesResponse, detail), 'release');
    }

    const repositories = await repositoriesResponse.json();
    const scriptRepository = (repositories.value || []).find(
      (item) => item.id === repositoryName || item.name === repositoryName
    );
    if (!scriptRepository?.id) {
      throw markErrorDomain(
        new Error(`Release script repository was not found: ${scriptProject}/${repositoryName}.`),
        'release'
      );
    }

    const scriptPath = rawPath.startsWith('/') ? rawPath : `/${rawPath}`;
    const fileUrl = `${hostUri}${encodeURIComponent(scriptProject)}/_apis/git/repositories/${encodeURIComponent(
      scriptRepository.id
    )}/items?path=${encodeURIComponent(scriptPath)}&versionDescriptor.version=${encodeURIComponent(
      branch
    )}&versionDescriptor.versionType=branch&%24format=text&api-version=6.0`;
    const fileResponse = await fetch(fileUrl, { headers: authHeaders(accessToken) });
    if (!fileResponse.ok) {
      const detail = await readErrorDetail(fileResponse);
      throw markErrorDomain(buildHttpError(`Failed to load Release Bash script ${scriptPath}`, fileResponse, detail), 'release');
    }

    const script = await fileResponse.text();
    if (!script.trim()) {
      throw markErrorDomain(new Error(`Release Bash script is empty: ${scriptProject}/${repositoryName}${scriptPath}.`), 'release');
    }
    return script;
  };

  const resolveReleaseInlineScript = async ({ releaseConfig, hostUri, accessToken }) => {
    const source = releaseConfig.scriptSource || {};
    if (source.type === 'inline') {
      if (typeof source.content !== 'string' || !source.content.trim()) {
        throw markErrorDomain(
          new Error('Release Bash inline script is empty. Set scriptSource.content in dist/release-config.js.'),
          'release'
        );
      }
      return source.content;
    }
    if (source.type === 'packagedFile') {
      const packagedPath = String(source.path || '').trim();
      if (!packagedPath) {
        throw markErrorDomain(
          new Error('Release Bash packagedFile path is empty in dist/release-config.js.'),
          'release'
        );
      }
      const packagedUrl = new URL(packagedPath, window.location.href).toString();
      const packagedResponse = await fetch(packagedUrl, { cache: 'no-store' });
      if (!packagedResponse.ok) {
        const detail = await readErrorDetail(packagedResponse);
        throw markErrorDomain(
          buildHttpError(`Failed to load packaged Release Bash script ${packagedPath}`, packagedResponse, detail),
          'release'
        );
      }
      const script = await packagedResponse.text();
      if (!script.trim()) {
        throw markErrorDomain(new Error(`Packaged Release Bash script is empty: ${packagedPath}.`), 'release');
      }
      return script;
    }
    if (source.type === 'azureReposFile') {
      return resolveReleaseScriptRepository({ hostUri, source, accessToken });
    }
    throw markErrorDomain(
      new Error(
        `Unsupported Release script source type: ${source.type || 'missing'}. Use inline, packagedFile, or azureReposFile.`
      ),
      'release'
    );
  };

  const resolveReleaseVariableGroup = async ({
    hostUri,
    projectId,
    groupName,
    requiredVariableNames,
    accessToken
  }) => {
    if (!groupName) {
      throw markErrorDomain(new Error('Release variableGroupName is empty in dist/release-config.js.'), 'release');
    }
    const searchParams = new URLSearchParams({
      groupName,
      actionFilter: 'Use',
      'api-version': '7.1'
    });
    const url = `${hostUri}${encodeURIComponent(projectId)}/_apis/distributedtask/variablegroups?${searchParams}`;
    const res = await fetch(url, { headers: authHeaders(accessToken) });
    if (!res.ok) {
      const detail = await readErrorDetail(res);
      throw markRequiredExtensionScope(
        markErrorDomain(buildHttpError(`Failed to load Release variable group ${groupName}`, res, detail), 'release'),
        'vso.variablegroups_read'
      );
    }

    const payload = await res.json();
    const group = (payload.value || [])
      .filter((item) => item?.name === groupName && item?.id != null)
      .sort((left, right) => Number(left.id) - Number(right.id))[0];
    if (!group) {
      throw markErrorDomain(
        new Error(`Release variable group was not found or cannot be used: ${groupName}. Check Pipelines → Library.`),
        'release'
      );
    }

    const variables = group.variables || {};
    const missingVariables = (requiredVariableNames || []).filter(
      (name) => !Object.prototype.hasOwnProperty.call(variables, name)
    );
    if (missingVariables.length) {
      throw markErrorDomain(
        new Error(
          `Release variable group ${groupName} is missing required variables: ${missingVariables.join(', ')}.`
        ),
        'release'
      );
    }
    return { id: Number(group.id), name: group.name };
  };

  const buildReleaseDefinitionPayload = ({
    releaseName,
    releaseConfig,
    inlineScript,
    projectId,
    projectName,
    repo,
    pipelineDefinition,
    pipelineName,
    branch,
    agentQueue,
    variableGroup
  }) => {
    const defaultBranch = `refs/heads/${branch}`;
    const artifactAlias = `_${pipelineName}`;
    return {
      name: releaseName,
      path: releaseConfig.folder,
      description: 'Generated by Pipeline Generator.',
      releaseNameFormat: 'Release-$(rev:r)',
      artifacts: [
        {
          sourceId: `${projectId}:${pipelineDefinition.id}`,
          type: 'Build',
          alias: artifactAlias,
          definitionReference: {
            definition: { id: String(pipelineDefinition.id), name: pipelineName },
            project: { id: projectId, name: projectName },
            repository: { id: repo.id, name: repo.name },
            defaultVersionBranch: { id: defaultBranch, name: defaultBranch },
            defaultVersionType: { id: 'latestType', name: 'Latest' },
            defaultVersionSpecific: { id: '', name: '' },
            defaultVersionTags: { id: '', name: '' },
            artifactSourceDefinitionUrl: { id: '', name: '' }
          },
          isPrimary: true,
          isRetained: false
        }
      ],
      environments: [
        {
          name: releaseConfig.environmentName,
          rank: 1,
          variables: {},
          variableGroups: [],
          demands: [],
          conditions: [{ name: 'ReleaseStarted', conditionType: 'event', value: '', result: null }],
          executionPolicy: { concurrencyCount: 0, queueDepthCount: 0 },
          schedules: [],
          retentionPolicy: { daysToKeep: 30, releasesToKeep: 3, retainBuild: true },
          processParameters: {},
          preDeployApprovals: {
            approvals: [{ rank: 1, isAutomated: true, isNotificationOn: false }],
            approvalOptions: {
              requiredApproverCount: null,
              releaseCreatorCanBeApprover: false,
              autoTriggeredAndPreviousEnvironmentApprovedCanBeSkipped: false,
              enforceIdentityRevalidation: false,
              timeoutInMinutes: 0,
              executionOrder: 'beforeGates'
            }
          },
          postDeployApprovals: {
            approvals: [{ rank: 1, isAutomated: true, isNotificationOn: false }],
            approvalOptions: {
              requiredApproverCount: null,
              releaseCreatorCanBeApprover: false,
              autoTriggeredAndPreviousEnvironmentApprovedCanBeSkipped: false,
              enforceIdentityRevalidation: false,
              timeoutInMinutes: 0,
              executionOrder: 'afterSuccessfulGates'
            }
          },
          deployPhases: [
            {
              name: 'Agent job',
              phaseType: 'agentBasedDeployment',
              rank: 1,
              workflowTasks: [
                {
                  taskId: BASH_TASK_ID,
                  version: '3.*',
                  name: releaseConfig.bashTaskName,
                  refName: '',
                  enabled: true,
                  alwaysRun: false,
                  continueOnError: false,
                  timeoutInMinutes: 0,
                  definitionType: 'task',
                  condition: 'succeeded()',
                  inputs: {
                    targetType: 'inline',
                    script: inlineScript,
                    workingDirectory: '',
                    failOnStderr: 'false',
                    noProfile: 'true',
                    noRc: 'true'
                  }
                }
              ],
              deploymentInput: {
                queueId: Number(agentQueue.id),
                queueName: agentQueue.name,
                demands: [],
                enableAccessToken: false,
                skipArtifactsDownload: false,
                timeoutInMinutes: 0,
                jobCancelTimeoutInMinutes: 1,
                condition: 'succeeded()',
                overrideInputs: {},
                parallelExecution: { parallelExecutionType: 'none' },
                artifactsDownloadInput: { downloadInputs: [] }
              }
            }
          ]
        }
      ],
      variables: {},
      variableGroups: [Number(variableGroup.id)],
      triggers: [],
      properties: {}
    };
  };

  const getReleaseWorkflowTask = (definition) =>
    definition?.environments?.[0]?.deployPhases?.[0]?.workflowTasks?.[0];

  const getReleaseDeploymentInput = (definition) =>
    definition?.environments?.[0]?.deployPhases?.[0]?.deploymentInput;

  const releaseDefinitionMatches = ({ current, desired }) => {
    const currentArtifact = current?.artifacts?.[0];
    const desiredArtifact = desired?.artifacts?.[0];
    const currentEnvironment = current?.environments?.[0];
    const desiredEnvironment = desired?.environments?.[0];
    const currentTask = getReleaseWorkflowTask(current);
    const desiredTask = getReleaseWorkflowTask(desired);
    const currentDeployment = getReleaseDeploymentInput(current);
    const desiredDeployment = getReleaseDeploymentInput(desired);
    const currentVariableGroupIds = new Set(
      (current?.variableGroups || []).map((groupId) => String(groupId))
    );
    const hasDesiredVariableGroups = (desired?.variableGroups || []).every((groupId) =>
      currentVariableGroupIds.has(String(groupId))
    );
    return (
      current?.name === desired?.name &&
      normalizeComparableFolder(current?.path) === normalizeComparableFolder(desired?.path) &&
      String(currentArtifact?.definitionReference?.definition?.id || '') ===
        String(desiredArtifact?.definitionReference?.definition?.id || '') &&
      String(currentArtifact?.definitionReference?.repository?.id || '') ===
        String(desiredArtifact?.definitionReference?.repository?.id || '') &&
      currentEnvironment?.name === desiredEnvironment?.name &&
      Number(currentDeployment?.queueId) === Number(desiredDeployment?.queueId) &&
      currentTask?.taskId === desiredTask?.taskId &&
      currentTask?.version === desiredTask?.version &&
      currentTask?.name === desiredTask?.name &&
      currentTask?.inputs?.targetType === 'inline' &&
      currentTask?.inputs?.script === desiredTask?.inputs?.script &&
      hasDesiredVariableGroups &&
      currentEnvironment?.conditions?.some((condition) => condition?.name === 'ReleaseStarted') &&
      currentEnvironment?.preDeployApprovals?.approvals?.some((approval) => approval?.isAutomated === true) &&
      currentEnvironment?.postDeployApprovals?.approvals?.some((approval) => approval?.isAutomated === true)
    );
  };

  const updateReleaseDefinition = async ({
    hostUri,
    projectId,
    definition,
    desired,
    accessToken
  }) => {
    const current = await getReleaseDefinitionById({
      hostUri,
      projectId,
      definitionId: definition.id,
      accessToken
    });
    if (releaseDefinitionMatches({ current, desired })) {
      return { ...current, created: false, updated: false };
    }
    const mergedEnvironments = (desired.environments || []).map((desiredEnvironment, index) => {
      const currentEnvironment = current.environments?.[index] || {};
      const mergedDeployPhases = (desiredEnvironment.deployPhases || []).map((desiredPhase, phaseIndex) => ({
        ...(currentEnvironment.deployPhases?.[phaseIndex] || {}),
        ...desiredPhase
      }));
      return {
        ...currentEnvironment,
        ...desiredEnvironment,
        ...(currentEnvironment.id ? { id: currentEnvironment.id } : {}),
        deployPhases: mergedDeployPhases
      };
    });
    const mergedVariableGroups = Array.from(
      new Set(
        [...(current.variableGroups || []), ...(desired.variableGroups || [])]
          .map((groupId) => Number(groupId))
          .filter(Number.isFinite)
      )
    );
    const updated = {
      ...current,
      ...desired,
      id: current.id,
      revision: current.revision,
      comment: 'Updated by Pipeline Generator.',
      environments: mergedEnvironments,
      variableGroups: mergedVariableGroups
    };
    const url = buildReleaseDefinitionsApiUrl({ hostUri, projectId });
    const res = await fetch(url, {
      method: 'PUT',
      headers: {
        ...authHeaders(accessToken),
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(updated)
    });
    const result = await readReleaseDefinitionResponse(
      res,
      `Failed to update classic Release definition ${current.id}`
    );
    return { ...result, created: false, updated: true };
  };

  const ensureReleaseDefinition = async ({
    hostUri,
    projectId,
    projectName,
    repo,
    pipelineDefinition,
    pipelineName,
    service,
    environment,
    stack = 'default',
    branch,
    queueName,
    accessToken,
    mode = 'pipeline'
  }) => {
    const releaseConfig = getReleaseConfig(mode);
    if (!releaseConfig.enabled) {
      return { skipped: true, reason: 'Release creation is disabled in dist/release-config.js.' };
    }
    if (!releaseConfig.environmentName) {
      throw markErrorDomain(new Error('Release environmentName is empty in dist/release-config.js.'), 'release');
    }
    if (!releaseConfig.variableGroupName) {
      throw markErrorDomain(new Error('Release variableGroupName is empty in dist/release-config.js.'), 'release');
    }

    const releaseName = buildReleaseName({ service, environment, stack, mode });
    const existingByName = await getReleaseDefinitionByName({ hostUri, projectId, releaseName, accessToken });
    const existingByPipeline = existingByName?.id
      ? undefined
      : await getReleaseDefinitionByPipelineId({
          hostUri,
          projectId,
          pipelineId: pipelineDefinition.id,
          accessToken
        });

    const [inlineScript, agentQueue, variableGroup] = await Promise.all([
      resolveReleaseInlineScript({ releaseConfig, hostUri, accessToken }),
      resolveReleaseAgentQueue({ hostUri, projectId, queueName, accessToken }),
      resolveReleaseVariableGroup({
        hostUri,
        projectId,
        groupName: releaseConfig.variableGroupName,
        requiredVariableNames: releaseConfig.requiredVariableNames,
        accessToken
      })
    ]);
    const body = buildReleaseDefinitionPayload({
      releaseName,
      releaseConfig,
      inlineScript,
      projectId,
      projectName,
      repo,
      pipelineDefinition,
      pipelineName,
      branch,
      agentQueue,
      variableGroup
    });
    const existing = existingByName || existingByPipeline;
    if (existing?.id) {
      return updateReleaseDefinition({
        hostUri,
        projectId,
        definition: existing,
        desired: body,
        accessToken
      });
    }

    const url = buildReleaseDefinitionsApiUrl({ hostUri, projectId });
    const res = await fetch(url, {
      method: 'POST',
      headers: {
        ...authHeaders(accessToken),
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(body)
    });
    if (!res.ok) {
      const detail = await readErrorDetail(res);
      if (res.status === 409 || /already exists|same name|duplicate/i.test(detail)) {
        const duplicate = await getReleaseDefinitionByName({ hostUri, projectId, releaseName, accessToken });
        if (duplicate?.id) {
          return updateReleaseDefinition({
            hostUri,
            projectId,
            definition: duplicate,
            desired: body,
            accessToken
          });
        }
      }
      throw markErrorDomain(buildHttpError(`Failed to create classic Release definition ${releaseName}`, res, detail), 'release');
    }
    return { ...(await res.json()), created: true, updated: false };
  };

  const fetchAgentQueues = async ({ hostUri, projectId, accessToken }) => {
    const url = `${hostUri}${encodeURIComponent(projectId)}/_apis/distributedtask/queues?api-version=6.0`;
    const res = await fetch(url, { headers: authHeaders(accessToken) });
    if (!res.ok) {
      const detail = await readErrorDetail(res);
      throw buildHttpError('Failed to load pools', res, detail);
    }
    const payload = await res.json();
    return Array.from(new Set((payload.value || []).map((queue) => queue.name).filter(Boolean)));
  };

  const fetchContainerRegistries = async ({ hostUri, projectId, accessToken }) => {
    const url = `${hostUri}${encodeURIComponent(projectId)}/_apis/serviceendpoint/endpoints?type=dockerregistry&projectIds=${encodeURIComponent(
      projectId
    )}&api-version=6.0`;
    const res = await fetch(url, { headers: authHeaders(accessToken) });
    if (!res.ok) {
      const detail = await readErrorDetail(res);
      throw buildHttpError('Failed to load container registries', res, detail);
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

  const buildMonorepoDeploymentContract = () => [
    'version: 1',
    'kind: nx-monorepo',
    'install_command: "pnpm install --frozen-lockfile"',
    'build_command: "node tools/scripts/with-env.cjs production pnpm exec nx run-many -t build --projects={projects} --parallel=3"',
    'artifact_name: "mr-drop"',
    'shell_projects:',
    '  - "shell"',
    '  - "host"',
    'bff_projects:',
    '  - "bff"',
    'bff_entry: "main.js"',
    'continue_on_module_error: true',
    'orphan_policy: "retain"',
    '',
    '# The extension creates this contract once and preserves later edits.',
    '# New Nx applications are discovered automatically. Each affected module is built independently.',
    '# A failed module keeps its previous deployed version; a failed shell blocks deployment.',
    '# A shell change rebuilds every application.',
    '# Removed or renamed applications are retained as orphans for manual review; nothing is auto-deleted.',
    ''
  ].join('\n');

  const quoteYaml = (value) => `'${String(value || '').replace(/'/g, "''")}'`;

  const buildMonorepoPipelineYaml = (payload, options = {}) => {
    const sourceBranchName = (options.sourceBranch || 'main').replace(/^refs\/heads\//, '');
    const sourceRepositoryName =
      options.rawRepositoryName || options.repositoryName || options.sourceRepositoryName || 'repository';
    const projectName = options.rawProjectName || options.projectName || 'PROJECTNAME';
    const compactProject = String(projectName).replace(/\s+/g, '');
    const projectKey = normalizeResourceSegment(compactProject.toLowerCase(), 'Project key');
    const serviceKey = normalizeResourceSegment(
      payload.service || deriveServiceNameFromRepository(sourceRepositoryName, projectName),
      'Service name'
    );
    const environment = normalizeResourceSegment(payload.environment, 'Environment');
    const stack = normalizeStackName(payload.stack);
    const containerStackSegment = isDefaultStack(stack) ? '' : `_${stack.replace(/-/g, '_')}`;
    const imageStackSegment = isDefaultStack(stack) ? '' : `-${stack}`;
    const registryAddress = String(payload.repositoryAddress || defaultValues.repositoryAddress)
      .trim()
      .replace(/^https?:\/\//i, '')
      .replace(/\/+$/, '')
      .toLowerCase();
    const normalizedService = serviceKey.replace(/-/g, '_');
    const staticContainer = `${projectKey}_${normalizedService}${containerStackSegment}_${environment}`;
    const bffContainer = `${projectKey}_${normalizedService}_bff${containerStackSegment}_${environment}`;
    const runtimeVariablePrefix = `MR_${projectKey}_${serviceKey}${imageStackSegment}_${environment}`
      .replace(/[^a-z0-9]+/gi, '_')
      .toUpperCase();
    const bffProfile = `mr-${serviceKey}${imageStackSegment}-bff`;
    const composeRepository = `${compactProject}_Docker_DevOps`;
    const composePath = `/${buildComposeDirectory({ environment, stack, projectName: compactProject })}/compose.yml`;
    const komodoResources = buildMonorepoKomodoResourceNames({ compactProject, environment, stack });
    return [
      'trigger: none',
      '',
      'resources:',
      '  repositories:',
      '    - repository: SharedTemplatesRepo',
      '      type: git',
      '      endpoint: ShonizCollection',
      '      name: SharedTemplates/SharedTemplates',
      '      ref: refs/heads/main',
      '',
      '    - repository: sourceRepo',
      '      type: git',
      `      name: ${quoteYaml(`${projectName}/${sourceRepositoryName}`)}`,
      `      ref: ${quoteYaml(`refs/heads/${sourceBranchName}`)}`,
      '      trigger:',
      '        branches:',
      '          include:',
      `            - ${quoteYaml(sourceBranchName)}`,
      '',
      'variables:',
      '  - group: KomodoAPI',
      '',
      'stages:',
      '  - template: monorepo/pipeline.yml@SharedTemplatesRepo',
      '    parameters:',
      `      pool: ${quoteYaml(payload.pool)}`,
      `      projectKey: ${quoteYaml(projectKey)}`,
      `      serviceKey: ${quoteYaml(serviceKey)}`,
      `      environment: ${quoteYaml(environment)}`,
      `      stack: ${quoteYaml(stack)}`,
      `      komodoServer: ${quoteYaml(payload.komodoServer)}`,
      `      staticContainer: ${quoteYaml(staticContainer)}`,
      `      bffContainer: ${quoteYaml(bffContainer)}`,
      `      runtimeVariablePrefix: ${quoteYaml(runtimeVariablePrefix)}`,
      `      bffProfile: ${quoteYaml(bffProfile)}`,
      `      composeRepository: ${quoteYaml(composeRepository)}`,
      `      composePath: ${quoteYaml(composePath)}`,
      `      komodoRepository: ${quoteYaml(komodoResources.repository)}`,
      `      komodoStack: ${quoteYaml(komodoResources.stack)}`,
      `      registryAddress: ${quoteYaml(registryAddress)}`,
      `      containerRegistryService: ${quoteYaml(payload.containerRegistryService || defaultValues.containerRegistryService)}`,
      `      staticRuntimeImage: ${quoteYaml(`${registryAddress}/nginx:1.27-alpine`)}`,
      `      bffRuntimeImage: ${quoteYaml(`${registryAddress}/node:20-alpine`)}`,
      `      nodeImage: ${quoteYaml(`${registryAddress}/node:22-bookworm`)}`,
      ''
    ].join('\n');
  };

  const buildPipelineYaml = (payload, options = {}) => {
    const sourceBranchName = (options.sourceBranch || 'main').replace(/^refs\/heads\//, '');
    const sourceRepositoryName =
      options.rawRepositoryName || options.repositoryName || options.sourceRepositoryName || 'repository';
    const projectName = options.rawProjectName || options.projectName || 'PROJECTNAME';
    const projectRepoName = `${projectName}/${sourceRepositoryName}`;
    return [
      "trigger: none",
      '',
      'resources:',
      '  repositories:',
      '    - repository: SharedTemplatesRepo',
      '      type: git',
      '      endpoint: ShonizCollection',
      '      name: SharedTemplates/SharedTemplates',
      '      ref: main',
      '',
      `    - repository: otherRepo`,
      '      type: git',
      `      name: "${projectRepoName}"`,
      `      ref: refs/heads/${sourceBranchName}`,
      '      trigger:',
      '        branches:',
      '          include:',
      `            - ${sourceBranchName}`,
      '#        paths:',
      '#          exclude:',
      '#            - server/**',
      '#          include:',
      '#            - client/**',
      '',
      'variables:',
      '- group: KomodoAPI',
      '',
      'stages:',
      '- template: build-push-komodo.yml@SharedTemplatesRepo',
      '  parameters:',
      `    pool: '${payload.pool || ''}'`,
      `    service: '${payload.service || ''}'                # service name`,
      `    environment: '${payload.environment || ''}'           # selected deployment environment`,
      `    stack: '${normalizeStackName(payload.stack)}'             # default preserves legacy names`,
      `    dockerfileDir: '${payload.dockerfileDir || '**'}'  # path of Dockerfile, Default is '**'`,
      `    repositoryAddress: '${payload.repositoryAddress || ''}'`,
      `    containerRegistryService: '${payload.containerRegistryService || ''}'`,
      "    tag: '1.0.$(Build.BuildId)'",
      `    komodoServer: '${payload.komodoServer || ''}' # selected Komodo server`,
      "    komodoApiKey: '$(KOMODO_API_KEY)'",
      "    komodoApiSecret: '$(KOMODO_API_SECRET)'",
      '    sourceRepo: otherRepo',
      ''
    ].join('\n');
  };

  const handleSubmit = async (event) => {
    event.preventDefault();
    if (state.provisioningComplete) return;
    setCompletionVisibility(false);
    if (initializationPromise) {
      try {
        await initializationPromise;
      } catch (error) {
        console.error('Initialization failed before submit', error);
      }
    }
    if (!state.deploymentTargetsReady) {
      setStatus(
        `Deployment targets are unavailable. Verify ${DEPLOYMENT_TARGETS_CONFIG.collection}/${DEPLOYMENT_TARGETS_CONFIG.project}/${DEPLOYMENT_TARGETS_CONFIG.repository}:${DEPLOYMENT_TARGETS_CONFIG.path} on ${DEPLOYMENT_TARGETS_CONFIG.branch} and reopen the generator.`,
        true
      );
      setSubmitting(false);
      return;
    }
    const payload = Object.fromEntries(new FormData(form).entries());
    payload.environment = String(payload.environment || '').trim();
    payload.stack = normalizeStackName(payload.stack);
    payload.service = normalizeServiceNameForForm(payload.service);
    if (serviceInput) serviceInput.value = payload.service;
    setEditableValue(environmentSelect, environmentCombobox, payload.environment);
    setEditableValue(stackInput, stackCombobox, payload.stack);
    const environmentConfig = state.deploymentTargets?.environmentConfigs?.find(
      ({ name }) => name.toLowerCase() === String(payload.environment || '').toLowerCase()
    );
    const includeNginx = Boolean(environmentConfig?.domain);
    payload.projectsRoot = environmentConfig?.projectsRoot || defaultProjectsRootForEnvironment(payload.environment);
    const generatorOptions = {
      sourceBranch: state.sourceBranch,
      rawProjectName: state.rawProjectName,
      projectName: state.projectName,
      rawRepositoryName: state.rawRepositoryName,
      repositoryName: state.repositoryName,
      sourceRepositoryName: state.rawRepositoryName || state.repositoryName || state.projectName
    };
    const yaml = isMonorepoMode()
      ? buildMonorepoPipelineYaml(payload, generatorOptions)
      : buildPipelineYaml(payload, generatorOptions);

    setStatus('Generating pipeline template...');
    setSubmitting(true);

    if (!state.accessToken) {
      const errorMessage = state.accessTokenError || 'Azure DevOps did not issue an extension access token.';
      const message = buildTokenRecoveryMessage(errorMessage);
      setReauthenticationVisibility(true, message);
      setStatus(message, true);
      setSubmitting(false);
      return;
    }

    if (!state.projectId && state.sdk?.getWebContext) {
      const context = state.sdk.getWebContext();
      state.projectId = context?.project?.id || state.projectId;
      const contextProjectName = context?.project?.name;
      state.rawProjectName = contextProjectName || state.rawProjectName;
      state.projectName = contextProjectName || state.projectId || state.projectName;
      state.repoId = context?.repository?.id || state.repoId;
      const contextRepositoryName = context?.repository?.name;
      state.rawRepositoryName = contextRepositoryName || state.rawRepositoryName;
      state.repositoryName = contextRepositoryName || state.repositoryName;
    }

    const sourceRepositoryName = state.rawRepositoryName || state.repositoryName || state.projectName;
    const pipelineFilename = buildPipelineFilename({
      projectName: state.projectName,
      repositoryName: sourceRepositoryName,
      service: payload.service,
      environment: payload.environment,
      branchName: state.sourceBranch,
      stack: payload.stack,
      mode: state.mode
    });
    const legacyServiceLessPipelineFilename = buildLegacyServiceLessPipelineFilename({
      projectName: state.projectName,
      repositoryName: sourceRepositoryName,
      environment: payload.environment,
      branchName: state.sourceBranch,
      mode: state.mode
    });
    const legacyPipelineFilename = buildLegacyPipelineFilename({
      projectName: state.projectName,
      repositoryName: sourceRepositoryName,
      branchName: state.sourceBranch
    });
    const legacyEnvironmentFirstPipelineFilename = buildLegacyEnvironmentFirstPipelineFilename({
      projectName: state.projectName,
      repositoryName: sourceRepositoryName,
      environment: payload.environment,
      branchName: state.sourceBranch
    });
    const pipelineName = buildPipelineName(pipelineFilename);
    const releaseName = buildReleaseName({
      service: payload.service,
      environment: payload.environment,
      stack: payload.stack,
      mode: state.mode
    });
    const pipelineFolder = getPipelineFolder();

    if (!state.accessToken || !state.projectId) {
      setStatus('Open the extension from Azure DevOps to create the repositories, files, Pipeline, and Release.', true);
      setSubmitting(false);
      return yaml;
    }

    const targetBranch = SCAFFOLD_BRANCH;
    try {
      const provisioningProjectName = state.rawProjectName || state.projectName;
      let supportRepositories = [];
      let nginxNetworkName = '';
      let serviceRoutingPlan;
      const repo = await runProvisioningStep(
        'Step 1/5: resolving the target Nginx network and creating or reusing the DevOps repositories...',
        async () => {
          nginxNetworkName = await resolveNginxNetworkForServer({
            hostUri: state.hostUri,
            server: payload.komodoServer
          });
          const projectRepositories = await listProjectRepositories({
            hostUri: state.hostUri,
            projectId: state.projectId,
            accessToken: state.accessToken
          });
          serviceRoutingPlan = buildProjectServiceRoutingPlan({
            repositories: projectRepositories,
            projectName: provisioningProjectName,
            repositoryName: sourceRepositoryName,
            service: payload.service
          });
          const pipelineRepo = await ensureRepo({
            hostUri: state.hostUri,
            projectId: state.projectId,
            projectName: provisioningProjectName,
            accessToken: state.accessToken
          });
          supportRepositories = await ensureSupportRepositories({
            hostUri: state.hostUri,
            projectId: state.projectId,
            projectName: provisioningProjectName,
            environment: payload.environment,
            domain: environmentConfig?.domain,
            stack: payload.stack,
            service: payload.service,
            projectsRoot: payload.projectsRoot,
            repositoryAddress: payload.repositoryAddress,
            nginxNetworkName,
            includeNginx,
            accessToken: state.accessToken,
            mode: state.mode,
            serviceRouting: serviceRoutingPlan.current.routing,
            routingOverrides: serviceRoutingPlan.routingOverrides
          });
          return pipelineRepo;
        }
      );
      state.generatedRepoId = repo.id || state.generatedRepoId;
      state.generatedRepositoryName = repo.name || state.generatedRepositoryName;
      state.branch = targetBranch;
      await runProvisioningStep(
        isMonorepoMode()
          ? `Step 2/5: saving MR Pipeline and /.devops deployment files...`
          : `Step 2/5: saving YAML file /${pipelineFilename}...`,
        async () => {
          if (!isMonorepoMode()) {
            return postScaffold({
              hostUri: state.hostUri,
              projectId: state.projectId,
              repoId: repo.id,
              branch: targetBranch,
              accessToken: state.accessToken,
              content: yaml,
              pipelineFilename
            });
          }
          return postGeneratedFiles({
            hostUri: state.hostUri,
            projectId: state.projectId,
            repoId: repo.id,
            accessToken: state.accessToken,
            comment: `Generate MR Pipeline ${pipelineFilename}`,
            files: [
              { path: `/${pipelineFilename}`, content: yaml },
              { path: '/.devops/deployments.yml', content: buildMonorepoDeploymentContract(), overwrite: false }
            ]
          });
        }
      );
      await runProvisioningStep('Step 3/5: setting the generated repository default branch...', () =>
        ensureDefaultBranch({
          hostUri: state.hostUri,
          projectId: state.projectId,
          repoId: repo.id,
          branchName: targetBranch,
          accessToken: state.accessToken
        })
      );

      const pipelineDefinition = await runProvisioningStep(
        `Step 4/5: creating or updating Pipeline ${pipelineName} in ${pipelineFolder}...`,
        () =>
          upsertPipelineDefinition({
            hostUri: state.hostUri,
            projectId: state.projectId,
            repo,
            pipelineName,
            pipelinePath: `/${pipelineFilename}`,
            legacyPipelineNames: !isDefaultStack(payload.stack)
              ? []
              : isMonorepoMode()
                ? [legacyServiceLessPipelineFilename]
                : [
                    legacyServiceLessPipelineFilename,
                    legacyEnvironmentFirstPipelineFilename,
                    legacyPipelineFilename
                  ],
            legacyPipelinePaths: !isDefaultStack(payload.stack)
              ? []
              : isMonorepoMode()
                ? [`/${legacyServiceLessPipelineFilename}`]
                : [
                    `/${legacyServiceLessPipelineFilename}`,
                    `/${legacyEnvironmentFirstPipelineFilename}`,
                    `/${legacyPipelineFilename}`
                  ],
            branch: targetBranch,
            pipelineFolder,
            accessToken: state.accessToken
          })
      );

      const releaseDefinition = await runProvisioningStep(
        `Step 5/5: creating or reusing the classic Release definition ${releaseName}...`,
        () =>
          ensureReleaseDefinition({
            hostUri: state.hostUri,
            projectId: state.projectId,
            projectName: state.rawProjectName || state.projectName,
            repo,
            pipelineDefinition,
            pipelineName,
            service: payload.service,
            environment: payload.environment,
            branch: targetBranch,
            stack: payload.stack,
            queueName: payload.pool,
            accessToken: state.accessToken,
            mode: state.mode
          })
      );

      const releaseMessage = releaseDefinition.skipped
        ? `Release skipped: ${releaseDefinition.reason}`
        : `Release definition ${
            releaseDefinition.created ? 'created' : releaseDefinition.updated ? 'updated' : 'already up to date'
          } (ID: ${releaseDefinition.id}).`;
      const generatedFilesMessage = includeNginx
        ? `${isMonorepoMode() ? 'deployment contract, ' : ''}Nginx and Compose files`
        : `${isMonorepoMode() ? 'deployment contract and ' : ''}Compose file; Nginx was skipped because ${payload.environment} is not a configured Environment`;
      setStatus(
        `Done. Pipeline ${pipelineName} is linked to /${pipelineFilename} in ${pipelineFolder} (ID: ${pipelineDefinition?.id || 'unknown'}). ${releaseMessage} Review the generated ${generatedFilesMessage} below, then run the Pipeline manually.`,
        false
      );
      showCompletionLinks({
        supportRepositories,
        pipelineDefinition,
        generatedRepo: repo,
        contractPath: isMonorepoMode() ? '/.devops/deployments.yml' : undefined
      });
    } catch (error) {
      console.error(error);
      const detail = sanitizeErrorDetail(error?.detail || error?.message || '');
      const step = error?.provisioningStep || 'Provisioning';
      const permissionHint = error?.requiredExtensionScope
        ? ` The installed extension token is missing or has not been reauthorized for ${error.requiredExtensionScope}. A Collection Administrator must authorize the updated Pipeline Generator scopes.`
        : error?.domain === 'komodo'
          ? ' Verify the selected Komodo server, central read credential, CORS policy, and that nginx-net or nginx-network exists on the target Docker host.'
        : error?.domain === 'release'
          ? ' Ask a project administrator to grant Manage release definitions, View releases, and Use the selected agent queue.'
          : error?.domain === 'pipeline'
            ? ' Ask a project administrator to grant Create/Edit pipeline permission under Project settings → Pipelines → Security.'
            : ' Ask a project administrator to grant the required Repos permissions.';
      const unauthorizedMessage = `Access denied during ${step}.${detail ? ` Details: ${detail}.` : ''}${permissionHint}`;
      const detailMessage = `Failed during ${step}.${detail ? ` Details: ${detail}` : ' No error details were returned by Azure DevOps.'}`;
      if (isUnauthorizedError(error)) {
        state.accessToken = null;
        setReauthenticationVisibility(
          true,
          error?.requiredExtensionScope
            ? `The installed extension token cannot access ${error.requiredExtensionScope}. Use Open extension authorization as a Collection Administrator, authorize the updated scopes, then reopen the generator from the branch.`
            : 'The Azure DevOps host token was denied. Sign out and authenticate again, then reopen the generator from the branch.'
        );
      }
      setStatus(isUnauthorizedError(error) ? unauthorizedMessage : detailMessage, true);
    }

    setSubmitting(false);
    return yaml;
  };

  form?.addEventListener('submit', handleSubmit);

  const fetchDockerfileDirectories = async ({ hostUri, projectId, repoId, branch, accessToken }) => {
    if (!repoId) return [];
    const versionDescriptor = branch
      ? `&versionDescriptor.version=${encodeURIComponent(branch)}&versionDescriptor.versionType=branch`
      : '';
    const url = `${hostUri}${encodeURIComponent(projectId)}/_apis/git/repositories/${repoId}/items?recursionLevel=Full&includeContentMetadata=true${versionDescriptor}&api-version=6.0`;
    const res = await fetch(url, { headers: authHeaders(accessToken) });
    if (!res.ok) {
      const detail = await readErrorDetail(res);
      throw buildHttpError('Failed to scan repository for Dockerfiles', res, detail);
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
    const modeFromQuery = normalizeGeneratorMode(getQueryValue(query.get('mode')));
    applyModePresentation(modeFromQuery);
    const branchFromQuery = getQueryValue(query.get('branch'));
    const projectIdFromQuery = getQueryValue(query.get('projectId'));
    const projectNameFromQuery = getQueryValue(query.get('projectName')) || projectIdFromQuery;
    const repoIdFromQuery = getQueryValue(query.get('repoId'));
    const repoNameFromQuery = getQueryValue(query.get('repoName'));
    const initialBranch = branchFromQuery || '(unknown branch)';
    const isFramed = window.parent !== window;

    const hostLooksLikeAzureDevOps = (() => {
      const candidateOrigins = new Set();
      const addOrigin = (value) => {
        try {
          if (value) {
            candidateOrigins.add(new URL(value).origin);
          }
        } catch {
          /* ignore invalid URLs */
        }
      };

      addOrigin(window.location.origin);
      addOrigin(document.referrer);
      if (window.location.ancestorOrigins) {
        try {
          const rawAncestors = window.location.ancestorOrigins;
          const ancestors = [];

          if (typeof rawAncestors.forEach === 'function') {
            rawAncestors.forEach((value) => ancestors.push(value));
          } else {
            const length = Number(rawAncestors.length) || 0;
            for (let i = 0; i < length; i += 1) {
              ancestors.push(rawAncestors[i]);
            }
          }

          ancestors.forEach(addOrigin);
        } catch (error) {
          console.warn('Skipping ancestorOrigins inspection', error);
        }
      }
      if (candidateOrigins.size === 0) return false;

      return Array.from(candidateOrigins).some((origin) => {
        try {
          const { hostname } = new URL(origin);
          return (
            origin === window.location.origin ||
            hostname.toLowerCase().endsWith('dev.azure.com') ||
            hostname.toLowerCase().endsWith('visualstudio.com')
          );
        } catch {
          return false;
        }
      });
    })();

    // Only attempt SDK initialization when the extension is running inside the
    // Azure DevOps dialog or hub iframe. Direct asset URLs remain in offline
    // mode to avoid noisy VSS handshake errors.
    const shouldAttemptSdk = isFramed && hostLooksLikeAzureDevOps;

    state.sourceBranch = initialBranch;
    state.projectId = projectIdFromQuery;
    state.rawProjectName = projectNameFromQuery;
    state.projectName = projectNameFromQuery;
    state.repoId = repoIdFromQuery;
    state.rawRepositoryName = repoNameFromQuery;
    state.repositoryName = repoNameFromQuery;
    state.hostUri = `${getHostBase().replace(/\/+$/, '')}/`;

    branchLabel.textContent = branchFromQuery
      ? `Target branch: ${SCAFFOLD_BRANCH} (source: ${initialBranch})`
      : 'Loading branch context...';
    if (branchInput && branchFromQuery) {
      branchInput.value = SCAFFOLD_BRANCH;
      branchInput.disabled = true;
    }
    targetRepoInput.value = `${projectNameFromQuery || 'project'}_Azure_DevOps`;
    setServiceNameFromRepository(repoNameFromQuery || projectNameFromQuery, projectNameFromQuery);
    // A host dialog is a trusted candidate even when an on-premises
    // referrer-policy removes document.referrer. The VSS handshake itself is
    // the authority; an unrelated parent cannot complete it successfully.
    const hasHostContext = isFramed;
    if (!hasHostContext || !shouldAttemptSdk) {
      setStatus(
        'Running outside Azure DevOps. Open the extension from a branch action to create the repository and pipeline file automatically.',
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
      const dialogConfiguration = getDialogConfiguration(sdk);
      const hostNavigationState = await getHostNavigationState(sdk);
      const hostedConfiguration = { ...hostNavigationState, ...dialogConfiguration };
      applyModePresentation(hostedConfiguration.mode || modeFromQuery);

      const branch =
        getQueryValue(hostedConfiguration.branch) ||
        branchFromQuery ||
        context?.repository?.defaultBranch?.replace(/^refs\/heads\//, '') ||
        '(unknown branch)';
      state.sourceBranch = branch;

      const projectId = getQueryValue(hostedConfiguration.projectId) || projectIdFromQuery || context?.project?.id;
      const projectName =
        getQueryValue(hostedConfiguration.projectName) || projectNameFromQuery || context?.project?.name || projectId;
      const repoId = getQueryValue(hostedConfiguration.repoId) || repoIdFromQuery || context?.repository?.id;
      let repositoryName =
        getQueryValue(hostedConfiguration.repoName) || repoNameFromQuery || context?.repository?.name;
      state.sdk = sdk;
      state.projectId = projectId;
      state.rawProjectName = projectName || state.rawProjectName;
      state.projectName = projectName;
      state.repoId = repoId;
      state.rawRepositoryName = repositoryName || state.rawRepositoryName;
      state.repositoryName = repositoryName;

      branchLabel.textContent = `Target branch: ${SCAFFOLD_BRANCH} (source: ${branch})`;
      if (branchInput) {
        branchInput.value = SCAFFOLD_BRANCH;
        branchInput.disabled = true;
      }
      targetRepoInput.value = `${projectName || 'project'}_Azure_DevOps`;
      setServiceNameFromRepository(repositoryName || projectName, projectName);

      if (!projectId) {
        setStatus('Project context was not provided by the branch action or hub.', true);
        sdk.notifyLoadFailed('Missing project context');
        return;
      }

      const hostUri = normalizeHostUri(hostedConfiguration.hostUri || context.collection?.uri || getHostBase());
      state.hostUri = hostUri;
      let accessToken = state.accessToken;
      let accessTokenError = null;

      if (!accessToken) {
        try {
          accessToken = await getAccessTokenWithRetry(sdk);
        } catch (tokenError) {
          console.error('Failed to acquire Azure DevOps access token', tokenError);
          accessTokenError = normalizeAccessTokenError(tokenError);
        }
      }

      state.accessTokenError = accessTokenError;

      if (!accessToken) {
        const errorMessage =
          accessTokenError ||
          'Failed to acquire access token from Azure DevOps. Reload the page and relaunch the generator from a branch action.';
        setReauthenticationVisibility(
          true,
          buildTokenRecoveryMessage(errorMessage)
        );
        setStatus(errorMessage, true);
        sdk.notifyLoadSucceeded?.();
        return;
      }

      try {
        state.accessToken = accessToken;
        setReauthenticationVisibility(false);
        if (!repositoryName && repoId) {
          try {
            const repoUrl = `${hostUri}${encodeURIComponent(projectId)}/_apis/git/repositories/${encodeURIComponent(
              repoId
            )}?api-version=6.0`;
            const repoRes = await fetch(repoUrl, { headers: authHeaders(accessToken) });
            if (repoRes.ok) {
              const repoPayload = await repoRes.json();
              repositoryName = repoPayload?.name || repositoryName;
              state.rawRepositoryName = repositoryName || state.rawRepositoryName;
              state.repositoryName = repositoryName;
              setServiceNameFromRepository(repositoryName, projectName);
            }
          } catch (repoError) {
            console.warn('Failed to fetch repository metadata', repoError);
          }
        }
        await loadDeploymentTargets({ hostUri, branch });
        const resourceLoads = [
          loadPools({ hostUri, projectId, accessToken }),
          loadProjectStacks({
            hostUri, projectId, projectName: state.rawProjectName || projectName, accessToken
          })
        ];
        if (!isMonorepoMode()) {
          resourceLoads.push(
            loadContainerRegistries({ hostUri, projectId, accessToken }),
            refreshDockerfiles({ hostUri, projectId, repoId, branch, accessToken })
          );
        }
        await Promise.all(resourceLoads);
        setStatus(
          isMonorepoMode()
            ? 'Azure DevOps context ready. Generate the MR Pipeline when you are ready.'
            : 'Azure DevOps context ready. Generate the pipeline when you are ready.'
        );
      } catch (tokenError) {
        console.error('Failed to initialize Azure DevOps context', tokenError);
        const detail = sanitizeErrorDetail(tokenError?.detail || tokenError?.message || '');
        let detailMessage;
        if (state.deploymentTargetsReady) {
          detailMessage =
            accessTokenError || 'Failed to initialize Azure DevOps resources. Reload the page and try again.';
        } else if (tokenError?.domain === 'komodo') {
          detailMessage = `Could not load active Komodo servers. ${
            detail || 'Verify the central credential file, Komodo API access, TLS certificate, and CORS origin.'
          }`;
        } else {
          detailMessage = `Could not load ${DEPLOYMENT_TARGETS_CONFIG.collection}/${DEPLOYMENT_TARGETS_CONFIG.project}/${DEPLOYMENT_TARGETS_CONFIG.repository}:${DEPLOYMENT_TARGETS_CONFIG.path}. ${
            detail || 'Verify the file, branch, YAML structure, and repository Read permission.'
          }`;
        }
        if (isUnauthorizedError(tokenError) && tokenError?.authenticationMode !== 'browser-session') {
          setReauthenticationVisibility(
            true,
            'The Azure DevOps host token could not access the required APIs. Sign out and authenticate again, then reopen the generator.'
          );
        }
        setStatus(detailMessage, true);
        sdk.notifyLoadSucceeded?.();
        return;
      }

      sdk.notifyLoadSucceeded();
    } catch (error) {
      console.error('Failed to initialize extension frame', error);
      const fallbackMessage = /Timed out waiting for Azure DevOps host/i.test(error?.message || '')
        ? 'Could not connect to the Azure DevOps host. If you opened this page directly, use the form to generate the YAML and copy it below.'
        : 'Failed to initialize extension frame. Check extension permissions and reload, or copy the template below.';
      if (state.projectId && state.hostUri) {
        setReauthenticationVisibility(
          true,
          `${fallbackMessage} Sign out and authenticate again below to rebuild the Azure DevOps host session.`
        );
      }
      setStatus(fallbackMessage, true);
      const sdk = normalizeSdk(window.VSS || window.parent?.VSS);
      sdk?.notifyLoadSucceeded?.();
    } finally {
      setSubmitting(false);
    }
  };

  const startInitialization = () => {
    if (initializationPromise) {
      return initializationPromise;
    }
    setStatus('Loading Azure DevOps context...');
    initializationPromise = init();
    return initializationPromise;
  };

  if (environmentSelect) {
    environmentSelect.addEventListener('change', (event) => {
      setKomodoServerFromEnvironment(event.target.value);
    });
  }

  reauthenticateButton?.addEventListener('click', () => {
    restartAzureDevOpsSession();
  });

  authorizeExtensionButton?.addEventListener('click', () => {
    openExtensionAuthorization();
  });

  window.addEventListener('message', (event) => {
    if (!event?.data || event.origin !== window.location.origin) return;
    if (event.data.type === 'pipeline-bootstrap') {
      applyBootstrapPayload(event.data.payload || {}, 'message');
      event.source?.postMessage({ type: 'pipeline-bootstrap-ack' }, event.origin);
    }
  });

  startInitialization();
})();
