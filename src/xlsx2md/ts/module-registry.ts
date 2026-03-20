(() => {
  type RegistryRecord = Record<string, unknown>;

  const registry = ((globalThis as typeof globalThis & {
    __xlsx2mdModuleRegistryStore?: RegistryRecord;
  }).__xlsx2mdModuleRegistryStore ??= {});

  function getModule<T>(name: string): T | undefined {
    return registry[name] as T | undefined;
  }

  function requireModule<T>(name: string, errorMessage: string): T {
    const moduleValue = getModule<T>(name);
    if (!moduleValue) {
      throw new Error(errorMessage);
    }
    return moduleValue;
  }

  function registerModule<T>(name: string, moduleValue: T): T {
    registry[name] = moduleValue as unknown;
    return moduleValue;
  }

  (globalThis as typeof globalThis & {
    __xlsx2mdModuleRegistry?: {
      getModule: typeof getModule;
      requireModule: typeof requireModule;
      registerModule: typeof registerModule;
    };
  }).__xlsx2mdModuleRegistry = {
    getModule,
    requireModule,
    registerModule
  };
})();
