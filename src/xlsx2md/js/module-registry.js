/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    var _a;
    var _b;
    const registry = ((_a = (_b = globalThis).__xlsx2mdModuleRegistryStore) !== null && _a !== void 0 ? _a : (_b.__xlsx2mdModuleRegistryStore = {}));
    function getModule(name) {
        return registry[name];
    }
    function requireModule(name, errorMessage) {
        const moduleValue = getModule(name);
        if (!moduleValue) {
            throw new Error(errorMessage);
        }
        return moduleValue;
    }
    function registerModule(name, moduleValue) {
        registry[name] = moduleValue;
        return moduleValue;
    }
    globalThis.__xlsx2mdModuleRegistry = {
        getModule,
        requireModule,
        registerModule
    };
})();
