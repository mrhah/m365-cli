import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { readFileSync } from 'fs';
import { homedir } from 'os';
import { join } from 'path';

vi.mock('fs', async () => {
  const actual = await vi.importActual('fs');
  return {
    ...actual,
    readFileSync: vi.fn(actual.readFileSync),
  };
});

const credsPath = join(homedir(), '.m365-cli/credentials.json');

let config;
async function loadConfig() {
  const mod = await import('../../src/utils/config.js');
  config = mod.default;
  return mod;
}

describe('Cloud Config', () => {
  let savedEnv;

  beforeEach(async () => {
    savedEnv = { ...process.env };
    delete process.env.M365_CLOUD;
    delete process.env.M365_GRAPH_API_URL;
    delete process.env.M365_AUTH_URL;
    delete process.env.M365_SCOPE_PREFIX;
    vi.resetModules();
    readFileSync.mockRestore?.();
    const fs = await import('fs');
    const actualFs = await vi.importActual('fs');
    fs.readFileSync.mockImplementation(actualFs.readFileSync);
  });

  afterEach(() => {
    process.env = savedEnv;
  });

  describe('getActiveCloud — priority resolution', () => {
    it('should default to "global" when no env, no creds', async () => {
      const { getActiveCloud } = await loadConfig();
      expect(getActiveCloud()).toBe('global');
    });

    it('should return M365_CLOUD env var when set to "china"', async () => {
      process.env.M365_CLOUD = 'china';
      const { getActiveCloud } = await loadConfig();
      expect(getActiveCloud()).toBe('china');
    });

    it('should normalize M365_CLOUD to lowercase', async () => {
      process.env.M365_CLOUD = 'China';
      const { getActiveCloud } = await loadConfig();
      expect(getActiveCloud()).toBe('china');
    });

    it('should ignore invalid M365_CLOUD and fall back to default', async () => {
      process.env.M365_CLOUD = 'invalid-cloud';
      const fs = await import('fs');
      fs.readFileSync.mockImplementation((path, ...args) => {
        if (path === credsPath) {
          const err = new Error('ENOENT');
          err.code = 'ENOENT';
          throw err;
        }
        const actualFs = require('fs');
        throw new Error(`unexpected readFileSync call: ${path}`);
      });

      const { getActiveCloud } = await loadConfig();
      expect(getActiveCloud()).toBe('global');
    });

    it('should read cloud from saved credentials when no env var', async () => {
      const fs = await import('fs');
      const actualFs = await vi.importActual('fs');
      fs.readFileSync.mockImplementation((path, ...args) => {
        if (path === credsPath) {
          return JSON.stringify({ accessToken: 'tok', cloud: 'china' });
        }
        return actualFs.readFileSync(path, ...args);
      });

      const { getActiveCloud } = await loadConfig();
      expect(getActiveCloud()).toBe('china');
    });

    it('should prefer M365_CLOUD env over saved credentials', async () => {
      process.env.M365_CLOUD = 'global';
      const fs = await import('fs');
      const actualFs = await vi.importActual('fs');
      fs.readFileSync.mockImplementation((path, ...args) => {
        if (path === credsPath) {
          return JSON.stringify({ accessToken: 'tok', cloud: 'china' });
        }
        return actualFs.readFileSync(path, ...args);
      });

      const { getActiveCloud } = await loadConfig();
      expect(getActiveCloud()).toBe('global');
    });
  });

  describe('getCloudConfig', () => {
    it('should return global endpoints by default', async () => {
      const { getCloudConfig } = await loadConfig();
      const cc = getCloudConfig('global');
      expect(cc.graphApiUrl).toBe('https://graph.microsoft.com/v1.0');
      expect(cc.authUrl).toBe('https://login.microsoftonline.com');
      expect(cc.scopePrefix).toBe('https://graph.microsoft.com/');
    });

    it('should return china endpoints for china cloud', async () => {
      const { getCloudConfig } = await loadConfig();
      const cc = getCloudConfig('china');
      expect(cc.graphApiUrl).toBe('https://microsoftgraph.chinacloudapi.cn/v1.0');
      expect(cc.authUrl).toBe('https://login.chinacloudapi.cn');
      expect(cc.scopePrefix).toBe('https://microsoftgraph.chinacloudapi.cn/');
      expect(cc.deviceLoginUrl).toBe('https://login.chinacloudapi.cn/common/oauth2/deviceauth');
    });

    it('should fall back to global for unknown cloud name', async () => {
      const { getCloudConfig } = await loadConfig();
      const cc = getCloudConfig('nonexistent');
      expect(cc.graphApiUrl).toBe('https://graph.microsoft.com/v1.0');
    });
  });

  describe('applyScopes', () => {
    it('should prepend scope prefix to bare scope names', async () => {
      const { applyScopes } = await loadConfig();
      const result = applyScopes(['Mail.Read', 'Files.ReadWrite'], 'https://graph.microsoft.com/');
      expect(result).toEqual([
        'https://graph.microsoft.com/Mail.Read',
        'https://graph.microsoft.com/Files.ReadWrite',
      ]);
    });

    it('should use china prefix for china cloud', async () => {
      const { applyScopes } = await loadConfig();
      const prefix = 'https://microsoftgraph.chinacloudapi.cn/';
      const result = applyScopes(['Mail.Read'], prefix);
      expect(result).toEqual(['https://microsoftgraph.chinacloudapi.cn/Mail.Read']);
    });

    it('should leave offline_access unchanged', async () => {
      const { applyScopes } = await loadConfig();
      const result = applyScopes(['offline_access', 'Mail.Read'], 'https://graph.microsoft.com/');
      expect(result[0]).toBe('offline_access');
      expect(result[1]).toBe('https://graph.microsoft.com/Mail.Read');
    });

    it('should leave already-prefixed scopes unchanged', async () => {
      const { applyScopes } = await loadConfig();
      const result = applyScopes(
        ['https://graph.microsoft.com/Mail.Read'],
        'https://microsoftgraph.chinacloudapi.cn/'
      );
      expect(result[0]).toBe('https://graph.microsoft.com/Mail.Read');
    });

    it('should return non-array input as-is', async () => {
      const { applyScopes } = await loadConfig();
      expect(applyScopes(null, 'prefix')).toBeNull();
      expect(applyScopes(undefined, 'prefix')).toBeUndefined();
    });
  });

  describe('getConfig — cloud-aware resolution', () => {
    it('should resolve graphApiUrl from active cloud', async () => {
      process.env.M365_CLOUD = 'china';
      const { getConfig } = await loadConfig();
      expect(getConfig('graphApiUrl')).toBe('https://microsoftgraph.chinacloudapi.cn/v1.0');
    });

    it('should resolve authUrl from active cloud', async () => {
      process.env.M365_CLOUD = 'china';
      const { getConfig } = await loadConfig();
      expect(getConfig('authUrl')).toBe('https://login.chinacloudapi.cn');
    });

    it('should resolve scopePrefix from active cloud', async () => {
      process.env.M365_CLOUD = 'china';
      const { getConfig } = await loadConfig();
      expect(getConfig('scopePrefix')).toBe('https://microsoftgraph.chinacloudapi.cn/');
    });

    it('should resolve deviceLoginUrl from active cloud', async () => {
      process.env.M365_CLOUD = 'china';
      const { getConfig } = await loadConfig();
      expect(getConfig('deviceLoginUrl')).toBe('https://login.chinacloudapi.cn/common/oauth2/deviceauth');
    });

    it('should return global endpoints when cloud is global', async () => {
      process.env.M365_CLOUD = 'global';
      const { getConfig } = await loadConfig();
      expect(getConfig('graphApiUrl')).toBe('https://graph.microsoft.com/v1.0');
      expect(getConfig('authUrl')).toBe('https://login.microsoftonline.com');
    });

    it('should apply china scope prefix to workScopes', async () => {
      process.env.M365_CLOUD = 'china';
      const { getConfig } = await loadConfig();
      const scopes = getConfig('workScopes');
      const chinaPrefix = 'https://microsoftgraph.chinacloudapi.cn/';
      for (const scope of scopes) {
        if (scope !== 'offline_access') {
          expect(scope.startsWith(chinaPrefix)).toBe(true);
        }
      }
    });

    it('should apply global scope prefix to personalScopes by default', async () => {
      const { getConfig } = await loadConfig();
      const scopes = getConfig('personalScopes');
      const globalPrefix = 'https://graph.microsoft.com/';
      for (const scope of scopes) {
        if (scope !== 'offline_access') {
          expect(scope.startsWith(globalPrefix)).toBe(true);
        }
      }
    });

    it('should allow env var override for cloud keys', async () => {
      process.env.M365_GRAPH_API_URL = 'https://custom.graph.api/v1.0';
      const { getConfig } = await loadConfig();
      expect(getConfig('graphApiUrl')).toBe('https://custom.graph.api/v1.0');
    });

    it('should return non-cloud keys from default config', async () => {
      const { getConfig } = await loadConfig();
      expect(getConfig('tokenRefreshBuffer')).toBe(300);
    });
  });

  describe('login guardrails', () => {
    it('should reject --cloud china --account-type personal', async () => {
      vi.resetModules();
      const fs = await import('fs');
      const actualFs = await vi.importActual('fs');
      fs.readFileSync.mockImplementation(actualFs.readFileSync);

      vi.mock('../../src/utils/config.js', async () => {
        const actual = await vi.importActual('../../src/utils/config.js');
        return { ...actual, default: actual.default || actual };
      });

      vi.mock('../../src/auth/device-flow.js', () => ({
        deviceCodeFlow: vi.fn(),
      }));

      process.env.M365_CLOUD = 'china';

      const { login } = await import('../../src/auth/token-manager.js');

      await expect(login({ accountType: 'personal', cloud: 'china' }))
        .rejects.toThrow('21Vianet (China) does not support personal Microsoft accounts');
    });
  });
});
