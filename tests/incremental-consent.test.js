import { describe, it, expect, vi, beforeEach } from 'vitest';

// ──────────────────────────────────────────────────────────
// Tests for Incremental Consent Logic
//
// Covers:
//   1. getExtraScopes config helper
//   2. loginWithScopes flow in token-manager
//   3. GraphClient._detectFeature endpoint mapping
//   4. GraphClient.request retry-with-consent on 403
//
// NOTE: parseGraphError + InsufficientPrivilegesError tests
//       are in insufficient-privileges.test.js (unmocked error module)
// ──────────────────────────────────────────────────────────

// ============================================================
// Section 2: getExtraScopes config helper
// ============================================================

describe('getExtraScopes', () => {
  let getExtraScopes;

  beforeEach(async () => {
    const mod = await import('../src/utils/config.js');
    getExtraScopes = mod.getExtraScopes;
  });

  it('should return SharePoint scopes for "sharepoint" feature', () => {
    const scopes = getExtraScopes('sharepoint');
    expect(scopes).toEqual(['https://graph.microsoft.com/Sites.ReadWrite.All']);
  });

  it('should return empty array for unknown feature', () => {
    const scopes = getExtraScopes('nonexistent');
    expect(scopes).toEqual([]);
  });

  it('should return empty array for undefined feature', () => {
    const scopes = getExtraScopes(undefined);
    expect(scopes).toEqual([]);
  });
});

// ============================================================
// Section 3: loginWithScopes
// ============================================================

describe('loginWithScopes', () => {
  // Mock console to suppress output
  const mockConsoleLog = vi.fn();
  vi.stubGlobal('console', {
    ...console,
    log: mockConsoleLog,
    error: vi.fn(),
  });

  // Mock config
  vi.mock('../src/utils/config.js', () => ({
    default: {
      get: vi.fn((key) => {
        const config = {
          tenantId: 'common',
          clientId: 'test-client-id',
          credsPath: '~/.m365-cli/credentials.json',
          tokenRefreshBuffer: 60,
          scopes: ['https://graph.microsoft.com/Mail.ReadWrite', 'offline_access'],
        };
        return config[key];
      }),
      getCredsPath: vi.fn(() => '/tmp/test-creds.json'),
    },
    getExtraScopes: vi.fn((feature) => {
      if (feature === 'sharepoint') return ['https://graph.microsoft.com/Sites.ReadWrite.All'];
      return [];
    }),
  }));

  // Mock error
  vi.mock('../src/utils/error.js', () => {
    class ApiError extends Error {
      constructor(message, statusCode, details) {
        super(message);
        this.name = 'ApiError';
        this.statusCode = statusCode;
        this.code = 'API_ERROR';
        this.details = details;
      }
    }
    class InsufficientPrivilegesError extends ApiError {
      constructor(message, details) {
        super(message || 'Insufficient privileges', 403, details);
        this.name = 'InsufficientPrivilegesError';
        this.code = 'INSUFFICIENT_PRIVILEGES';
      }
    }
    return {
      AuthError: class AuthError extends Error {
        constructor(message, details) {
          super(message);
          this.name = 'AuthError';
          this.details = details;
        }
      },
      TokenExpiredError: class TokenExpiredError extends Error {
        constructor() {
          super('Token expired');
          this.name = 'TokenExpiredError';
        }
      },
      ApiError,
      InsufficientPrivilegesError,
      parseGraphError: vi.fn(),
    };
  });

  // Mock fs module
  vi.mock('fs', async () => {
    const actual = await vi.importActual('fs');
    return {
      ...actual,
      readFileSync: vi.fn(),
      writeFileSync: vi.fn(),
      mkdirSync: vi.fn(),
      chmodSync: vi.fn(),
      unlinkSync: vi.fn(),
    };
  });

  // Mock device-flow
  vi.mock('../src/auth/device-flow.js', () => ({
    deviceCodeFlow: vi.fn(),
  }));

  let loginWithScopes, loadCreds, deviceCodeFlow;
  let readFileSync, writeFileSync;

  beforeEach(async () => {
    vi.clearAllMocks();

    const tokenMod = await import('../src/auth/token-manager.js');
    loginWithScopes = tokenMod.loginWithScopes;
    loadCreds = tokenMod.loadCreds;

    const flowMod = await import('../src/auth/device-flow.js');
    deviceCodeFlow = flowMod.deviceCodeFlow;

    const fsMod = await import('fs');
    readFileSync = fsMod.readFileSync;
    writeFileSync = fsMod.writeFileSync;
  });

  it('should call deviceCodeFlow with additional scopes', async () => {
    const additionalScopes = ['https://graph.microsoft.com/Sites.ReadWrite.All'];

    // Existing creds
    readFileSync.mockReturnValue(JSON.stringify({
      accessToken: 'old-token',
      refreshToken: 'old-refresh',
      expiresAt: Math.floor(Date.now() / 1000) + 3600,
      grantedScopes: ['https://graph.microsoft.com/Mail.ReadWrite', 'offline_access'],
    }));

    deviceCodeFlow.mockResolvedValue({
      accessToken: 'new-token',
      refreshToken: 'new-refresh',
      expiresIn: 3600,
    });

    const token = await loginWithScopes(additionalScopes);

    expect(deviceCodeFlow).toHaveBeenCalledWith(additionalScopes);
    expect(token).toBe('new-token');
  });

  it('should save merged scopes after successful consent', async () => {
    const additionalScopes = ['https://graph.microsoft.com/Sites.ReadWrite.All'];

    readFileSync.mockReturnValue(JSON.stringify({
      accessToken: 'old-token',
      refreshToken: 'old-refresh',
      expiresAt: Math.floor(Date.now() / 1000) + 3600,
      grantedScopes: ['https://graph.microsoft.com/Mail.ReadWrite', 'offline_access'],
    }));

    deviceCodeFlow.mockResolvedValue({
      accessToken: 'new-token',
      refreshToken: 'new-refresh',
      expiresIn: 3600,
    });

    await loginWithScopes(additionalScopes);

    // Verify saveCreds was called with merged scopes
    expect(writeFileSync).toHaveBeenCalled();
    const savedCreds = JSON.parse(writeFileSync.mock.calls[0][1]);
    expect(savedCreds.grantedScopes).toContain('https://graph.microsoft.com/Mail.ReadWrite');
    expect(savedCreds.grantedScopes).toContain('offline_access');
    expect(savedCreds.grantedScopes).toContain('https://graph.microsoft.com/Sites.ReadWrite.All');
  });

  it('should deduplicate scopes when merging', async () => {
    const additionalScopes = [
      'https://graph.microsoft.com/Sites.ReadWrite.All',
      'https://graph.microsoft.com/Mail.ReadWrite', // already granted
    ];

    readFileSync.mockReturnValue(JSON.stringify({
      accessToken: 'old-token',
      refreshToken: 'old-refresh',
      expiresAt: Math.floor(Date.now() / 1000) + 3600,
      grantedScopes: ['https://graph.microsoft.com/Mail.ReadWrite', 'offline_access'],
    }));

    deviceCodeFlow.mockResolvedValue({
      accessToken: 'new-token',
      refreshToken: 'new-refresh',
      expiresIn: 3600,
    });

    await loginWithScopes(additionalScopes);

    const savedCreds = JSON.parse(writeFileSync.mock.calls[0][1]);
    // Mail.ReadWrite should appear only once
    const mailCount = savedCreds.grantedScopes.filter(
      s => s === 'https://graph.microsoft.com/Mail.ReadWrite'
    ).length;
    expect(mailCount).toBe(1);
  });
});

// ============================================================
// Section 4: GraphClient._detectFeature
// ============================================================

describe('GraphClient._detectFeature', () => {
  let graphClient;

  beforeEach(async () => {
    const mod = await import('../src/graph/client.js');
    graphClient = mod.default;
  });

  it('should return "sharepoint" for /sites endpoints', () => {
    expect(graphClient._detectFeature('/sites/abc')).toBe('sharepoint');
    expect(graphClient._detectFeature('/sites/abc/lists')).toBe('sharepoint');
    expect(graphClient._detectFeature('/sites/host,guid,guid/drive/root/children')).toBe('sharepoint');
  });

  it('should return "sharepoint" for /search/query endpoint', () => {
    expect(graphClient._detectFeature('/search/query')).toBe('sharepoint');
  });

  it('should return null for non-SharePoint endpoints', () => {
    expect(graphClient._detectFeature('/me/messages')).toBeNull();
    expect(graphClient._detectFeature('/me/events')).toBeNull();
    expect(graphClient._detectFeature('/me/drive/root/children')).toBeNull();
    expect(graphClient._detectFeature('/users')).toBeNull();
  });
});

// ============================================================
// Section 5: GraphClient.request — retry with consent on 403
// ============================================================

describe('GraphClient.request — incremental consent retry', () => {
  let graphClient, loginWithScopes, InsufficientPrivilegesError;

  // We need fetch to be mockable
  const mockFetch = vi.fn();
  vi.stubGlobal('fetch', mockFetch);

  beforeEach(async () => {
    vi.clearAllMocks();

    const errorMod = await import('../src/utils/error.js');
    InsufficientPrivilegesError = errorMod.InsufficientPrivilegesError;

    const tokenMod = await import('../src/auth/token-manager.js');
    loginWithScopes = tokenMod.loginWithScopes;

    const clientMod = await import('../src/graph/client.js');
    graphClient = clientMod.default;
  });

  it('should retry request after successful incremental consent on SharePoint 403', async () => {
    const { getAccessToken } = await import('../src/auth/token-manager.js');
    const { parseGraphError } = await import('../src/utils/error.js');
    const { deviceCodeFlow } = await import('../src/auth/device-flow.js');
    const { readFileSync } = await import('fs');

    // First call: getAccessToken returns old token
    readFileSync.mockReturnValue(JSON.stringify({
      accessToken: 'old-token',
      refreshToken: 'refresh-token',
      expiresAt: Math.floor(Date.now() / 1000) + 3600,
      grantedScopes: ['https://graph.microsoft.com/Mail.ReadWrite'],
    }));

    // Make parseGraphError return an InsufficientPrivilegesError
    parseGraphError.mockReturnValue(
      new InsufficientPrivilegesError('Insufficient privileges')
    );

    // deviceCodeFlow for loginWithScopes
    deviceCodeFlow.mockResolvedValue({
      accessToken: 'new-token',
      refreshToken: 'new-refresh',
      expiresIn: 3600,
    });

    // First request: 403 error
    mockFetch.mockResolvedValueOnce({
      ok: false,
      status: 403,
      json: async () => ({
        error: {
          code: 'Authorization_RequestDenied',
          message: 'Insufficient privileges',
        },
      }),
    });

    // Second request (after consent): success
    mockFetch.mockResolvedValueOnce({
      ok: true,
      status: 200,
      json: async () => ({ value: [{ id: 'site1', name: 'Team Site' }] }),
    });

    const result = await graphClient.request('/sites/abc/lists');

    // Should have been called twice: first fails with 403, retry succeeds
    expect(mockFetch).toHaveBeenCalledTimes(2);
    expect(result.value).toHaveLength(1);
    expect(result.value[0].name).toBe('Team Site');
  });

  it('should not retry more than once (prevent infinite loops)', async () => {
    const { parseGraphError } = await import('../src/utils/error.js');
    const { deviceCodeFlow } = await import('../src/auth/device-flow.js');
    const { readFileSync } = await import('fs');

    readFileSync.mockReturnValue(JSON.stringify({
      accessToken: 'old-token',
      refreshToken: 'refresh-token',
      expiresAt: Math.floor(Date.now() / 1000) + 3600,
      grantedScopes: ['https://graph.microsoft.com/Mail.ReadWrite'],
    }));

    parseGraphError.mockReturnValue(
      new InsufficientPrivilegesError('Insufficient privileges')
    );

    deviceCodeFlow.mockResolvedValue({
      accessToken: 'new-token',
      refreshToken: 'new-refresh',
      expiresIn: 3600,
    });

    // Both requests return 403
    mockFetch.mockResolvedValue({
      ok: false,
      status: 403,
      json: async () => ({
        error: {
          code: 'Authorization_RequestDenied',
          message: 'Still insufficient privileges',
        },
      }),
    });

    await expect(graphClient.request('/sites/abc/lists'))
      .rejects.toThrow();

    // Should be called exactly twice: first attempt + one retry
    expect(mockFetch).toHaveBeenCalledTimes(2);
  });

  it('should not trigger consent for non-SharePoint 403 errors', async () => {
    const { parseGraphError } = await import('../src/utils/error.js');
    const { readFileSync } = await import('fs');

    readFileSync.mockReturnValue(JSON.stringify({
      accessToken: 'old-token',
      refreshToken: 'refresh-token',
      expiresAt: Math.floor(Date.now() / 1000) + 3600,
    }));

    parseGraphError.mockReturnValue(
      new InsufficientPrivilegesError('Insufficient privileges')
    );

    mockFetch.mockResolvedValueOnce({
      ok: false,
      status: 403,
      json: async () => ({
        error: {
          code: 'Authorization_RequestDenied',
          message: 'Forbidden',
        },
      }),
    });

    // /me/messages is not a SharePoint endpoint — _detectFeature returns null
    await expect(graphClient.request('/me/messages'))
      .rejects.toThrow();

    // Only one fetch call — no retry
    expect(mockFetch).toHaveBeenCalledTimes(1);
  });
});
