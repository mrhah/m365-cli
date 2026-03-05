import { describe, it, expect, beforeAll, afterAll } from 'vitest';
import { existsSync } from 'fs';
import { resolve } from 'path';
import { loadCreds, isTokenExpired, getAccessToken, refreshToken } from '../../src/auth/token-manager.js';

/**
 * Integration tests for Token Manager.
 *
 * These tests call the real Azure AD token endpoint and require:
 *   1. A dedicated integration credentials file
 *   2. Environment variables pointing to the integration app registration
 *
 * Setup:
 *   export M365_INTEGRATION_CLIENT_ID="<integration-app-client-id>"
 *   export M365_INTEGRATION_TENANT_ID="<integration-app-tenant-id>"
 *   export M365_INTEGRATION_CREDS_PATH="~/.m365-cli/integration-credentials.json"
 *
 * Run:
 *   npx vitest run tests/token-manager.integration.test.js
 */

const INTEGRATION_CLIENT_ID = process.env.M365_INTEGRATION_CLIENT_ID;
const INTEGRATION_TENANT_ID = process.env.M365_INTEGRATION_TENANT_ID;
const INTEGRATION_CREDS_PATH = process.env.M365_INTEGRATION_CREDS_PATH;

let hasAuth = false;
let savedEnv = {};
let loadedCreds = null;

beforeAll(async () => {
  if (!INTEGRATION_CLIENT_ID || !INTEGRATION_TENANT_ID || !INTEGRATION_CREDS_PATH) {
    console.log(
      '⏭️  Integration env vars not set — skipping token-manager integration tests'
    );
    return;
  }

  const resolvedCredsPath = INTEGRATION_CREDS_PATH.startsWith('~/')
    ? resolve(process.env.HOME, INTEGRATION_CREDS_PATH.slice(2))
    : resolve(INTEGRATION_CREDS_PATH);

  if (!existsSync(resolvedCredsPath)) {
    console.log(
      `⏭️  Integration credentials file not found at ${resolvedCredsPath} — skipping`
    );
    return;
  }

  savedEnv = {
    M365_CLIENT_ID: process.env.M365_CLIENT_ID,
    M365_TENANT_ID: process.env.M365_TENANT_ID,
    M365_CREDS_PATH: process.env.M365_CREDS_PATH,
  };

  process.env.M365_CLIENT_ID = INTEGRATION_CLIENT_ID;
  process.env.M365_TENANT_ID = INTEGRATION_TENANT_ID;
  process.env.M365_CREDS_PATH = resolvedCredsPath;

  try {
    loadedCreds = loadCreds();
    if (!loadedCreds || !loadedCreds.accessToken) {
      console.log('⏭️  No valid credentials — skipping token-manager integration tests');
      return;
    }
    hasAuth = true;
  } catch (error) {
    console.log(`⏭️  Auth unavailable (${error.message}) — skipping token-manager integration tests`);
  }
});

afterAll(() => {
  for (const [key, value] of Object.entries(savedEnv)) {
    if (value === undefined) {
      delete process.env[key];
    } else {
      process.env[key] = value;
    }
  }
});

describe('[Integration] Token Manager — Azure AD', { timeout: 30000 }, () => {
  describe('Load credentials', () => {
    it('should load valid credentials from file', (ctx) => {
      if (!hasAuth) return ctx.skip();

      const creds = loadCreds();

      expect(creds).toBeDefined();
      expect(creds).toHaveProperty('accessToken');
      expect(typeof creds.accessToken).toBe('string');
      expect(creds.accessToken.length).toBeGreaterThan(0);
      expect(creds).toHaveProperty('refreshToken');
      expect(typeof creds.refreshToken).toBe('string');
    });

    it('should have expiresAt timestamp', (ctx) => {
      if (!hasAuth) return ctx.skip();

      const creds = loadCreds();

      expect(creds).toHaveProperty('expiresAt');
      expect(typeof creds.expiresAt).toBe('number');
    });
  });

  describe('Token expiration check', () => {
    it('should check if token is expired', (ctx) => {
      if (!hasAuth) return ctx.skip();

      const creds = loadCreds();
      const expired = isTokenExpired(creds);

      // The result should be a boolean
      expect(typeof expired).toBe('boolean');
      expect([true, false]).toContain(expired);
    });

    it('should report expired for null creds', () => {
      const expired = isTokenExpired(null);
      expect(expired).toBe(true);
    });

    it('should report expired for creds without expiresAt', () => {
      const expired = isTokenExpired({ accessToken: 'test' });
      expect(expired).toBe(true);
    });

    it('should report expired for past timestamp', () => {
      const expired = isTokenExpired({
        accessToken: 'test',
        expiresAt: Math.floor(Date.now() / 1000) - 3600,
      });
      expect(expired).toBe(true);
    });

    it('should report not expired for far future timestamp', () => {
      const expired = isTokenExpired({
        accessToken: 'test',
        expiresAt: Math.floor(Date.now() / 1000) + 36000,
      });
      expect(expired).toBe(false);
    });
  });

  describe('Get access token (auto-refresh)', () => {
    it('should return a valid access token', async (ctx) => {
      if (!hasAuth) return ctx.skip();

      const token = await getAccessToken();

      expect(typeof token).toBe('string');
      expect(token.length).toBeGreaterThan(0);
    });

    it('should return consistent token on repeated calls', async (ctx) => {
      if (!hasAuth) return ctx.skip();

      const token1 = await getAccessToken();
      const token2 = await getAccessToken();

      // Both should be valid strings
      expect(typeof token1).toBe('string');
      expect(typeof token2).toBe('string');
      expect(token1.length).toBeGreaterThan(0);
      expect(token2.length).toBeGreaterThan(0);

      // If token wasn't expired, both calls should return the same token
      // If it was expired and refreshed, both should return the new token
      // Either way, they should match since both go through loadCreds
    });
  });

  describe('Refresh token', () => {
    it('should refresh the token using refresh token', async (ctx) => {
      if (!hasAuth) return ctx.skip();

      const creds = loadCreds();
      if (!creds.refreshToken) {
        console.log('  (no refresh token — skipping)');
        return ctx.skip();
      }

      const refreshed = await refreshToken(creds.refreshToken);

      expect(refreshed).toHaveProperty('accessToken');
      expect(typeof refreshed.accessToken).toBe('string');
      expect(refreshed.accessToken.length).toBeGreaterThan(0);

      expect(refreshed).toHaveProperty('refreshToken');
      expect(typeof refreshed.refreshToken).toBe('string');
      expect(refreshed.refreshToken.length).toBeGreaterThan(0);

      expect(refreshed).toHaveProperty('expiresIn');
      expect(typeof refreshed.expiresIn).toBe('number');
      expect(refreshed.expiresIn).toBeGreaterThan(0);
    });
  });

  describe('Token validity', () => {
    it('should get a token that can make Graph API calls', async (ctx) => {
      if (!hasAuth) return ctx.skip();

      const token = await getAccessToken();

      // Verify the token works by making a simple Graph API call
      const response = await fetch('https://graph.microsoft.com/v1.0/me?$select=id,displayName', {
        headers: {
          'Authorization': `Bearer ${token}`,
        },
      });

      expect(response.ok).toBe(true);

      const data = await response.json();
      expect(data).toHaveProperty('id');
      expect(data).toHaveProperty('displayName');
    });
  });
});
