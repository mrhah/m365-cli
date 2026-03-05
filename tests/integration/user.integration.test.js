import { describe, it, expect, beforeAll, afterAll } from 'vitest';
import { existsSync } from 'fs';
import { resolve } from 'path';
import { loadCreds, isTokenExpired, getAccessToken } from '../../src/auth/token-manager.js';
import graphClient from '../../src/graph/client.js';
import { searchUser } from '../../src/commands/user.js';

/**
 * Integration tests for user search feature.
 *
 * These tests call the real Microsoft Graph API and require:
 *   1. A dedicated integration credentials file (separate from your daily-use creds)
 *   2. Environment variables pointing to the integration app registration
 *
 * Setup:
 *   - Create an Azure AD app registration for integration testing
 *   - Authenticate with it and save the credentials file
 *   - Copy .env.integration.example → .env.integration and fill in your values
 *   - Set the env vars before running tests:
 *
 *       export M365_CLIENT_ID="<integration-app-client-id>"
 *       export M365_TENANT_ID="<integration-app-tenant-id>"
 *       export M365_CREDS_PATH="~/.m365-cli/integration-credentials.json"
 *
 *   Or source from the env file:
 *       source .env.integration && npx vitest run tests/user.integration.test.js
 *
 * Run:
 *   npx vitest run tests/user.integration.test.js
 */

// --- Dedicated integration test credentials via env vars ---
// These MUST be set externally (not hardcoded) to isolate integration tests
// from global/production credentials.
const INTEGRATION_CLIENT_ID = process.env.M365_INTEGRATION_CLIENT_ID;
const INTEGRATION_TENANT_ID = process.env.M365_INTEGRATION_TENANT_ID;
const INTEGRATION_CREDS_PATH = process.env.M365_INTEGRATION_CREDS_PATH;

let hasAuth = false;
let accessToken = null;

// Saved original env vars for restore
let savedEnv = {};

beforeAll(async () => {
  // Check if integration credentials are configured
  if (!INTEGRATION_CLIENT_ID || !INTEGRATION_TENANT_ID || !INTEGRATION_CREDS_PATH) {
    console.log(
      '⏭️  Integration env vars not set (M365_INTEGRATION_CLIENT_ID, M365_INTEGRATION_TENANT_ID, M365_INTEGRATION_CREDS_PATH) — skipping integration tests'
    );
    return;
  }

  // Resolve the creds path (expand ~ manually since env vars don't auto-expand)
  const resolvedCredsPath = INTEGRATION_CREDS_PATH.startsWith('~/')
    ? resolve(process.env.HOME, INTEGRATION_CREDS_PATH.slice(2))
    : resolve(INTEGRATION_CREDS_PATH);

  if (!existsSync(resolvedCredsPath)) {
    console.log(
      `⏭️  Integration credentials file not found at ${resolvedCredsPath} — skipping integration tests`
    );
    return;
  }

  // Save original env vars
  savedEnv = {
    M365_CLIENT_ID: process.env.M365_CLIENT_ID,
    M365_TENANT_ID: process.env.M365_TENANT_ID,
    M365_CREDS_PATH: process.env.M365_CREDS_PATH,
  };

  // Override with dedicated integration test values
  process.env.M365_CLIENT_ID = INTEGRATION_CLIENT_ID;
  process.env.M365_TENANT_ID = INTEGRATION_TENANT_ID;
  process.env.M365_CREDS_PATH = resolvedCredsPath;

  try {
    const creds = loadCreds();
    if (!creds || !creds.accessToken) {
      console.log('⏭️  No valid credentials in integration creds file — skipping integration tests');
      return;
    }

    // Try to get a valid (possibly refreshed) token
    accessToken = await getAccessToken();
    hasAuth = true;
  } catch (error) {
    console.log(`⏭️  Auth unavailable (${error.message}) — skipping integration tests`);
  }
});

afterAll(() => {
  // Restore original env vars
  for (const [key, value] of Object.entries(savedEnv)) {
    if (value === undefined) {
      delete process.env[key];
    } else {
      process.env[key] = value;
    }
  }
});

describe('[Integration] User search — Graph API', { timeout: 30000 }, () => {
  describe('Organization users (/users endpoint)', () => {
    it('should search organization users by name', { retry: 2 }, async (ctx) => {
      if (!hasAuth) return ctx.skip();

      const results = await graphClient.user.searchUsers('a', { top: 5 });

      expect(Array.isArray(results)).toBe(true);

      // Each result should have expected Graph API fields
      for (const user of results) {
        expect(user).toHaveProperty('id');
        expect(user).toHaveProperty('displayName');
        expect(user).toHaveProperty('userPrincipalName');
        // mail can be null for some users, but the field should exist
        expect('mail' in user).toBe(true);
      }
    });

    it('should respect the top parameter', async (ctx) => {
      if (!hasAuth) return ctx.skip();

      const results = await graphClient.user.searchUsers('a', { top: 2 });

      expect(Array.isArray(results)).toBe(true);
      expect(results.length).toBeLessThanOrEqual(2);
    });

    it('should return empty array for nonsense query', async (ctx) => {
      if (!hasAuth) return ctx.skip();

      const results = await graphClient.user.searchUsers('zzxqqnonexistent99', { top: 5 });

      expect(Array.isArray(results)).toBe(true);
      expect(results.length).toBe(0);
    });
  });

  describe('Personal contacts (/me/contacts endpoint)', () => {
    it('should search personal contacts by name', async (ctx) => {
      if (!hasAuth) return ctx.skip();

      // This may return empty if the user has no contacts — that's fine.
      // We're testing that the API call succeeds and returns an array.
      const results = await graphClient.user.searchContacts('a', { top: 5 });

      expect(Array.isArray(results)).toBe(true);

      for (const contact of results) {
        expect(contact).toHaveProperty('id');
        expect(contact).toHaveProperty('displayName');
        // emailAddresses is an array (possibly empty)
        expect(Array.isArray(contact.emailAddresses)).toBe(true);
      }
    });

    it('should return empty array for nonsense query', async (ctx) => {
      if (!hasAuth) return ctx.skip();

      const results = await graphClient.user.searchContacts('zzxqqnonexistent99', { top: 5 });

      expect(Array.isArray(results)).toBe(true);
      expect(results.length).toBe(0);
    });
  });

  describe('Full searchUser() command flow', () => {
    it('should execute the full search flow without throwing', async (ctx) => {
      if (!hasAuth) return ctx.skip();

      // searchUser calls both endpoints, normalizes, deduplicates, and outputs.
      // We just verify it completes without error.
      // The function calls console.log for output — we don't assert on that here.
      await expect(
        searchUser('a', { top: 5, json: true })
      ).resolves.not.toThrow();
    });

    it('should handle empty results gracefully', async (ctx) => {
      if (!hasAuth) return ctx.skip();

      await expect(
        searchUser('zzxqqnonexistent99', { top: 5, json: true })
      ).resolves.not.toThrow();
    });
  });
});
