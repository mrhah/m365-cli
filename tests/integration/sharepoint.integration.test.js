import { describe, it, expect, beforeAll, afterAll } from 'vitest';
import { existsSync} from 'fs';
import { resolve } from 'path';
import { loadCreds, getAccessToken } from '../../src/auth/token-manager.js';
import graphClient from '../../src/graph/client.js';
import sharepointCommands from '../../src/commands/sharepoint.js';

/**
 * Integration tests for SharePoint feature.
 *
 * These tests call the real Microsoft Graph API and require:
 *   1. A dedicated integration credentials file
 *   2. Environment variables pointing to the integration app registration
 *   3. Optionally, M365_INTEGRATION_SP_SITE env var with a SharePoint site URL
 *      (e.g., "contoso.sharepoint.com:/sites/teamsite")
 *
 * Setup:
 *   export M365_INTEGRATION_CLIENT_ID="<integration-app-client-id>"
 *   export M365_INTEGRATION_TENANT_ID="<integration-app-tenant-id>"
 *   export M365_INTEGRATION_CREDS_PATH="~/.m365-cli/integration-credentials.json"
 *   export M365_INTEGRATION_SP_SITE="contoso.sharepoint.com:/sites/teamsite"  # optional
 *
 * Run:
 *   npx vitest run tests/sharepoint.integration.test.js
 */

const INTEGRATION_CLIENT_ID = process.env.M365_INTEGRATION_CLIENT_ID;
const INTEGRATION_TENANT_ID = process.env.M365_INTEGRATION_TENANT_ID;
const INTEGRATION_CREDS_PATH = process.env.M365_INTEGRATION_CREDS_PATH;
const INTEGRATION_SP_SITE = process.env.M365_INTEGRATION_SP_SITE;

let hasAuth = false;
let savedEnv = {};

// Resolved site ID for tests that need a specific site
let resolvedSiteId = null;

beforeAll(async () => {
  if (!INTEGRATION_CLIENT_ID || !INTEGRATION_TENANT_ID || !INTEGRATION_CREDS_PATH) {
    console.log(
      '⏭️  Integration env vars not set — skipping SharePoint integration tests'
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
    const creds = loadCreds();
    if (!creds || !creds.accessToken) {
      console.log('⏭️  No valid credentials — skipping SharePoint integration tests');
      return;
    }
    await getAccessToken();
    hasAuth = true;

    // Try to resolve the SharePoint site if provided
    if (INTEGRATION_SP_SITE) {
      try {
        resolvedSiteId = await graphClient.sharepoint._parseSite(INTEGRATION_SP_SITE);
      } catch (error) {
        console.log(`⚠️  Could not resolve SP site (${error.message}) — site-specific tests will skip`);
      }
    }
  } catch (error) {
    console.log(`⏭️  Auth unavailable (${error.message}) — skipping SharePoint integration tests`);
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

describe('[Integration] SharePoint — Graph API', { timeout: 30000 }, () => {
  describe('List sites', () => {
    it('should list followed sites', { retry: 2 }, async (ctx) => {
      if (!hasAuth) return ctx.skip();

      try {
        const sites = await graphClient.sharepoint.sites({ top: 5 });

        expect(Array.isArray(sites)).toBe(true);

        for (const site of sites) {
          expect(site).toHaveProperty('id');
          expect(site).toHaveProperty('name');
          expect(site).toHaveProperty('webUrl');
        }
      } catch (error) {
        // SharePoint may not be available for all accounts
        if (error.statusCode === 403 || error.statusCode === 401) {
          console.log('  (SharePoint access denied — skipping)');
          return ctx.skip();
        }
        // Network errors (fetch failed) on first call — retry handles this
        throw error;
      }
    });

    it('should search sites by keyword', { retry: 2 }, async (ctx) => {
      if (!hasAuth) return ctx.skip();

      try {
        const sites = await graphClient.sharepoint.sites({ search: 'team', top: 5 });

        expect(Array.isArray(sites)).toBe(true);
        // May return empty if no sites match
      } catch (error) {
        if (error.statusCode === 403 || error.statusCode === 401) {
          console.log('  (SharePoint access denied — skipping)');
          return ctx.skip();
        }
        throw error;
      }
    });

    it('should respect top parameter', async (ctx) => {
      if (!hasAuth) return ctx.skip();

      try {
        const sites = await graphClient.sharepoint.sites({ top: 2 });

        expect(Array.isArray(sites)).toBe(true);
        expect(sites.length).toBeLessThanOrEqual(2);
      } catch (error) {
        if (error.statusCode === 403 || error.statusCode === 401) {
          return ctx.skip();
        }
        throw error;
      }
    });
  });

  describe('Parse site URL', () => {
    it('should parse a Graph API ID (hostname,guid,guid) format', async (ctx) => {
      if (!hasAuth) return ctx.skip();
      if (!resolvedSiteId) {
        console.log('  (no SP site configured — skipping)');
        return ctx.skip();
      }

      // The resolved site ID should be a parseable Graph ID (hostname,guid,guid)
      const parsed = await graphClient.sharepoint._parseSite(resolvedSiteId);
      expect(typeof parsed).toBe('string');
      expect(parsed.length).toBeGreaterThan(0);
    });

    it('should parse hostname:/path format', async (ctx) => {
      if (!hasAuth) return ctx.skip();
      if (!INTEGRATION_SP_SITE || !INTEGRATION_SP_SITE.includes(':/')) {
        console.log('  (no hostname:/path site configured — skipping)');
        return ctx.skip();
      }

      try {
        const siteId = await graphClient.sharepoint._parseSite(INTEGRATION_SP_SITE);
        expect(typeof siteId).toBe('string');
        expect(siteId.length).toBeGreaterThan(0);
      } catch (error) {
        if (error.statusCode === 403 || error.statusCode === 404) {
          console.log(`  (site not accessible: ${error.message} — skipping)`);
          return ctx.skip();
        }
        throw error;
      }
    });

    it('should throw for invalid site', async (ctx) => {
      if (!hasAuth) return ctx.skip();

      await expect(
        graphClient.sharepoint._parseSite('invalid-nonexistent-host.example.com:/sites/fakefake')
      ).rejects.toThrow();
    });
  });

  describe('List site lists', () => {
    it('should list site lists and document libraries', async (ctx) => {
      if (!hasAuth || !resolvedSiteId) {
        return ctx.skip();
      }

      try {
        const lists = await graphClient.sharepoint.lists(resolvedSiteId, { top: 10 });

        expect(Array.isArray(lists)).toBe(true);

        for (const list of lists) {
          expect(list).toHaveProperty('id');
          expect(list).toHaveProperty('displayName');
        }
      } catch (error) {
        if (error.statusCode === 403 || error.statusCode === 401) {
          console.log('  (access denied — skipping)');
          return ctx.skip();
        }
        throw error;
      }
    });

    it('should respect top parameter for lists', async (ctx) => {
      if (!hasAuth || !resolvedSiteId) {
        return ctx.skip();
      }

      try {
        const lists = await graphClient.sharepoint.lists(resolvedSiteId, { top: 2 });

        expect(Array.isArray(lists)).toBe(true);
        expect(lists.length).toBeLessThanOrEqual(2);
      } catch (error) {
        if (error.statusCode === 403 || error.statusCode === 401) {
          return ctx.skip();
        }
        throw error;
      }
    });
  });

  describe('List site files', () => {
    it('should list files in site default document library root', async (ctx) => {
      if (!hasAuth || !resolvedSiteId) {
        return ctx.skip();
      }

      try {
        const files = await graphClient.sharepoint.files(resolvedSiteId, '', { top: 10 });

        expect(Array.isArray(files)).toBe(true);

        for (const file of files) {
          expect(file).toHaveProperty('id');
          expect(file).toHaveProperty('name');
        }
      } catch (error) {
        if (error.statusCode === 403 || error.statusCode === 401) {
          console.log('  (access denied — skipping)');
          return ctx.skip();
        }
        throw error;
      }
    });
  });

  describe('List items in a list', () => {
    it('should list items in a SharePoint list', async (ctx) => {
      if (!hasAuth || !resolvedSiteId) {
        return ctx.skip();
      }

      try {
        // First, get a list to query
        const lists = await graphClient.sharepoint.lists(resolvedSiteId, { top: 5 });
        if (lists.length === 0) {
          console.log('  (no lists found — skipping)');
          return ctx.skip();
        }

        const items = await graphClient.sharepoint.items(resolvedSiteId, lists[0].id, { top: 5 });

        expect(Array.isArray(items)).toBe(true);
        // Items may be empty depending on the list
      } catch (error) {
        if (error.statusCode === 403 || error.statusCode === 401) {
          console.log('  (access denied — skipping)');
          return ctx.skip();
        }
        throw error;
      }
    });
  });

  describe('Search SharePoint content (/search/query)', () => {
    it('should search SharePoint content', async (ctx) => {
      if (!hasAuth) return ctx.skip();

      try {
        const results = await graphClient.sharepoint.search('report', { top: 5 });

        expect(Array.isArray(results)).toBe(true);
        // Results may be empty
      } catch (error) {
        if (error.statusCode === 403 || error.statusCode === 401) {
          console.log('  (SharePoint search access denied — skipping)');
          return ctx.skip();
        }
        throw error;
      }
    });

    it('should respect top parameter in search', async (ctx) => {
      if (!hasAuth) return ctx.skip();

      try {
        const results = await graphClient.sharepoint.search('test', { top: 2 });

        expect(Array.isArray(results)).toBe(true);
        expect(results.length).toBeLessThanOrEqual(2);
      } catch (error) {
        if (error.statusCode === 403 || error.statusCode === 401) {
          return ctx.skip();
        }
        throw error;
      }
    });
  });

  describe('Full command flows', () => {
    it('should execute listSites command without throwing', async (ctx) => {
      if (!hasAuth) return ctx.skip();

      try {
        await sharepointCommands.sites({ top: 5, json: true });
      } catch (error) {
        // Skip if insufficient permissions (Sites.Read.All required)
        if (error.message?.includes('process.exit') || error.statusCode === 403) {
          return ctx.skip();
        }
        throw error;
      }
    });

    it('should execute searchSites command without throwing', async (ctx) => {
      if (!hasAuth) return ctx.skip();

      try {
        await sharepointCommands.sites({ search: 'team', top: 5, json: true });
      } catch (error) {
        if (error.message?.includes('process.exit') || error.statusCode === 403) {
          return ctx.skip();
        }
        throw error;
      }
    });

    it('should execute listLists command without throwing', async (ctx) => {
      if (!hasAuth || !INTEGRATION_SP_SITE) return ctx.skip();

      await expect(
        sharepointCommands.lists(INTEGRATION_SP_SITE, { top: 5, json: true })
      ).resolves.not.toThrow();
    });

    it('should execute listFiles command without throwing', async (ctx) => {
      if (!hasAuth || !INTEGRATION_SP_SITE) return ctx.skip();

      await expect(
        sharepointCommands.files(INTEGRATION_SP_SITE, '', { top: 5, json: true })
      ).resolves.not.toThrow();
    });

    it('should execute searchContent command without throwing', async (ctx) => {
      if (!hasAuth) return ctx.skip();

      try {
        await sharepointCommands.search('report', { top: 5, json: true });
      } catch (error) {
        // Skip if insufficient permissions (Sites.Read.All required for search)
        if (error.message?.includes('process.exit') || error.statusCode === 403) {
          return ctx.skip();
        }
        throw error;
      }
    });

    it('should handle empty search results gracefully', async (ctx) => {
      if (!hasAuth) return ctx.skip();

      try {
        await sharepointCommands.search('zzxqqnonexistent99integration', { top: 5, json: true });
      } catch (error) {
        if (error.message?.includes('process.exit') || error.statusCode === 403) {
          return ctx.skip();
        }
        throw error;
      }
    });
  });
});
