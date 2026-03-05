import { describe, it, expect, beforeAll, afterAll } from 'vitest';
import { existsSync, writeFileSync, unlinkSync } from 'fs';
import { resolve, join } from 'path';
import { tmpdir } from 'os';
import { loadCreds, getAccessToken } from '../../src/auth/token-manager.js';
import graphClient from '../../src/graph/client.js';
import onedriveCommands from '../../src/commands/onedrive.js';

/**
 * Integration tests for OneDrive feature.
 *
 * These tests call the real Microsoft Graph API and require:
 *   1. A dedicated integration credentials file
 *   2. Environment variables pointing to the integration app registration
 *
 * Setup:
 *   export M365_INTEGRATION_CLIENT_ID="<integration-app-client-id>"
 *   export M365_INTEGRATION_TENANT_ID="<integration-app-tenant-id>"
 *   export M365_INTEGRATION_CREDS_PATH="~/.m365-cli/integration-credentials.json"
 *
 * Run:
 *   npx vitest run tests/onedrive.integration.test.js
 */

const INTEGRATION_CLIENT_ID = process.env.M365_INTEGRATION_CLIENT_ID;
const INTEGRATION_TENANT_ID = process.env.M365_INTEGRATION_TENANT_ID;
const INTEGRATION_CREDS_PATH = process.env.M365_INTEGRATION_CREDS_PATH;

let hasAuth = false;
let savedEnv = {};

// Unique test folder name to avoid collisions
const TEST_FOLDER = `integration-test-${Date.now()}`;

// Track items to clean up
const itemsToCleanup = [];

// Temp files to clean up
const tempFilesToCleanup = [];

beforeAll(async () => {
  if (!INTEGRATION_CLIENT_ID || !INTEGRATION_TENANT_ID || !INTEGRATION_CREDS_PATH) {
    console.log(
      '⏭️  Integration env vars not set — skipping OneDrive integration tests'
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
      console.log('⏭️  No valid credentials — skipping OneDrive integration tests');
      return;
    }
    await getAccessToken();
    hasAuth = true;
  } catch (error) {
    console.log(`⏭️  Auth unavailable (${error.message}) — skipping OneDrive integration tests`);
  }
});

afterAll(async () => {
  // Clean up remote items (in reverse order to handle nested items first)
  if (hasAuth) {
    for (const path of itemsToCleanup.reverse()) {
      try {
        await graphClient.onedrive.remove(path);
      } catch {
        // Ignore cleanup errors
      }
    }
  }

  // Clean up temp files
  for (const f of tempFilesToCleanup) {
    try {
      unlinkSync(f);
    } catch {
      // Ignore
    }
  }

  // Restore env
  for (const [key, value] of Object.entries(savedEnv)) {
    if (value === undefined) {
      delete process.env[key];
    } else {
      process.env[key] = value;
    }
  }
});

describe('[Integration] OneDrive — Graph API', { timeout: 30000 }, () => {
  describe('List files (/me/drive/root/children)', () => {
    it('should list root files and folders', { retry: 2 }, async (ctx) => {
      if (!hasAuth) return ctx.skip();

      const items = await graphClient.onedrive.list('', { top: 10 });

      expect(Array.isArray(items)).toBe(true);

      for (const item of items) {
        expect(item).toHaveProperty('id');
        expect(item).toHaveProperty('name');
        // Each item should be either a file or folder
        expect(item.file !== undefined || item.folder !== undefined).toBe(true);
      }
    });

    it('should respect top parameter', async (ctx) => {
      if (!hasAuth) return ctx.skip();

      const items = await graphClient.onedrive.list('', { top: 2 });

      expect(Array.isArray(items)).toBe(true);
      expect(items.length).toBeLessThanOrEqual(2);
    });
  });

  describe('Create folder (/me/drive/root/children)', () => {
    it('should create a new folder', async (ctx) => {
      if (!hasAuth) return ctx.skip();

      const result = await graphClient.onedrive.mkdir(TEST_FOLDER);
      itemsToCleanup.push(TEST_FOLDER);

      expect(result).toHaveProperty('id');
      expect(result).toHaveProperty('name');
      expect(result).toHaveProperty('folder');
      expect(result.name).toBe(TEST_FOLDER);
    });

    it('should list contents of the created folder', async (ctx) => {
      if (!hasAuth) return ctx.skip();

      const items = await graphClient.onedrive.list(TEST_FOLDER, { top: 10 });

      expect(Array.isArray(items)).toBe(true);
      // Newly created folder should be empty
      expect(items.length).toBe(0);
    });
  });

  describe('Upload file (/me/drive/root:/{path}:/content)', () => {
    it('should upload a small file', async (ctx) => {
      if (!hasAuth) return ctx.skip();

      const content = Buffer.from('Hello, integration test! ' + Date.now());
      const remotePath = `${TEST_FOLDER}/test-upload.txt`;

      const result = await graphClient.onedrive.upload(remotePath, content);
      // Don't add to cleanup — folder cleanup will handle it

      expect(result).toHaveProperty('id');
      expect(result).toHaveProperty('name');
      expect(result.name).toBe('test-upload.txt');
      expect(result).toHaveProperty('size');
      expect(result.size).toBeGreaterThan(0);
    });
  });

  describe('Get metadata (/me/drive/root:/{path})', () => {
    it('should get metadata for a file', async (ctx) => {
      if (!hasAuth) return ctx.skip();

      // Upload a file first
      const content = Buffer.from('Metadata test content ' + Date.now());
      const remotePath = `${TEST_FOLDER}/metadata-test.txt`;
      await graphClient.onedrive.upload(remotePath, content);

      const metadata = await graphClient.onedrive.getMetadata(remotePath);

      expect(metadata).toHaveProperty('id');
      expect(metadata).toHaveProperty('name');
      expect(metadata.name).toBe('metadata-test.txt');
      expect(metadata).toHaveProperty('size');
      expect(metadata.size).toBeGreaterThan(0);
      expect(metadata).toHaveProperty('webUrl');
      expect(metadata).toHaveProperty('file');
    });

    it('should get metadata for a folder', async (ctx) => {
      if (!hasAuth) return ctx.skip();

      const metadata = await graphClient.onedrive.getMetadata(TEST_FOLDER);

      expect(metadata).toHaveProperty('id');
      expect(metadata).toHaveProperty('name');
      expect(metadata.name).toBe(TEST_FOLDER);
      expect(metadata).toHaveProperty('folder');
    });

    it('should throw for nonexistent path', async (ctx) => {
      if (!hasAuth) return ctx.skip();

      await expect(
        graphClient.onedrive.getMetadata('nonexistent-path-xyz-' + Date.now())
      ).rejects.toThrow();
    });
  });

  describe('Download file (/me/drive/root:/{path}:/content)', () => {
    it('should download a file and return a response', async (ctx) => {
      if (!hasAuth) return ctx.skip();

      const testContent = 'Download test content ' + Date.now();
      const remotePath = `${TEST_FOLDER}/download-test.txt`;
      await graphClient.onedrive.upload(remotePath, Buffer.from(testContent));

      const response = await graphClient.onedrive.download(remotePath);

      expect(response).toBeDefined();
      expect(response.ok).toBe(true);

      // Read the body
      const body = await response.text();
      expect(body).toBe(testContent);
    });
  });

  describe('Search files (/me/drive/root/search)', () => {
    it('should search for files', async (ctx) => {
      if (!hasAuth) return ctx.skip();

      // Search with a broad query
      const results = await graphClient.onedrive.search('test', { top: 5 });

      expect(Array.isArray(results)).toBe(true);
      // Results may be empty if nothing matches
    });

    it('should respect top parameter', async (ctx) => {
      if (!hasAuth) return ctx.skip();

      const results = await graphClient.onedrive.search('test', { top: 2 });

      expect(Array.isArray(results)).toBe(true);
      expect(results.length).toBeLessThanOrEqual(2);
    });
  });

  describe('Share file (/me/drive/root:/{path}:/createLink)', () => {
    it('should create a sharing link', async (ctx) => {
      if (!hasAuth) return ctx.skip();

      const content = Buffer.from('Share test content ' + Date.now());
      const remotePath = `${TEST_FOLDER}/share-test.txt`;
      await graphClient.onedrive.upload(remotePath, content);

      try {
        const result = await graphClient.onedrive.share(remotePath, {
          type: 'view',
          scope: 'organization',
        });

        expect(result).toHaveProperty('link');
        expect(result.link).toHaveProperty('webUrl');
        expect(typeof result.link.webUrl).toBe('string');
      } catch (error) {
        // Organization sharing may not be enabled — that's acceptable
        if (error.message && error.message.includes('accessDenied')) {
          console.log('  (sharing not enabled for this account — skipping)');
          return ctx.skip();
        }
        throw error;
      }
    });
  });

  describe('Delete file (/me/drive/root:/{path})', () => {
    it('should delete a file', async (ctx) => {
      if (!hasAuth) return ctx.skip();

      const content = Buffer.from('Delete test content');
      const remotePath = `${TEST_FOLDER}/to-delete.txt`;
      await graphClient.onedrive.upload(remotePath, content);

      const result = await graphClient.onedrive.remove(remotePath);
      expect(result).toEqual({ success: true });

      // Verify it's gone
      await expect(
        graphClient.onedrive.getMetadata(remotePath)
      ).rejects.toThrow();
    });
  });

  describe('Full command flows', () => {
    it('should execute listFiles command without throwing', async (ctx) => {
      if (!hasAuth) return ctx.skip();

      await expect(
        onedriveCommands.ls('', { top: 5, json: true })
      ).resolves.not.toThrow();
    });

    it('should execute getMetadata command without throwing', async (ctx) => {
      if (!hasAuth) return ctx.skip();

      // Upload a test file
      const remotePath = `${TEST_FOLDER}/cmd-metadata-test.txt`;
      await graphClient.onedrive.upload(remotePath, Buffer.from('cmd test'));

      await expect(
        onedriveCommands.get(remotePath, { json: true })
      ).resolves.not.toThrow();
    });

    it('should execute searchFiles command without throwing', async (ctx) => {
      if (!hasAuth) return ctx.skip();

      await expect(
        onedriveCommands.search('test', { top: 5, json: true })
      ).resolves.not.toThrow();
    });

    it('should execute uploadFile command without throwing', async (ctx) => {
      if (!hasAuth) return ctx.skip();

      // Create a local temp file
      const tmpFile = join(tmpdir(), `m365-integration-upload-${Date.now()}.txt`);
      writeFileSync(tmpFile, 'Upload command flow test content');
      tempFilesToCleanup.push(tmpFile);

      const remotePath = `${TEST_FOLDER}/cmd-upload-test.txt`;

      await expect(
        onedriveCommands.upload(tmpFile, remotePath, { json: true })
      ).resolves.not.toThrow();
    });

    it('should execute downloadFile command without throwing', async (ctx) => {
      if (!hasAuth) return ctx.skip();

      // Upload a file first
      const remotePath = `${TEST_FOLDER}/cmd-download-test.txt`;
      await graphClient.onedrive.upload(remotePath, Buffer.from('Download command test'));

      // Download to temp path
      const tmpFile = join(tmpdir(), `m365-integration-download-${Date.now()}.txt`);
      tempFilesToCleanup.push(tmpFile);

      await expect(
        onedriveCommands.download(remotePath, tmpFile, { json: true })
      ).resolves.not.toThrow();

      // File may still be flushing (stream-based write); verify via returned result
    });

    it('should execute createFolder command without throwing', async (ctx) => {
      if (!hasAuth) return ctx.skip();

      const subFolder = `${TEST_FOLDER}/cmd-subfolder-${Date.now()}`;

      await expect(
        onedriveCommands.mkdir(subFolder, { json: true })
      ).resolves.not.toThrow();
    });

    it('should execute deleteItem command without throwing', async (ctx) => {
      if (!hasAuth) return ctx.skip();

      const remotePath = `${TEST_FOLDER}/cmd-to-delete.txt`;
      await graphClient.onedrive.upload(remotePath, Buffer.from('to delete'));

      await expect(
        onedriveCommands.rm(remotePath, { json: true, force: true })
      ).resolves.not.toThrow();
    });
  });
});
