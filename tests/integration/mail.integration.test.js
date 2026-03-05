import { describe, it, expect, beforeAll, afterAll } from 'vitest';
import { existsSync } from 'fs';
import { resolve } from 'path';
import { loadCreds, getAccessToken } from '../../src/auth/token-manager.js';
import graphClient from '../../src/graph/client.js';
import mailCommands from '../../src/commands/mail.js';

/**
 * Integration tests for Mail feature.
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
 *   npx vitest run tests/mail.integration.test.js
 */

const INTEGRATION_CLIENT_ID = process.env.M365_INTEGRATION_CLIENT_ID;
const INTEGRATION_TENANT_ID = process.env.M365_INTEGRATION_TENANT_ID;
const INTEGRATION_CREDS_PATH = process.env.M365_INTEGRATION_CREDS_PATH;

let hasAuth = false;
let savedEnv = {};

beforeAll(async () => {
  if (!INTEGRATION_CLIENT_ID || !INTEGRATION_TENANT_ID || !INTEGRATION_CREDS_PATH) {
    console.log(
      '⏭️  Integration env vars not set — skipping mail integration tests'
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
      console.log('⏭️  No valid credentials — skipping mail integration tests');
      return;
    }
    await getAccessToken();
    hasAuth = true;
  } catch (error) {
    console.log(`⏭️  Auth unavailable (${error.message}) — skipping mail integration tests`);
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

describe('[Integration] Mail — Graph API', { timeout: 30000 }, () => {
  describe('Current user', () => {
    it('should get current user profile', { retry: 2 }, async (ctx) => {
      if (!hasAuth) return ctx.skip();

      const user = await graphClient.getCurrentUser();

      expect(user).toHaveProperty('id');
      expect(user).toHaveProperty('displayName');
      // mail or userPrincipalName should exist
      expect(user.mail || user.userPrincipalName).toBeTruthy();
    });
  });

  describe('List emails (/me/messages)', () => {
    it('should list inbox emails', async (ctx) => {
      if (!hasAuth) return ctx.skip();

      const mails = await graphClient.mail.list({ top: 5, folder: 'inbox' });

      expect(Array.isArray(mails)).toBe(true);

      for (const mail of mails) {
        expect(mail).toHaveProperty('id');
        expect(mail).toHaveProperty('subject');
        expect(mail).toHaveProperty('receivedDateTime');
        expect(mail).toHaveProperty('isRead');
      }
    });

    it('should respect top parameter', async (ctx) => {
      if (!hasAuth) return ctx.skip();

      const mails = await graphClient.mail.list({ top: 2, folder: 'inbox' });

      expect(Array.isArray(mails)).toBe(true);
      expect(mails.length).toBeLessThanOrEqual(2);
    });

    it('should list sent items', async (ctx) => {
      if (!hasAuth) return ctx.skip();

      const mails = await graphClient.mail.list({ top: 5, folder: 'sent' });

      expect(Array.isArray(mails)).toBe(true);
      // Sent items may be empty for a test account, that's fine
    });

    it('should list drafts', async (ctx) => {
      if (!hasAuth) return ctx.skip();

      const mails = await graphClient.mail.list({ top: 5, folder: 'drafts' });

      expect(Array.isArray(mails)).toBe(true);
    });
  });

  describe('Get email (/me/messages/{id})', () => {
    it('should retrieve an email by ID with body and attachments info', async (ctx) => {
      if (!hasAuth) return ctx.skip();

      // Get a message from inbox first
      const mails = await graphClient.mail.list({ top: 1, folder: 'inbox' });
      if (mails.length === 0) {
        console.log('  (no emails in inbox — cannot test get)');
        return ctx.skip();
      }

      const mail = await graphClient.mail.get(mails[0].id);

      expect(mail).toHaveProperty('id');
      expect(mail).toHaveProperty('subject');
      expect(mail).toHaveProperty('body');
      expect(mail.body).toHaveProperty('content');
      expect(mail.body).toHaveProperty('contentType');
      expect(mail).toHaveProperty('from');
    });
  });

  describe('Search emails (/me/messages?$search)', () => {
    it('should search emails by keyword', async (ctx) => {
      if (!hasAuth) return ctx.skip();

      // Search for a common word likely to match something
      const results = await graphClient.mail.search('test', { top: 5 });

      expect(Array.isArray(results)).toBe(true);
      // May be empty if no emails match, that's fine
    });

    it('should respect top parameter in search', async (ctx) => {
      if (!hasAuth) return ctx.skip();

      const results = await graphClient.mail.search('the', { top: 2 });

      expect(Array.isArray(results)).toBe(true);
      expect(results.length).toBeLessThanOrEqual(2);
    });
  });

  describe('Attachments (/me/messages/{id}/attachments)', () => {
    it('should list attachments for an email', async (ctx) => {
      if (!hasAuth) return ctx.skip();

      // Find an email with attachments
      const mails = await graphClient.mail.list({ top: 20, folder: 'inbox' });
      const mailWithAttachments = mails.find(m => m.hasAttachments);

      if (!mailWithAttachments) {
        console.log('  (no emails with attachments found — skipping)');
        return ctx.skip();
      }

      const attachments = await graphClient.mail.attachments(mailWithAttachments.id);

      expect(Array.isArray(attachments)).toBe(true);
      expect(attachments.length).toBeGreaterThan(0);

      for (const att of attachments) {
        expect(att).toHaveProperty('id');
        expect(att).toHaveProperty('name');
        expect(att).toHaveProperty('size');
      }
    });

    it('should return empty array for email without attachments', async (ctx) => {
      if (!hasAuth) return ctx.skip();

      const mails = await graphClient.mail.list({ top: 20, folder: 'inbox' });
      const mailWithout = mails.find(m => !m.hasAttachments);

      if (!mailWithout) {
        console.log('  (all emails have attachments — skipping)');
        return ctx.skip();
      }

      const attachments = await graphClient.mail.attachments(mailWithout.id);

      expect(Array.isArray(attachments)).toBe(true);
      expect(attachments.length).toBe(0);
    });
  });

  describe('Download attachment (/me/messages/{id}/attachments/{id})', () => {
    it('should download a specific attachment', async (ctx) => {
      if (!hasAuth) return ctx.skip();

      // Find an email with attachments
      const mails = await graphClient.mail.list({ top: 20, folder: 'inbox' });
      const mailWithAttachments = mails.find(m => m.hasAttachments);

      if (!mailWithAttachments) {
        console.log('  (no emails with attachments found — skipping)');
        return ctx.skip();
      }

      const attachments = await graphClient.mail.attachments(mailWithAttachments.id);
      if (attachments.length === 0) {
        return ctx.skip();
      }

      const attachment = await graphClient.mail.downloadAttachment(
        mailWithAttachments.id,
        attachments[0].id
      );

      expect(attachment).toHaveProperty('name');
      expect(attachment).toHaveProperty('contentBytes');
      expect(typeof attachment.contentBytes).toBe('string');
    });
  });

  describe('Send email (/me/sendMail)', () => {
    it('should send an email to self', async (ctx) => {
      if (!hasAuth) return ctx.skip();

      // Get current user email for self-send
      const user = await graphClient.getCurrentUser();
      const selfEmail = user.mail || user.userPrincipalName;
      if (!selfEmail) {
        console.log('  (cannot determine self email — skipping send test)');
        return ctx.skip();
      }

      const message = {
        subject: '[Integration Test] Mail Send Test',
        body: {
          contentType: 'HTML',
          content: '<p>This is an automated integration test email. Safe to delete.</p>',
        },
        toRecipients: [
          { emailAddress: { address: selfEmail } },
        ],
      };

      // sendMail returns { success: true } on 202 Accepted
      const result = await graphClient.mail.send(message);

      // Graph API returns 202 Accepted (empty response), our client returns { success: true }
      expect(result).toBeDefined();
    });
  });

  describe('Full command flows', () => {
    it('should execute listMails command without throwing', async (ctx) => {
      if (!hasAuth) return ctx.skip();

      await expect(
        mailCommands.list({ top: 5, folder: 'inbox', json: true })
      ).resolves.not.toThrow();
    });

    it('should execute readMail command without throwing', async (ctx) => {
      if (!hasAuth) return ctx.skip();

      const mails = await graphClient.mail.list({ top: 1, folder: 'inbox' });
      if (mails.length === 0) {
        return ctx.skip();
      }

      await expect(
        mailCommands.read(mails[0].id, { json: true })
      ).resolves.not.toThrow();
    });

    it('should execute readMail with --force without throwing', async (ctx) => {
      if (!hasAuth) return ctx.skip();

      const mails = await graphClient.mail.list({ top: 1, folder: 'inbox' });
      if (mails.length === 0) {
        return ctx.skip();
      }

      await expect(
        mailCommands.read(mails[0].id, { json: true, force: true })
      ).resolves.not.toThrow();
    });

    it('should execute searchMails command without throwing', async (ctx) => {
      if (!hasAuth) return ctx.skip();

      await expect(
        mailCommands.search('test', { top: 5, json: true })
      ).resolves.not.toThrow();
    });

    it('should execute listAttachments command without throwing', async (ctx) => {
      if (!hasAuth) return ctx.skip();

      const mails = await graphClient.mail.list({ top: 1, folder: 'inbox' });
      if (mails.length === 0) {
        return ctx.skip();
      }

      try {
        await mailCommands.attachments(mails[0].id, { json: true });
      } catch (error) {
        // process.exit called by handleError for permission issues
        if (error.message?.includes('process.exit') || error.statusCode === 403) {
          return ctx.skip();
        }
        throw error;
      }
    });

    it('should handle empty search results gracefully', async (ctx) => {
      if (!hasAuth) return ctx.skip();

      await expect(
        mailCommands.search('zzxqqnonexistent99integration', { top: 5, json: true })
      ).resolves.not.toThrow();
    });
  });
});
