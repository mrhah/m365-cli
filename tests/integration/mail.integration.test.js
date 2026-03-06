import { describe, it, expect, beforeAll, afterAll } from 'vitest';
import graphClient from '../../src/graph/client.js';
import mailCommands from '../../src/commands/mail.js';
import { getAvailableAccounts, setupAuth, teardownAuth } from './helpers/setup.js';

const accounts = getAvailableAccounts();

describe('[Integration] Mail — Graph API', { timeout: 90000 }, () => {
  if (accounts.length === 0) {
    it('requires integration env vars', (ctx) => {
      console.log('⏭️  Integration env vars not set — skipping mail integration tests');
      ctx.skip();
    });
    return;
  }

  describe.each(accounts)('$type account', (account) => {
    let hasAuth = false;
    let savedEnv = {};

    beforeAll(async () => {
      const result = await setupAuth(account);
      hasAuth = result.hasAuth;
      savedEnv = result.savedEnv;
    });

    afterAll(() => teardownAuth(savedEnv));

    describe('Current user', () => {
      it('should get current user profile', { retry: 2 }, async (ctx) => {
        if (!hasAuth) return ctx.skip();

        const user = await graphClient.getCurrentUser();

        expect(user).toHaveProperty('id');
        expect(user).toHaveProperty('displayName');
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

        const results = await graphClient.mail.search('test', { top: 5 });

        expect(Array.isArray(results)).toBe(true);
      });

      it('should respect top parameter in search', async (ctx) => {
        if (!hasAuth) return ctx.skip();

        const results = await graphClient.mail.search('the', { top: 2 });

        expect(Array.isArray(results)).toBe(true);
        expect(results.length).toBeLessThanOrEqual(2);
      });
    });

    describe('Attachments (/me/messages/{id}/attachments)', () => {
      let sentMailId = null;

      // Send an email with attachment to self so we have a guaranteed test target
      beforeAll(async () => {
        if (!hasAuth) return;
        try {
          const user = await graphClient.getCurrentUser();
          const selfEmail = user.mail || user.userPrincipalName;
          if (!selfEmail) return;

          const testContent = Buffer.from('integration-test-attachment-content').toString('base64');
          const message = {
            subject: `[Integration Test] Attachment Test (${account.type}) ${Date.now()}`,
            body: { contentType: 'Text', content: 'Automated test — safe to delete.' },
            toRecipients: [{ emailAddress: { address: selfEmail } }],
            attachments: [{
              '@odata.type': '#microsoft.graph.fileAttachment',
              name: 'test-attachment.txt',
              contentBytes: testContent,
            }],
          };
          await graphClient.mail.send(message);

          // Wait for delivery
          for (let i = 0; i < 12; i++) {
            await new Promise(r => setTimeout(r, 5000));
            const mails = await graphClient.mail.list({ top: 5, folder: 'inbox' });
            const found = mails.find(m => m.subject?.includes('Attachment Test') && m.hasAttachments);
            if (found) { sentMailId = found.id; break; }
          }
        } catch { /* best effort */ }
      }, 75000);

      it('should list attachments for an email', async (ctx) => {
        if (!hasAuth || !sentMailId) return ctx.skip();

        const attachments = await graphClient.mail.attachments(sentMailId);

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

        const user = await graphClient.getCurrentUser();
        const selfEmail = user.mail || user.userPrincipalName;
        if (!selfEmail) {
          console.log('  (cannot determine self email — skipping send test)');
          return ctx.skip();
        }

        const message = {
          subject: `[Integration Test] Mail Send Test (${account.type})`,
          body: {
            contentType: 'HTML',
            content: '<p>This is an automated integration test email. Safe to delete.</p>',
          },
          toRecipients: [
            { emailAddress: { address: selfEmail } },
          ],
        };

        const result = await graphClient.mail.send(message);

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

        try {
          await mailCommands.read(mails[0].id, { json: true });
        } catch (error) {
          if (error.message?.includes('process.exit')) {
            return ctx.skip();
          }
          throw error;
        }
      });

      it('should execute readMail with --force without throwing', async (ctx) => {
        if (!hasAuth) return ctx.skip();

        const mails = await graphClient.mail.list({ top: 1, folder: 'inbox' });
        if (mails.length === 0) {
          return ctx.skip();
        }

        try {
          await mailCommands.read(mails[0].id, { json: true, force: true });
        } catch (error) {
          if (error.message?.includes('process.exit')) {
            return ctx.skip();
          }
          throw error;
        }
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
});
