import { describe, it, expect, beforeAll, afterAll } from 'vitest';
import graphClient from '../../src/graph/client.js';
import config from '../../src/utils/config.js';
import { getAvailableAccounts, setupAuth, teardownAuth } from './helpers/setup.js';

const accounts = getAvailableAccounts({ cloud: 'china' });

async function suppressConsole(fn) {
  const origLog = console.log;
  const origError = console.error;
  console.log = () => {};
  console.error = () => {};
  try {
    return await fn();
  } finally {
    console.log = origLog;
    console.error = origError;
  }
}

describe('[Integration] 21Vianet (China) Cloud — Graph API', { timeout: 60000 }, () => {
  if (accounts.length === 0) {
    it('requires 21Vianet integration env vars (M365_INTEGRATION_CHINA_*)', (ctx) => {
      console.log('⏭️  21Vianet integration env vars not set — skipping China cloud tests');
      ctx.skip();
    });
    return;
  }

  describe.each(accounts)('$type account ($cloud)', (account) => {
    let hasAuth = false;
    let savedEnv = {};

    beforeAll(async () => {
      const result = await setupAuth(account);
      hasAuth = result.hasAuth;
      savedEnv = result.savedEnv;
    });

    afterAll(() => teardownAuth(savedEnv));

    describe('Cloud config resolution', () => {
      it('should resolve active cloud as china', (ctx) => {
        if (!hasAuth) return ctx.skip();

        expect(config.getActiveCloud()).toBe('china');
      });

      it('should resolve graphApiUrl to chinacloudapi.cn', (ctx) => {
        if (!hasAuth) return ctx.skip();

        const graphApiUrl = config.get('graphApiUrl');
        expect(graphApiUrl).toContain('chinacloudapi.cn');
      });

      it('should resolve authUrl to chinacloudapi.cn', (ctx) => {
        if (!hasAuth) return ctx.skip();

        const authUrl = config.get('authUrl');
        expect(authUrl).toContain('chinacloudapi.cn');
      });

      it('should resolve scopePrefix to chinacloudapi.cn', (ctx) => {
        if (!hasAuth) return ctx.skip();

        const scopePrefix = config.get('scopePrefix');
        expect(scopePrefix).toContain('chinacloudapi.cn');
      });
    });

    describe('Graph API connectivity', () => {
      it('should get current user profile via China Graph endpoint', { retry: 2 }, async (ctx) => {
        if (!hasAuth) return ctx.skip();

        const user = await graphClient.getCurrentUser();

        expect(user).toHaveProperty('id');
        expect(user).toHaveProperty('displayName');
        expect(user.mail || user.userPrincipalName).toBeTruthy();
      });

      it('should list inbox emails via China Graph endpoint', async (ctx) => {
        if (!hasAuth) return ctx.skip();

        const mails = await graphClient.mail.list({ top: 3, folder: 'inbox' });

        expect(Array.isArray(mails)).toBe(true);
        for (const mail of mails) {
          expect(mail).toHaveProperty('id');
          expect(mail).toHaveProperty('subject');
        }
      });

      it('should list calendar events via China Graph endpoint', async (ctx) => {
        if (!hasAuth) return ctx.skip();

        const events = await graphClient.calendar.list({ days: 7, top: 5 });

        expect(Array.isArray(events)).toBe(true);
      });

      it('should list OneDrive root via China Graph endpoint', async (ctx) => {
        if (!hasAuth) return ctx.skip();

        const items = await graphClient.onedrive.list({ top: 5 });

        expect(Array.isArray(items)).toBe(true);
      });
    });

    describe('Token validity against China endpoint', () => {
      it('should make a raw fetch to China Graph API successfully', async (ctx) => {
        if (!hasAuth) return ctx.skip();

        const { getAccessToken } = await import('../../src/auth/token-manager.js');
        const token = await getAccessToken();
        const graphApiUrl = config.get('graphApiUrl');

        const response = await fetch(`${graphApiUrl}/me?$select=id,displayName`, {
          headers: { Authorization: `Bearer ${token}` },
        });

        expect(response.ok).toBe(true);

        const data = await response.json();
        expect(data).toHaveProperty('id');
        expect(data).toHaveProperty('displayName');
      });
    });
  });
});
