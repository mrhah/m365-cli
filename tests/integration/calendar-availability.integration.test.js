import { describe, it, expect, beforeAll, afterAll } from 'vitest';
import graphClient from '../../src/graph/client.js';
import calendarCommands from '../../src/commands/calendar.js';
import { enrichScheduleResult } from '../../src/utils/availability.js';
import { getAvailableAccounts, setupAuth, teardownAuth } from './helpers/setup.js';

const accounts = getAvailableAccounts({ workOnly: true });

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

describe('[Integration] Calendar Availability — Graph API', { timeout: 30000 }, () => {
  if (accounts.length === 0) {
    it('requires integration env vars', (ctx) => {
      console.log('⏭️  Integration env vars not set — skipping calendar availability integration tests');
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

    afterAll(async () => {
      teardownAuth(savedEnv);
    });

    describe('getSchedule (/me/calendar/getSchedule)', () => {
      it('should call getSchedule API and get array response', async (ctx) => {
        if (!hasAuth) return ctx.skip();

        const start = new Date();
        const end = new Date(Date.now() + 4 * 60 * 60 * 1000);
        const me = await graphClient.getCurrentUser();
        const email = me.mail || me.userPrincipalName;

        const result = await graphClient.calendar.getSchedule(
          [email],
          start.toISOString(),
          end.toISOString(),
          { availabilityViewInterval: 30 }
        );

        expect(Array.isArray(result)).toBe(true);
      });

      it('should return schedule data with expected shape', async (ctx) => {
        if (!hasAuth) return ctx.skip();

        const start = new Date();
        const end = new Date(Date.now() + 4 * 60 * 60 * 1000);
        const me = await graphClient.getCurrentUser();
        const email = me.mail || me.userPrincipalName;

        const result = await graphClient.calendar.getSchedule(
          [email],
          start.toISOString(),
          end.toISOString(),
          { availabilityViewInterval: 30 }
        );

        expect(Array.isArray(result)).toBe(true);
        if (result.length > 0) {
          const first = result[0];
          expect(first).toHaveProperty('scheduleId');
          expect(first).toHaveProperty('availabilityView');
          expect(first).toHaveProperty('scheduleItems');
          expect(first).toHaveProperty('workingHours');
          expect(Array.isArray(first.scheduleItems)).toBe(true);
        }
      });

      it('should respect availabilityViewInterval parameter', async (ctx) => {
        if (!hasAuth) return ctx.skip();

        const start = new Date();
        const end = new Date(Date.now() + 4 * 60 * 60 * 1000);
        const me = await graphClient.getCurrentUser();
        const email = me.mail || me.userPrincipalName;

        const result30 = await graphClient.calendar.getSchedule(
          [email],
          start.toISOString(),
          end.toISOString(),
          { availabilityViewInterval: 30 }
        );

        const result60 = await graphClient.calendar.getSchedule(
          [email],
          start.toISOString(),
          end.toISOString(),
          { availabilityViewInterval: 60 }
        );

        expect(Array.isArray(result30)).toBe(true);
        expect(Array.isArray(result60)).toBe(true);

        if (result30.length > 0 && result60.length > 0) {
          expect(result30[0].availabilityView.length).toBeGreaterThanOrEqual(result60[0].availabilityView.length);
        }
      });

      it('should return enriched schedule data with slots and segments', async (ctx) => {
        if (!hasAuth) return ctx.skip();

        const start = new Date();
        const end = new Date(Date.now() + 4 * 60 * 60 * 1000);
        const me = await graphClient.getCurrentUser();
        const email = me.mail || me.userPrincipalName;
        const interval = 30;
        const timeZone = await graphClient.getTimezone();

        const result = await graphClient.calendar.getSchedule(
          [email],
          start.toISOString(),
          end.toISOString(),
          { availabilityViewInterval: interval, timezone: timeZone }
        );

        expect(Array.isArray(result)).toBe(true);
        if (result.length > 0) {
          const enriched = enrichScheduleResult(result[0], {
            startDateTime: start.toISOString().slice(0, 19),
            endDateTime: end.toISOString().slice(0, 19),
            timeZone,
            intervalMinutes: interval,
          });

          expect(Array.isArray(enriched.slots)).toBe(true);
          expect(Array.isArray(enriched.segments)).toBe(true);
          expect(Array.isArray(enriched.freeSegments)).toBe(true);
        }
      });

      it('should execute full command flow without throwing', async (ctx) => {
        if (!hasAuth) return ctx.skip();

        const start = new Date();
        const end = new Date(Date.now() + 4 * 60 * 60 * 1000);
        const me = await graphClient.getCurrentUser();
        const email = me.mail || me.userPrincipalName;

        await expect(
          suppressConsole(() => calendarCommands.availability({
            users: email,
            startDateTime: start.toISOString(),
            endDateTime: end.toISOString(),
            interval: 30,
            json: true,
          }))
        ).resolves.not.toThrow();
      });
    });
  });
});
