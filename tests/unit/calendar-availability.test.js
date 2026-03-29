import { describe, it, expect, vi, beforeEach } from 'vitest';
import { mapStatus, expandSlots, mergeSegments, filterFreeSegments, enrichScheduleResult } from '../../src/utils/availability.js';

const mockCalendarGetSchedule = vi.fn();
const mockGetTimezone = vi.fn();

vi.mock('../../src/graph/client.js', () => ({
  default: {
    getTimezone: (...args) => mockGetTimezone(...args),
    calendar: {
      getSchedule: (...args) => mockCalendarGetSchedule(...args),
    },
  },
}));

const mockOutputAvailability = vi.fn();

vi.mock('../../src/utils/output.js', async () => {
  const actual = await vi.importActual('../../src/utils/output.js');
  return {
    ...actual,
    outputAvailability: (...args) => mockOutputAvailability(...args),
    __actualOutputAvailability: actual.outputAvailability,
  };
});

const mockHandleError = vi.fn();
vi.mock('../../src/utils/error.js', () => ({
  handleError: (...args) => mockHandleError(...args),
}));

const mockEnsureWorkAccount = vi.fn();
vi.mock('../../src/utils/account.js', () => ({
  ensureWorkAccount: (...args) => mockEnsureWorkAccount(...args),
}));

import { getAvailability } from '../../src/commands/calendar.js';
import { __actualOutputAvailability as realOutputAvailability } from '../../src/utils/output.js';

describe('Calendar availability command', () => {
  beforeEach(() => {
    vi.clearAllMocks();
    mockGetTimezone.mockResolvedValue('China Standard Time');
  });

  it('should call ensureWorkAccount with correct command name', async () => {
    mockCalendarGetSchedule.mockResolvedValue([]);

    await getAvailability({
      users: 'user1@contoso.com',
      startDateTime: '2026-04-01T09:00:00',
      endDateTime: '2026-04-01T18:00:00',
      json: false,
    });

    expect(mockEnsureWorkAccount).toHaveBeenCalledWith('calendar availability get');
  });

  it('should throw error when users is missing', async () => {
    await getAvailability({
      startDateTime: '2026-04-01T09:00:00',
      endDateTime: '2026-04-01T18:00:00',
      json: false,
    });

    expect(mockHandleError).toHaveBeenCalled();
    expect(mockHandleError.mock.calls[0][0].message).toBe('Users are required');
  });

  it('should throw error when startDateTime is missing', async () => {
    await getAvailability({
      users: 'user1@contoso.com',
      endDateTime: '2026-04-01T18:00:00',
      json: false,
    });

    expect(mockHandleError).toHaveBeenCalled();
    expect(mockHandleError.mock.calls[0][0].message).toBe('Start date/time is required');
  });

  it('should throw error when endDateTime is missing', async () => {
    await getAvailability({
      users: 'user1@contoso.com',
      startDateTime: '2026-04-01T09:00:00',
      json: false,
    });

    expect(mockHandleError).toHaveBeenCalled();
    expect(mockHandleError.mock.calls[0][0].message).toBe('End date/time is required');
  });

  it('should call getSchedule with correct parameters', async () => {
    mockCalendarGetSchedule.mockResolvedValue([]);

    await getAvailability({
      users: 'user1@contoso.com, user2@contoso.com',
      startDateTime: '2026-04-01T09:00:00',
      endDateTime: '2026-04-01T18:00:00',
      interval: 60,
      json: false,
    });

    expect(mockCalendarGetSchedule).toHaveBeenCalledWith(
      ['user1@contoso.com', 'user2@contoso.com'],
      '2026-04-01T09:00:00',
      '2026-04-01T18:00:00',
      {
        availabilityViewInterval: 60,
        timezone: 'China Standard Time',
      }
    );
  });

  it('should pass default interval of 30 when not specified', async () => {
    mockCalendarGetSchedule.mockResolvedValue([]);

    await getAvailability({
      users: 'user1@contoso.com',
      startDateTime: '2026-04-01T09:00:00',
      endDateTime: '2026-04-01T18:00:00',
      json: false,
    });

    expect(mockCalendarGetSchedule).toHaveBeenCalledWith(
      ['user1@contoso.com'],
      '2026-04-01T09:00:00',
      '2026-04-01T18:00:00',
      {
        availabilityViewInterval: 30,
        timezone: 'China Standard Time',
      }
    );
  });

  it('should pass custom interval when specified', async () => {
    mockCalendarGetSchedule.mockResolvedValue([]);

    await getAvailability({
      users: 'user1@contoso.com',
      startDateTime: '2026-04-01T09:00:00',
      endDateTime: '2026-04-01T18:00:00',
      interval: 15,
      json: false,
    });

    expect(mockCalendarGetSchedule).toHaveBeenCalledWith(
      ['user1@contoso.com'],
      '2026-04-01T09:00:00',
      '2026-04-01T18:00:00',
      {
        availabilityViewInterval: 15,
        timezone: 'China Standard Time',
      }
    );
  });

  it('should pass timezone when specified', async () => {
    mockCalendarGetSchedule.mockResolvedValue([]);

    await getAvailability({
      users: 'user1@contoso.com',
      startDateTime: '2026-04-01T09:00:00',
      endDateTime: '2026-04-01T18:00:00',
      timezone: 'Pacific Standard Time',
      json: false,
    });

    expect(mockCalendarGetSchedule).toHaveBeenCalledWith(
      ['user1@contoso.com'],
      '2026-04-01T09:00:00',
      '2026-04-01T18:00:00',
      {
        availabilityViewInterval: 30,
        timezone: 'Pacific Standard Time',
      }
    );
  });

  it('should call outputAvailability with result and correct options', async () => {
    const result = [{ scheduleId: 'user1@contoso.com', availabilityView: '0222' }];
    mockCalendarGetSchedule.mockResolvedValue(result);

    await getAvailability({
      users: 'user1@contoso.com',
      startDateTime: '2026-04-01T09:00:00',
      endDateTime: '2026-04-01T18:00:00',
      details: true,
      json: true,
    });

    expect(mockOutputAvailability).toHaveBeenCalledWith(
      [
        expect.objectContaining({
          scheduleId: 'user1@contoso.com',
          availabilityView: '0222',
          startDateTime: '2026-04-01T09:00:00',
          endDateTime: '2026-04-01T18:00:00',
          timeZone: 'China Standard Time',
          intervalMinutes: 30,
          slots: expect.any(Array),
          segments: expect.any(Array),
          freeSegments: expect.any(Array),
        }),
      ],
      {
        json: true,
        details: true,
        startDateTime: '2026-04-01T09:00:00',
        endDateTime: '2026-04-01T18:00:00',
        timeZone: 'China Standard Time',
        intervalMinutes: 30,
      }
    );
  });

  it('should call handleError on API failure', async () => {
    const apiError = new Error('Graph API error');
    mockCalendarGetSchedule.mockRejectedValue(apiError);

    await getAvailability({
      users: 'user1@contoso.com',
      startDateTime: '2026-04-01T09:00:00',
      endDateTime: '2026-04-01T18:00:00',
      json: true,
    });

    expect(mockHandleError).toHaveBeenCalledWith(apiError, { json: true });
  });

  it('should call handleError when ensureWorkAccount throws', async () => {
    mockEnsureWorkAccount.mockImplementation(() => {
      throw new Error('Work account required');
    });

    await getAvailability({
      users: 'user1@contoso.com',
      startDateTime: '2026-04-01T09:00:00',
      endDateTime: '2026-04-01T18:00:00',
      json: false,
    });

    expect(mockCalendarGetSchedule).not.toHaveBeenCalled();
    expect(mockHandleError).toHaveBeenCalled();
    expect(mockHandleError.mock.calls[0][0].message).toBe('Work account required');
  });
});

describe('availability utility functions', () => {
  describe('mapStatus', () => {
    it('should map 0 to free', () => { expect(mapStatus('0')).toBe('free'); });
    it('should map 1 to tentative', () => { expect(mapStatus('1')).toBe('tentative'); });
    it('should map 2 to busy', () => { expect(mapStatus('2')).toBe('busy'); });
    it('should map 3 to oof', () => { expect(mapStatus('3')).toBe('oof'); });
    it('should map 4 to workingElsewhere', () => { expect(mapStatus('4')).toBe('workingElsewhere'); });
    it('should map unknown char to unknown', () => { expect(mapStatus('9')).toBe('unknown'); });
  });

  describe('expandSlots', () => {
    it('should expand availabilityView to slots with correct times', () => {
      const slots = expandSlots('002', '2026-04-01T09:00:00', 30);
      expect(slots).toEqual([
        { start: '2026-04-01T09:00:00', end: '2026-04-01T09:30:00', status: 'free' },
        { start: '2026-04-01T09:30:00', end: '2026-04-01T10:00:00', status: 'free' },
        { start: '2026-04-01T10:00:00', end: '2026-04-01T10:30:00', status: 'busy' },
      ]);
    });

    it('should handle hour rollover correctly', () => {
      const slots = expandSlots('00', '2026-04-01T23:30:00', 30);
      expect(slots[0].start).toBe('2026-04-01T23:30:00');
      expect(slots[0].end).toBe('2026-04-02T00:00:00');
      expect(slots[1].start).toBe('2026-04-02T00:00:00');
      expect(slots[1].end).toBe('2026-04-02T00:30:00');
    });

    it('should handle custom interval', () => {
      const slots = expandSlots('02', '2026-04-01T09:00:00', 60);
      expect(slots).toEqual([
        { start: '2026-04-01T09:00:00', end: '2026-04-01T10:00:00', status: 'free' },
        { start: '2026-04-01T10:00:00', end: '2026-04-01T11:00:00', status: 'busy' },
      ]);
    });

    it('should return empty array for empty view', () => {
      expect(expandSlots('', '2026-04-01T09:00:00', 30)).toEqual([]);
    });
  });

  describe('mergeSegments', () => {
    it('should merge consecutive same-status slots', () => {
      const slots = [
        { start: '2026-04-01T09:00:00', end: '2026-04-01T09:30:00', status: 'free' },
        { start: '2026-04-01T09:30:00', end: '2026-04-01T10:00:00', status: 'free' },
        { start: '2026-04-01T10:00:00', end: '2026-04-01T10:30:00', status: 'busy' },
      ];
      const segments = mergeSegments(slots);
      expect(segments).toEqual([
        { start: '2026-04-01T09:00:00', end: '2026-04-01T10:00:00', status: 'free', durationMinutes: 60 },
        { start: '2026-04-01T10:00:00', end: '2026-04-01T10:30:00', status: 'busy', durationMinutes: 30 },
      ]);
    });

    it('should not merge different statuses', () => {
      const slots = [
        { start: '2026-04-01T09:00:00', end: '2026-04-01T09:30:00', status: 'free' },
        { start: '2026-04-01T09:30:00', end: '2026-04-01T10:00:00', status: 'busy' },
      ];
      const segments = mergeSegments(slots);
      expect(segments).toHaveLength(2);
    });

    it('should handle single slot', () => {
      const segments = mergeSegments([
        { start: '2026-04-01T09:00:00', end: '2026-04-01T09:30:00', status: 'free' },
      ]);
      expect(segments).toEqual([
        { start: '2026-04-01T09:00:00', end: '2026-04-01T09:30:00', status: 'free', durationMinutes: 30 },
      ]);
    });

    it('should handle empty array', () => {
      expect(mergeSegments([])).toEqual([]);
    });
  });

  describe('filterFreeSegments', () => {
    it('should filter only free segments', () => {
      const segments = [
        { start: 'a', end: 'b', status: 'free', durationMinutes: 60 },
        { start: 'b', end: 'c', status: 'busy', durationMinutes: 30 },
        { start: 'c', end: 'd', status: 'free', durationMinutes: 30 },
      ];
      const free = filterFreeSegments(segments);
      expect(free).toEqual([
        { start: 'a', end: 'b', durationMinutes: 60 },
        { start: 'c', end: 'd', durationMinutes: 30 },
      ]);
    });

    it('should return empty for no free segments', () => {
      expect(filterFreeSegments([
        { start: 'a', end: 'b', status: 'busy', durationMinutes: 30 },
      ])).toEqual([]);
    });
  });

  describe('enrichScheduleResult', () => {
    it('should add derived fields to schedule result', () => {
      const raw = {
        scheduleId: 'alice@contoso.com',
        availabilityView: '002',
        scheduleItems: [],
        workingHours: {},
      };
      const ctx = {
        startDateTime: '2026-04-01T09:00:00',
        endDateTime: '2026-04-01T10:30:00',
        timeZone: 'China Standard Time',
        intervalMinutes: 30,
      };
      const enriched = enrichScheduleResult(raw, ctx);

      expect(enriched.scheduleId).toBe('alice@contoso.com');
      expect(enriched.startDateTime).toBe('2026-04-01T09:00:00');
      expect(enriched.endDateTime).toBe('2026-04-01T10:30:00');
      expect(enriched.timeZone).toBe('China Standard Time');
      expect(enriched.intervalMinutes).toBe(30);
      expect(enriched.availabilityView).toBe('002');
      expect(enriched.slots).toHaveLength(3);
      expect(enriched.segments).toBeDefined();
      expect(enriched.freeSegments).toBeDefined();
      expect(enriched.scheduleItems).toEqual([]);
      expect(enriched.workingHours).toEqual({});
    });
  });
});

describe('outputAvailability text output', () => {
  let logOutput;

  beforeEach(() => {
    logOutput = [];
    vi.spyOn(console, 'log').mockImplementation((...args) => {
      logOutput.push(args.join(' '));
    });
  });

  it('single user: should show real time segments, not Slot numbers', () => {
    const data = [
      {
        scheduleId: 'alice@contoso.com',
        availabilityView: '0021',
        startDateTime: '2026-04-01T09:00:00',
        endDateTime: '2026-04-01T11:00:00',
        timeZone: 'China Standard Time',
        intervalMinutes: 30,
        slots: [
          { start: '2026-04-01T09:00:00', end: '2026-04-01T09:30:00', status: 'free' },
          { start: '2026-04-01T09:30:00', end: '2026-04-01T10:00:00', status: 'free' },
          { start: '2026-04-01T10:00:00', end: '2026-04-01T10:30:00', status: 'busy' },
          { start: '2026-04-01T10:30:00', end: '2026-04-01T11:00:00', status: 'tentative' },
        ],
        segments: [
          { start: '2026-04-01T09:00:00', end: '2026-04-01T10:00:00', status: 'free', durationMinutes: 60 },
          { start: '2026-04-01T10:00:00', end: '2026-04-01T10:30:00', status: 'busy', durationMinutes: 30 },
          { start: '2026-04-01T10:30:00', end: '2026-04-01T11:00:00', status: 'tentative', durationMinutes: 30 },
        ],
        freeSegments: [{ start: '2026-04-01T09:00:00', end: '2026-04-01T10:00:00', durationMinutes: 60 }],
      },
    ];

    realOutputAvailability(data, { json: false });

    const joined = logOutput.join('\n');
    expect(joined).not.toContain('Slot');
    expect(joined).toContain('09:00');
    expect(joined).toContain('→');
  });

  it('multi user: should show time labels as row headers', () => {
    const data = [
      enrichScheduleResult(
        { scheduleId: 'alice@contoso.com', availabilityView: '02', scheduleItems: [], workingHours: {} },
        {
          startDateTime: '2026-04-01T09:00:00',
          endDateTime: '2026-04-01T10:00:00',
          timeZone: 'China Standard Time',
          intervalMinutes: 30,
        }
      ),
      enrichScheduleResult(
        { scheduleId: 'bob@contoso.com', availabilityView: '20', scheduleItems: [], workingHours: {} },
        {
          startDateTime: '2026-04-01T09:00:00',
          endDateTime: '2026-04-01T10:00:00',
          timeZone: 'China Standard Time',
          intervalMinutes: 30,
        }
      ),
    ];

    realOutputAvailability(data, { json: false });

    const joined = logOutput.join('\n');
    expect(joined).toContain('Time');
    expect(joined).toContain('09:00 → 09:30');
    expect(joined).not.toContain('Slot');
  });

  it('JSON output should include derived fields', () => {
    const data = [
      enrichScheduleResult(
        { scheduleId: 'alice@contoso.com', availabilityView: '002', scheduleItems: [], workingHours: {} },
        {
          startDateTime: '2026-04-01T09:00:00',
          endDateTime: '2026-04-01T10:30:00',
          timeZone: 'China Standard Time',
          intervalMinutes: 30,
        }
      ),
    ];

    realOutputAvailability(data, { json: true });

    const parsed = JSON.parse(logOutput[0]);
    expect(parsed[0]).toHaveProperty('slots');
    expect(parsed[0]).toHaveProperty('segments');
    expect(parsed[0]).toHaveProperty('freeSegments');
  });
});
