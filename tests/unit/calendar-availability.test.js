import { describe, it, expect, vi, beforeEach } from 'vitest';

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

vi.mock('../../src/utils/output.js', () => ({
  outputAvailability: (...args) => mockOutputAvailability(...args),
}));

const mockHandleError = vi.fn();
vi.mock('../../src/utils/error.js', () => ({
  handleError: (...args) => mockHandleError(...args),
}));

const mockEnsureWorkAccount = vi.fn();
vi.mock('../../src/utils/account.js', () => ({
  ensureWorkAccount: (...args) => mockEnsureWorkAccount(...args),
}));

import { getAvailability } from '../../src/commands/calendar.js';

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
        timezone: undefined,
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
        timezone: undefined,
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
        timezone: undefined,
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

    expect(mockOutputAvailability).toHaveBeenCalledWith(result, { json: true, details: true });
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
