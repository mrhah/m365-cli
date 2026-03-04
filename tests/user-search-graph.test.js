import { describe, it, expect, vi, beforeEach } from 'vitest';

vi.mock('../src/utils/config.js', () => ({
  default: {
    get: vi.fn((key) => {
      if (key === 'graphApiUrl') return 'https://graph.microsoft.com/v1.0';
      return undefined;
    }),
  },
}));

vi.mock('../src/auth/token-manager.js', () => ({
  getAccessToken: vi.fn(() => 'mock-token'),
}));

vi.mock('../src/utils/error.js', () => ({
  ApiError: class ApiError extends Error {
    constructor(message, statusCode) {
      super(message);
      this.name = 'ApiError';
      this.statusCode = statusCode;
    }
  },
  parseGraphError: vi.fn((data, status) => {
    const err = new Error(data?.error?.message || `API Error (${status})`);
    err.name = 'ApiError';
    err.statusCode = status;
    return err;
  }),
}));

const mockFetch = vi.fn();
vi.stubGlobal('fetch', mockFetch);

import graphClient from '../src/graph/client.js';

describe('GraphClient user search', () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it('merges users and contacts and de-duplicates by email', async () => {
    mockFetch
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: async () => ({
          value: [
            { id: 'u1', displayName: 'Jerry A', mail: 'jerry@contoso.com', department: 'Sales', jobTitle: 'Manager' },
            { id: 'u2', displayName: 'Jerry B', userPrincipalName: 'jerryb@contoso.com' },
          ],
        }),
      })
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: async () => ({
          value: [
            { id: 'c1', displayName: 'Jerry Personal', emailAddresses: [{ address: 'jerry.personal@example.com' }], companyName: 'Freelance' },
            { id: 'c2', displayName: 'Jerry Duplicate', emailAddresses: [{ address: 'jerry@contoso.com' }] },
          ],
        }),
      });

    const results = await graphClient.user.search('Jerry', { top: 5 });

    expect(mockFetch).toHaveBeenCalledTimes(2);
    expect(results).toHaveLength(3);
    expect(results.map(r => r.email)).toEqual([
      'jerry@contoso.com',
      'jerryb@contoso.com',
      'jerry.personal@example.com',
    ]);
    expect(results[0].source).toBe('organization');
    expect(results[2].source).toBe('contact');
  });

  it('returns available results when one source fails', async () => {
    mockFetch
      .mockResolvedValueOnce({
        ok: false,
        status: 403,
        json: async () => ({ error: { message: 'Forbidden' } }),
      })
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: async () => ({
          value: [
            { id: 'c1', displayName: 'Alice C', emailAddresses: [{ address: 'alice@example.com' }] },
          ],
        }),
      });

    const results = await graphClient.user.search('Alice', { top: 3 });

    expect(results).toEqual([
      {
        id: 'c1',
        name: 'Alice C',
        email: 'alice@example.com',
        department: '',
        jobTitle: '',
        source: 'contact',
      },
    ]);
  });
});
