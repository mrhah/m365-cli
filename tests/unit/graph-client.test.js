import { describe, it, expect, vi, beforeEach } from 'vitest';

// Mock token manager
vi.mock('../../src/auth/token-manager.js', () => ({
  getAccessToken: vi.fn().mockResolvedValue('mock-token'),
}));

// Mock config
vi.mock('../../src/utils/config.js', () => ({
  default: {
    get: vi.fn((key) => {
      const config = {
        graphApiUrl: 'https://graph.microsoft.com/v1.0',
      };
      return config[key];
    }),
  },
}));

// Mock fetch globally
const mockFetch = vi.fn();
vi.stubGlobal('fetch', mockFetch);

// Import AFTER mocks — vitest hoists vi.mock
import graphClient from '../../src/graph/client.js';
import { InsufficientPrivilegesError, ApiError } from '../../src/utils/error.js';

describe('GraphClient', () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  describe('request error handling', () => {
    it('should throw InsufficientPrivilegesError on 403 AccessDenied without retrying', async () => {
      mockFetch.mockResolvedValueOnce({
        ok: false,
        status: 403,
        json: () => Promise.resolve({
          error: {
            code: 'Authorization_RequestDenied',
            message: 'Insufficient privileges to complete the operation.',
          },
        }),
      });

      await expect(graphClient.get('/sites/test-site/drive')).rejects.toBeInstanceOf(InsufficientPrivilegesError);

      // Should NOT have made a second fetch call (no retry)
      expect(mockFetch).toHaveBeenCalledTimes(1);
    });

    it('should throw ApiError on non-403 errors', async () => {
      mockFetch.mockResolvedValueOnce({
        ok: false,
        status: 404,
        json: () => Promise.resolve({
          error: {
            code: 'itemNotFound',
            message: 'The resource could not be found.',
          },
        }),
      });

      await expect(graphClient.get('/me/messages/nonexistent')).rejects.toBeInstanceOf(ApiError);
      expect(mockFetch).toHaveBeenCalledTimes(1);
    });
  });
});
