import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';

// Mock config
vi.mock('../src/utils/config.js', () => ({
  default: {
    get: vi.fn((key) => {
      const config = {
        tenantId: 'common',
        clientId: 'test-client-id',
        credsPath: '~/.m365-cli/credentials.json',
        tokenRefreshBuffer: 60,
        scopes: ['Mail.Read', 'Files.Read'],
      };
      return config[key];
    }),
    getCredsPath: vi.fn(() => '/home/testuser/.m365-cli/credentials.json'),
  },
}));

// Mock error
vi.mock('../src/utils/error.js', () => ({
  AuthError: class AuthError extends Error {
    constructor(message, details) {
      super(message);
      this.name = 'AuthError';
      this.details = details;
    }
  },
}));

// Mock fs module
vi.mock('fs', async () => {
  const actual = await vi.importActual('fs');
  return {
    ...actual,
    readFileSync: vi.fn(),
    writeFileSync: vi.fn(),
    mkdirSync: vi.fn(),
    chmodSync: vi.fn(),
    unlinkSync: vi.fn(),
  };
});

// Mock device-flow
vi.mock('../src/auth/device-flow.js', () => ({
  deviceCodeFlow: vi.fn(),
}));

import { readFileSync, writeFileSync, mkdirSync, chmodSync, unlinkSync } from 'fs';
import { isTokenExpired, loadCreds, saveCreds, logout } from '../../src/auth/token-manager.js';

describe('Token Manager', () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  describe('isTokenExpired', () => {
    it('should return true if creds is null', () => {
      expect(isTokenExpired(null)).toBe(true);
    });

    it('should return true if creds is undefined', () => {
      expect(isTokenExpired(undefined)).toBe(true);
    });

    it('should return true if expiresAt is missing', () => {
      expect(isTokenExpired({ accessToken: 'abc' })).toBe(true);
    });

    it('should return true if token is expired', () => {
      // Set expiresAt to past time
      const pastTime = Math.floor(Date.now() / 1000) - 100;
      expect(isTokenExpired({ expiresAt: pastTime })).toBe(true);
    });

    it('should return true if token expires within buffer (default 60s)', () => {
      // Set expiresAt to within 30 seconds (less than buffer)
      const soonTime = Math.floor(Date.now() / 1000) + 30;
      expect(isTokenExpired({ expiresAt: soonTime })).toBe(true);
    });

    it('should return false if token is valid (more than buffer time)', () => {
      // Set expiresAt to more than 60 seconds from now
      const futureTime = Math.floor(Date.now() / 1000) + 120;
      expect(isTokenExpired({ expiresAt: futureTime })).toBe(false);
    });
  });

  describe('loadCreds', () => {
    beforeEach(() => {
      readFileSync.mockReset();
    });

    it('should return null if credentials file does not exist', () => {
      readFileSync.mockImplementation(() => {
        throw { code: 'ENOENT' };
      });
      
      const result = loadCreds();
      expect(result).toBeNull();
    });

    it('should return null if credentials file is not valid JSON', () => {
      readFileSync.mockReturnValue('not valid json');
      
      const result = loadCreds();
      expect(result).toBeNull();
    });

    it('should return null if credentials file has no accessToken', () => {
      readFileSync.mockReturnValue(JSON.stringify({ refreshToken: 'abc' }));
      
      const result = loadCreds();
      expect(result).toBeNull();
    });

    it('should return credentials if valid', () => {
      const creds = { accessToken: 'test-token', refreshToken: 'refresh-token' };
      readFileSync.mockReturnValue(JSON.stringify(creds));
      
      const result = loadCreds();
      expect(result.accessToken).toBe('test-token');
      expect(result.refreshToken).toBe('refresh-token');
    });
  });

  describe('saveCreds', () => {
    beforeEach(() => {
      mkdirSync.mockReset();
      writeFileSync.mockReset();
      chmodSync.mockReset();
    });

    it('should write credentials to file', () => {
      mkdirSync.mockReturnValue(undefined);
      writeFileSync.mockReturnValue(undefined);
      chmodSync.mockReturnValue(undefined);
      
      const creds = { accessToken: 'test-token' };
      saveCreds(creds);
      
      expect(writeFileSync).toHaveBeenCalledWith(
        '/home/testuser/.m365-cli/credentials.json',
        JSON.stringify(creds, null, 2),
        'utf-8'
      );
    });

    it('should set file permissions to 600', () => {
      mkdirSync.mockReturnValue(undefined);
      writeFileSync.mockReturnValue(undefined);
      chmodSync.mockReturnValue(undefined);
      
      saveCreds({ accessToken: 'test' });
      
      expect(chmodSync).toHaveBeenCalledWith(
        '/home/testuser/.m365-cli/credentials.json',
        0o600
      );
    });
  });

  describe('logout', () => {
    beforeEach(() => {
      unlinkSync.mockReset();
    });

    it('should delete credentials file', () => {
      unlinkSync.mockReturnValue(undefined);
      
      logout();
      
      expect(unlinkSync).toHaveBeenCalledWith('/home/testuser/.m365-cli/credentials.json');
    });

    it('should return true if file does not exist', () => {
      unlinkSync.mockImplementation(() => {
        throw { code: 'ENOENT' };
      });
      
      const result = logout();
      expect(result).toBe(true);
    });
  });
});
