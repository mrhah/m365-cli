import { readFileSync, writeFileSync, mkdirSync, chmodSync, unlinkSync } from 'fs';
import { dirname, join } from 'path';
import config from '../utils/config.js';
import { AuthError, TokenExpiredError } from '../utils/error.js';
import { deviceCodeFlow } from './device-flow.js';

/**
 * Token Manager
 * Handles token storage, refresh, and validation
 */

/**
 * Load credentials from file
 * Supports migration from old token file
 */
export function loadCreds() {
  const credsPath = config.getCredsPath();
  const oldTokenPath = join(dirname(credsPath), '../.m365-token.json');
  
  // Try to load from new creds file
  try {
    const data = readFileSync(credsPath, 'utf-8');
    const creds = JSON.parse(data);
    
    // Validate it's a proper JSON object
    if (creds && typeof creds === 'object' && creds.accessToken) {
      return creds;
    }
  } catch (error) {
    // File doesn't exist or not valid JSON, continue to migration
  }
  
  // Try to migrate from old token file
  try {
    const oldData = readFileSync(oldTokenPath, 'utf-8');
    const oldToken = JSON.parse(oldData);
    
    if (oldToken && oldToken.access_token) {
      // Migrate to new format
      const creds = {
        tenantId: config.get('tenantId'),
        clientId: config.get('clientId'),
        accessToken: oldToken.access_token,
        refreshToken: oldToken.refresh_token,
        expiresAt: oldToken.expires_at,
      };
      
      // Save to new location
      saveCreds(creds);
      
      console.log('ℹ️  Migrated token from old format.');
      
      return creds;
    }
  } catch (error) {
    // Old file doesn't exist either
  }
  
  return null; // No credentials found
}

/**
 * Save credentials to file
 */
export function saveCreds(creds) {
  const credsPath = config.getCredsPath();
  
  try {
    // Ensure directory exists
    mkdirSync(dirname(credsPath), { recursive: true });
    
    // Write credentials
    writeFileSync(credsPath, JSON.stringify(creds, null, 2), 'utf-8');
    
    // Set permissions to 600 (user read/write only)
    chmodSync(credsPath, 0o600);
  } catch (error) {
    throw new AuthError(`Failed to save credentials: ${error.message}`);
  }
}

/**
 * Check if token is expired
 */
export function isTokenExpired(creds) {
  if (!creds || !creds.expiresAt) {
    return true;
  }
  
  const now = Math.floor(Date.now() / 1000);
  const buffer = config.get('tokenRefreshBuffer') || 60;
  
  return creds.expiresAt <= (now + buffer);
}

/**
 * Refresh access token using refresh token
 */
export async function refreshToken(refreshToken) {
  const tenantId = config.get('tenantId');
  const clientId = config.get('clientId');
  const scopes = config.get('scopes').join(' ');
  const authUrl = config.get('authUrl');
  
  const url = `${authUrl}/${tenantId}/oauth2/v2.0/token`;
  
  try {
    const response = await fetch(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
      },
      body: new URLSearchParams({
        client_id: clientId,
        grant_type: 'refresh_token',
        refresh_token: refreshToken,
        scope: scopes,
      }),
    });
    
    if (!response.ok) {
      const error = await response.json();
      throw new AuthError(
        error.error_description || 'Token refresh failed',
        error
      );
    }
    
    const data = await response.json();
    
    return {
      accessToken: data.access_token,
      refreshToken: data.refresh_token || refreshToken,
      expiresIn: data.expires_in || 3600,
    };
  } catch (error) {
    if (error instanceof AuthError) {
      throw error;
    }
    throw new AuthError(`Failed to refresh token: ${error.message}`);
  }
}

/**
 * Get valid access token (auto-refresh if needed)
 */
export async function getAccessToken() {
  const creds = loadCreds();
  
  if (!creds || !creds.accessToken) {
    throw new AuthError('Not authenticated. Please run: m365 login');
  }
  
  // Check if token is expired
  if (!isTokenExpired(creds)) {
    return creds.accessToken;
  }
  
  // Try to refresh
  if (!creds.refreshToken) {
    throw new TokenExpiredError();
  }
  
  try {
    const refreshed = await refreshToken(creds.refreshToken);
    
    // Save new credentials (preserve grantedScopes)
    const newCreds = {
      tenantId: config.get('tenantId'),
      clientId: config.get('clientId'),
      accessToken: refreshed.accessToken,
      refreshToken: refreshed.refreshToken,
      expiresAt: Math.floor(Date.now() / 1000) + refreshed.expiresIn,
      grantedScopes: creds.grantedScopes || [],
    };
    
    saveCreds(newCreds);
    
    return refreshed.accessToken;
  } catch (error) {
    throw new TokenExpiredError();
  }
}

/**
 * Perform login (device code flow)
 * @param {Object} [options]
 * @param {string} [options.scopes] - Comma-separated scopes to request (overrides defaults)
 * @param {string} [options.exclude] - Comma-separated scopes to exclude from defaults
 */
export async function login({ scopes, exclude } = {}) {
  // Resolve final scope list
  let overrideScopes;
  let effectiveScopes;

  if (scopes && exclude) {
    throw new AuthError('Cannot use --scopes and --exclude together. Use one or the other.');
  }

  const GRAPH_PREFIX = 'https://graph.microsoft.com/';

  if (scopes) {
    // User specified exact scopes — normalize to full URIs
    overrideScopes = scopes.split(',').map(s => {
      s = s.trim();
      if (s === 'offline_access' || s.startsWith('https://')) return s;
      return `${GRAPH_PREFIX}${s}`;
    });
    effectiveScopes = overrideScopes;
  } else if (exclude) {
    // User wants to exclude specific scopes from defaults
    const excludeList = exclude.split(',').map(s => {
      s = s.trim();
      if (s === 'offline_access' || s.startsWith('https://')) return s;
      return `${GRAPH_PREFIX}${s}`;
    });
    const defaultScopes = config.get('scopes');
    overrideScopes = defaultScopes.filter(s => !excludeList.includes(s));
    effectiveScopes = overrideScopes;

    const removed = defaultScopes.filter(s => excludeList.includes(s));
    if (removed.length > 0) {
      console.log(`ℹ️  Excluding scopes: ${removed.map(s => s.replace(GRAPH_PREFIX, '')).join(', ')}\n`);
    }
  } else {
    // Default — use all scopes from config
    effectiveScopes = config.get('scopes');
  }

  try {
    const flowOptions = overrideScopes ? { overrideScopes } : {};
    const result = await deviceCodeFlow(flowOptions);
    
    const creds = {
      tenantId: config.get('tenantId'),
      clientId: config.get('clientId'),
      accessToken: result.accessToken,
      refreshToken: result.refreshToken,
      expiresAt: Math.floor(Date.now() / 1000) + result.expiresIn,
      grantedScopes: effectiveScopes,
    };
    
    saveCreds(creds);
    
    console.log('\n✅ Authentication successful!');
    console.log(`   Credentials saved to: ${config.getCredsPath()}`);
    
    return true;
  } catch (error) {
    throw error;
  }
}

/**
 * Perform login with additional scopes (incremental consent)
 * Re-authenticates with device code flow including extra scopes
 * @param {string[]} additionalScopes - Extra scopes to request
 * @returns {Promise<string>} New access token
 */
export async function loginWithScopes(additionalScopes = []) {
  try {
    const result = await deviceCodeFlow({ additionalScopes });
    
    // Merge previously granted scopes with new ones
    const existingCreds = loadCreds();
    const previousScopes = existingCreds?.grantedScopes || config.get('scopes');
    const allScopes = [...new Set([...previousScopes, ...additionalScopes])];
    
    const creds = {
      tenantId: config.get('tenantId'),
      clientId: config.get('clientId'),
      accessToken: result.accessToken,
      refreshToken: result.refreshToken,
      expiresAt: Math.floor(Date.now() / 1000) + result.expiresIn,
      grantedScopes: allScopes,
    };
    
    saveCreds(creds);
    
    console.log('\n✅ Additional permissions granted!');
    
    return result.accessToken;
  } catch (error) {
    throw error;
  }
}

/**
 * Logout (clear credentials)
 */
export function logout() {
  const credsPath = config.getCredsPath();
  
  try {
    unlinkSync(credsPath);
    console.log('✅ Logged out successfully.');
    return true;
  } catch (error) {
    if (error.code === 'ENOENT') {
      console.log('ℹ️  No credentials found.');
      return true;
    }
    throw new AuthError(`Failed to logout: ${error.message}`);
  }
}

export default {
  loadCreds,
  saveCreds,
  isTokenExpired,
  refreshToken,
  getAccessToken,
  login,
  loginWithScopes,
  logout,
};
