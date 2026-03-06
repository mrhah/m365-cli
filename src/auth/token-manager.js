import { readFileSync, writeFileSync, mkdirSync, chmodSync, unlinkSync } from 'fs';
import { dirname, join } from 'path';
import config from '../utils/config.js';
import { AuthError, TokenExpiredError } from '../utils/error.js';
import { deviceCodeFlow } from './device-flow.js';

/**
 * Token Manager
 * Handles token storage, refresh, and validation
 */

// Personal Microsoft account tenant ID
const MSA_TENANT_ID = '9188040d-6c67-4c5b-b112-36a304b66dad';

/**
 * Detect account type from access token JWT
 * @param {string} accessToken
 * @param {string} [hint] - Fallback when token is opaque (personal MSA tokens aren't JWTs)
 */
function detectAccountType(accessToken, hint) {
  try {
    const parts = accessToken.split('.');
    if (parts.length !== 3) {
      // Opaque token (personal Microsoft accounts) — use hint or default
      return hint || 'work';
    }
    const payload = JSON.parse(
      Buffer.from(parts[1], 'base64url').toString()
    );
    return payload.tid === MSA_TENANT_ID ? 'personal' : 'work';
  } catch {
    return hint || 'work';
  }
}

/**
 * Get current account type from stored credentials
 * Returns 'work' by default (backward compatible)
 */
export function getAccountType() {
  const creds = loadCreds();
  return creds?.accountType || 'work';
}

/**
 * Get default scopes based on account type
 */
export function getDefaultScopes(accountType = 'work') {
  return accountType === 'personal'
    ? config.get('personalScopes')
    : config.get('workScopes');
}

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
  const creds = loadCreds();
  const accountType = creds?.accountType || 'work';
  // Personal account tokens are issued by /common, so refresh must use /common too
  const tenantId = accountType === 'personal' ? 'common' : config.get('tenantId');
  const clientId = config.get('clientId');
  const scopes = getDefaultScopes(accountType).join(' ');
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
      accountType: creds.accountType || 'work',
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
 * @param {string} [options.addScopes] - Comma-separated scopes to add to defaults
 * @param {string} [options.exclude] - Comma-separated scopes to exclude from defaults
 */
export async function login({ scopes, addScopes, exclude, accountType } = {}) {
  // Default account type is 'work'
  const effectiveAccountType = accountType || 'work';

  // Resolve final scope list
  let overrideScopes;
  let effectiveScopes;

  const optionCount = [scopes, addScopes, exclude].filter(Boolean).length;
  if (optionCount > 1) {
    throw new AuthError('Cannot combine --scopes, --add-scopes, and --exclude. Use only one.');
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
  } else if (addScopes) {
    // User wants to add extra scopes on top of defaults
    const additionalList = addScopes.split(',').map(s => {
      s = s.trim();
      if (s === 'offline_access' || s.startsWith('https://')) return s;
      return `${GRAPH_PREFIX}${s}`;
    });
    const defaultScopes = getDefaultScopes(effectiveAccountType);
    overrideScopes = [...new Set([...defaultScopes, ...additionalList])];
    effectiveScopes = overrideScopes;

    const added = additionalList.filter(s => !defaultScopes.includes(s));
    if (added.length > 0) {
      console.log(`ℹ️  Adding scopes: ${added.map(s => s.replace(GRAPH_PREFIX, '')).join(', ')}\n`);
    }
  } else if (exclude) {
    // User wants to exclude specific scopes from defaults
    const excludeList = exclude.split(',').map(s => {
      s = s.trim();
      if (s === 'offline_access' || s.startsWith('https://')) return s;
      return `${GRAPH_PREFIX}${s}`;
    });
    const defaultScopes = getDefaultScopes(effectiveAccountType);
    overrideScopes = defaultScopes.filter(s => !excludeList.includes(s));
    effectiveScopes = overrideScopes;

    const removed = defaultScopes.filter(s => excludeList.includes(s));
    if (removed.length > 0) {
      console.log(`ℹ️  Excluding scopes: ${removed.map(s => s.replace(GRAPH_PREFIX, '')).join(', ')}\n`);
    }
  } else {
    // Default — use all scopes from config
    effectiveScopes = getDefaultScopes(effectiveAccountType);
    // For personal accounts, we must pass personalScopes explicitly
    // because requestDeviceCode defaults to workScopes
    if (effectiveAccountType === 'personal') {
      overrideScopes = effectiveScopes;
    }
  }

  try {
    // Use 'common' tenant for personal accounts — /consumers device codes
    // aren't recognized by the Microsoft verification page, but /common
    // supports both work and personal accounts (user picks in browser)
    const overrideTenant = effectiveAccountType === 'personal' ? 'common' : undefined;
    const flowOptions = {
      ...(overrideScopes ? { overrideScopes } : {}),
      ...(overrideTenant ? { overrideTenant } : {}),
    };
    const result = await deviceCodeFlow(flowOptions);
    
    // Detect actual account type from JWT (pass user hint for opaque MSA tokens)
    const detectedType = detectAccountType(result.accessToken, effectiveAccountType);
    
    const creds = {
      tenantId: config.get('tenantId'),
      clientId: config.get('clientId'),
      accessToken: result.accessToken,
      refreshToken: result.refreshToken,
      expiresAt: Math.floor(Date.now() / 1000) + result.expiresIn,
      grantedScopes: effectiveScopes,
      accountType: detectedType,
    };
    
    saveCreds(creds);
    
    const typeLabel = detectedType === 'personal' ? 'Personal Microsoft Account' : 'Work/School Account';
    console.log('\n✅ Authentication successful!');
    console.log(`   Account type: ${typeLabel}`);
    console.log(`   Credentials saved to: ${config.getCredsPath()}`);
    
    return true;
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
  getAccountType,
  getDefaultScopes,
  login,
  logout,
};
