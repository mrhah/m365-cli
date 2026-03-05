import config from '../utils/config.js';
import { AuthError } from '../utils/error.js';

/**
 * Device Code Flow authentication
 * https://learn.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-device-code
 */

/**
 * Request device code from Microsoft
 * @param {string[]} [additionalScopes] - Extra scopes to request beyond default
 */
export async function requestDeviceCode(additionalScopes = []) {
  const tenantId = config.get('tenantId');
  const clientId = config.get('clientId');
  const defaultScopes = config.get('scopes');
  const allScopes = [...new Set([...defaultScopes, ...additionalScopes])];
  const scopes = allScopes.join(' ');
  const authUrl = config.get('authUrl');
  
  const url = `${authUrl}/${tenantId}/oauth2/v2.0/devicecode`;
  
  const response = await fetch(url, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/x-www-form-urlencoded',
    },
    body: new URLSearchParams({
      client_id: clientId,
      scope: scopes,
    }),
  });
  
  if (!response.ok) {
    throw new AuthError('Failed to request device code', await response.json());
  }
  
  const data = await response.json();
  
  return {
    deviceCode: data.device_code,
    userCode: data.user_code,
    verificationUri: data.verification_uri,
    expiresIn: data.expires_in || 900,
    interval: data.interval || 5,
    message: data.message,
  };
}

/**
 * Poll for access token
 */
export async function pollForToken(deviceCode) {
  const tenantId = config.get('tenantId');
  const clientId = config.get('clientId');
  const authUrl = config.get('authUrl');
  
  const url = `${authUrl}/${tenantId}/oauth2/v2.0/token`;
  
  const response = await fetch(url, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/x-www-form-urlencoded',
    },
    body: new URLSearchParams({
      client_id: clientId,
      grant_type: 'urn:ietf:params:oauth:grant-type:device_code',
      device_code: deviceCode,
    }),
  });
  
  const data = await response.json();
  
  // Check for errors
  if (data.error) {
    if (data.error === 'authorization_pending') {
      return { pending: true };
    }
    
    if (data.error === 'slow_down') {
      return { slowDown: true };
    }
    
    throw new AuthError(
      data.error_description || data.error,
      { error: data.error }
    );
  }
  
  // Success - return token data
  return {
    success: true,
    accessToken: data.access_token,
    refreshToken: data.refresh_token,
    expiresIn: data.expires_in || 3600,
  };
}

/**
 * Full device code flow
 * @param {string[]} [additionalScopes] - Extra scopes for incremental consent
 */
export async function deviceCodeFlow(additionalScopes = []) {
  // Step 1: Request device code
  if (additionalScopes.length > 0) {
    console.log('🔐 Additional permissions required. Starting re-authentication...\n');
  } else {
    console.log('🔐 Starting authentication...\n');
  }
  const deviceCodeData = await requestDeviceCode(additionalScopes);
  
  // Step 2: Show user instructions
  console.log('━'.repeat(60));
  console.log('📱 Please authenticate:');
  console.log('');
  console.log(`   1. Open: ${deviceCodeData.verificationUri}`);
  console.log(`   2. Enter code: ${deviceCodeData.userCode}`);
  console.log('');
  console.log('━'.repeat(60));
  console.log('\n⏳ Waiting for authentication...\n');
  
  // Step 3: Poll for token
  const startTime = Date.now();
  const expiresAt = startTime + deviceCodeData.expiresIn * 1000;
  let interval = deviceCodeData.interval * 1000;
  
  while (Date.now() < expiresAt) {
    await new Promise(resolve => setTimeout(resolve, interval));
    
    try {
      const result = await pollForToken(deviceCodeData.deviceCode);
      
      if (result.success) {
        return {
          accessToken: result.accessToken,
          refreshToken: result.refreshToken,
          expiresIn: result.expiresIn,
        };
      }
      
      if (result.slowDown) {
        // Increase polling interval
        interval += 1000;
      }
      
      // Keep waiting if pending
      if (result.pending) {
        continue;
      }
    } catch (error) {
      throw error;
    }
  }
  
  throw new AuthError('Authentication timed out. Please try again.');
}

export default {
  requestDeviceCode,
  pollForToken,
  deviceCodeFlow,
};
