import { readFileSync } from 'fs';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';
import { homedir } from 'os';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

// Load default config
const configPath = join(__dirname, '../../config/default.json');
const defaultConfig = JSON.parse(readFileSync(configPath, 'utf-8'));

// Cloud endpoint keys — these are resolved from the active cloud config
const CLOUD_KEYS = new Set(['graphApiUrl', 'authUrl', 'deviceLoginUrl', 'scopePrefix']);

/**
 * Expand ~ to home directory
 */
function expandHome(filepath) {
  if (filepath.startsWith('~/')) {
    return join(homedir(), filepath.slice(2));
  }
  return filepath;
}

/**
 * Get the active cloud name.
 * Priority: M365_CLOUD env var > saved credentials > config default > 'global'
 */
export function getActiveCloud() {
  // 1. Environment variable
  const envCloud = process.env.M365_CLOUD;
  if (envCloud) {
    const normalized = envCloud.toLowerCase();
    if (defaultConfig.clouds[normalized]) {
      return normalized;
    }
    // Invalid cloud name in env — fall through
  }

  // 2. Saved credentials (read lazily to avoid circular deps at import time)
  try {
    const credsPath = expandHome(defaultConfig.credsPath);
    const creds = JSON.parse(readFileSync(credsPath, 'utf-8'));
    if (creds.cloud && defaultConfig.clouds[creds.cloud]) {
      return creds.cloud;
    }
  } catch {
    // No credentials file or invalid — fall through
  }

  // 3. Config default
  return defaultConfig.cloud || 'global';
}

/**
 * Get the cloud endpoint config for the active cloud.
 */
export function getCloudConfig(cloudName) {
  const cloud = cloudName || getActiveCloud();
  return defaultConfig.clouds[cloud] || defaultConfig.clouds.global;
}

/**
 * Apply scope prefix to bare scope names.
 * Scopes like 'offline_access' or already-prefixed scopes are left as-is.
 */
export function applyScopes(bareScopes, scopePrefix) {
  if (!bareScopes || !Array.isArray(bareScopes)) return bareScopes;
  return bareScopes.map(s => {
    if (s === 'offline_access' || s.startsWith('https://')) return s;
    return `${scopePrefix}${s}`;
  });
}

/**
 * Get configuration value
 * Priority: ENV > cloud config (for cloud keys) > default config
 */
export function getConfig(key) {
  // Convert camelCase to UPPER_SNAKE_CASE for env var lookup
  // e.g., 'clientId' -> 'M365_CLIENT_ID', 'tenantId' -> 'M365_TENANT_ID'
  const snakeKey = key.replace(/([a-z])([A-Z])/g, '$1_$2').replace(/\./g, '_').toUpperCase();
  const envKey = `M365_${snakeKey}`;
  if (process.env[envKey]) {
    return process.env[envKey];
  }

  // Cloud-specific keys resolve from the active cloud config
  if (CLOUD_KEYS.has(key)) {
    const cloudConfig = getCloudConfig();
    return cloudConfig[key];
  }

  // Scope keys — return with prefix applied
  if (key === 'workScopes' || key === 'personalScopes') {
    const bareScopes = defaultConfig[key];
    const cloudConfig = getCloudConfig();
    return applyScopes(bareScopes, cloudConfig.scopePrefix);
  }

  // Return default config
  return defaultConfig[key];
}

/**
 * Get credentials file path
 */
export function getCredsPath() {
  const path = getConfig('credsPath');
  return expandHome(path);
}


/**
 * Get all config as object
 */
export function getAllConfig() {
  return {
    ...defaultConfig,
    ...getCloudConfig(),
    credsPath: getCredsPath(),
    cloud: getActiveCloud(),
  };
}

export default {
  get: getConfig,
  getCredsPath,
  getAll: getAllConfig,
  getActiveCloud,
  getCloudConfig,
  applyScopes,
};
