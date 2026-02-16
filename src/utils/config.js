import { readFileSync } from 'fs';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';
import { homedir } from 'os';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

// Load default config
const configPath = join(__dirname, '../../config/default.json');
const defaultConfig = JSON.parse(readFileSync(configPath, 'utf-8'));

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
 * Get configuration value
 * Priority: ENV > default config
 */
export function getConfig(key) {
  // Check environment variable
  const envKey = `M365_${key.toUpperCase().replace(/\./g, '_')}`;
  if (process.env[envKey]) {
    return process.env[envKey];
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
    credsPath: getCredsPath(),
  };
}

export default {
  get: getConfig,
  getCredsPath,
  getAll: getAllConfig,
};
