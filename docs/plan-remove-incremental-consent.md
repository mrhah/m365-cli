# Implementation Plan: Remove Incremental Consent + Add `--add-scopes`

## Overview

Remove the automatic incremental consent mechanism (auto-retry on 403 with Device Code Flow) and replace it with:
1. Clear error messages suggesting `m365 login --add-scopes <scope>` when permissions are insufficient
2. A new `--add-scopes` CLI option that **appends** scopes to defaults (vs `--scopes` which **replaces** all defaults)

**Design Principle**: All permissions are handled the same way — no special SharePoint auto-re-auth flow. Insufficient permissions → clear error message with actionable suggestion.

---

## Part A: Remove Incremental Consent

### A1. `config/default.json` — Remove `extraScopes`

**File**: `config/default.json` (lines 15-19)

**Remove**:
```json
"extraScopes": {
  "sharepoint": [
    "https://graph.microsoft.com/Sites.ReadWrite.All"
  ]
},
```

**After** (lines 14-15 become adjacent):
```json
    "https://graph.microsoft.com/MailboxSettings.Read"
  ],
  "graphApiUrl": "https://graph.microsoft.com/v1.0",
```

---

### A2. `src/utils/config.js` — Remove `getExtraScopes()`

**File**: `src/utils/config.js`

**Remove function** (lines 48-58):
```js
/**
 * Get extra scopes for a specific feature (e.g., 'sharepoint')
 * Returns an array of scope strings, or empty array if not defined
 */
export function getExtraScopes(feature) {
  const extraScopes = defaultConfig.extraScopes;
  if (!extraScopes || !extraScopes[feature]) {
    return [];
  }
  return extraScopes[feature];
}
```

**Remove from default export** (line 74):
```js
  getExtraScopes,  // DELETE this line
```

---

### A3. `src/graph/client.js` — Remove auto-retry and feature detection

**File**: `src/graph/client.js`

#### A3a. Remove imports (lines 1-2)

**Before**:
```js
import { getAccessToken, loginWithScopes } from '../auth/token-manager.js';
import config, { getExtraScopes } from '../utils/config.js';
```

**After**:
```js
import { getAccessToken } from '../auth/token-manager.js';
import config from '../utils/config.js';
```

#### A3b. Remove `_detectFeature()` method (lines 16-28)

**Remove entire block**:
```js
/**
 * Detect which feature an endpoint belongs to, for incremental consent.
 * Maps API endpoint patterns to feature names defined in config.extraScopes.
 * @param {string} endpoint - Graph API endpoint path
 * @returns {string|null} Feature name (e.g., 'sharepoint') or null
 */
_detectFeature(endpoint) {
  // SharePoint endpoints: /sites/..., /search/query
  if (endpoint.startsWith('/sites') || endpoint === '/search/query') {
    return 'sharepoint';
  }
  return null;
}
```

#### A3c. Remove `_retried` parameter from `request()` (line 89)

**Before** (line 89):
```js
      _retried = false,
```

**After**: Remove this line entirely.

#### A3d. Replace auto-retry catch block (lines 137-149)

**Before**:
```js
    } catch (error) {
      // Incremental consent: on InsufficientPrivilegesError, try to acquire additional scopes
      if (error instanceof InsufficientPrivilegesError && !_retried) {
        const feature = this._detectFeature(endpoint);
        if (feature) {
          const extraScopes = getExtraScopes(feature);
          if (extraScopes.length > 0) {
            await loginWithScopes(extraScopes);
            // Retry the original request once with _retried flag to prevent loops
            return this.request(endpoint, { ...options, _retried: true });
          }
        }
      }

      if (error instanceof ApiError) {
```

**After**:
```js
    } catch (error) {
      if (error instanceof ApiError) {
```

#### A3e. Clean up unused import

After removing the auto-retry block, `InsufficientPrivilegesError` is no longer referenced in `client.js`.

**Before** (line 3):
```js
import { ApiError, InsufficientPrivilegesError, parseGraphError } from '../utils/error.js';
```

**After**:
```js
import { ApiError, parseGraphError } from '../utils/error.js';
```

---

### A4. `src/auth/token-manager.js` — Remove `loginWithScopes()`

**File**: `src/auth/token-manager.js`

#### A4a. Remove `loginWithScopes()` function (lines 254-286)

**Remove entire block**:
```js
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
```

#### A4b. Remove from default export (line 314)

**Before**:
```js
  loginWithScopes,
```

**After**: Remove this line.

---

### A5. `src/auth/device-flow.js` — Remove `additionalScopes` parameter

**File**: `src/auth/device-flow.js`

#### A5a. `requestDeviceCode()` — simplify signature and body

**Before** (lines 11-28):
```js
/**
 * Request device code from Microsoft
 * @param {Object} [options]
 * @param {string[]} [options.additionalScopes] - Extra scopes to request beyond default
 * @param {string[]} [options.overrideScopes] - Complete scope list (replaces defaults entirely)
 */
export async function requestDeviceCode({ additionalScopes = [], overrideScopes } = {}) {
  const tenantId = config.get('tenantId');
  const clientId = config.get('clientId');
  const authUrl = config.get('authUrl');

  let allScopes;
  if (overrideScopes) {
    // Complete replacement — user specified exact scopes via --scopes or --exclude
    allScopes = overrideScopes;
  } else {
    // Default behavior — merge default scopes with any additional ones
    const defaultScopes = config.get('scopes');
    allScopes = [...new Set([...defaultScopes, ...additionalScopes])];
  }
```

**After**:
```js
/**
 * Request device code from Microsoft
 * @param {Object} [options]
 * @param {string[]} [options.overrideScopes] - Complete scope list (replaces defaults entirely)
 */
export async function requestDeviceCode({ overrideScopes } = {}) {
  const tenantId = config.get('tenantId');
  const clientId = config.get('clientId');
  const authUrl = config.get('authUrl');

  let allScopes;
  if (overrideScopes) {
    allScopes = overrideScopes;
  } else {
    allScopes = config.get('scopes');
  }
```

#### A5b. `deviceCodeFlow()` — remove `additionalScopes` parameter

**Before** (lines 123-138):
```js
/**
 * Full device code flow
 * @param {Object} [options]
 * @param {string[]} [options.additionalScopes] - Extra scopes for incremental consent
 * @param {string[]} [options.overrideScopes] - Complete scope list (replaces defaults entirely)
 */
export async function deviceCodeFlow({ additionalScopes = [], overrideScopes } = {}) {
  // Step 1: Request device code
  if (overrideScopes) {
    console.log('🔐 Starting authentication with custom scopes...\n');
  } else if (additionalScopes.length > 0) {
    console.log('🔐 Additional permissions required. Starting re-authentication...\n');
  } else {
    console.log('🔐 Starting authentication...\n');
  }
  const deviceCodeData = await requestDeviceCode({ additionalScopes, overrideScopes });
```

**After**:
```js
/**
 * Full device code flow
 * @param {Object} [options]
 * @param {string[]} [options.overrideScopes] - Complete scope list (replaces defaults entirely)
 */
export async function deviceCodeFlow({ overrideScopes } = {}) {
  // Step 1: Request device code
  if (overrideScopes) {
    console.log('🔐 Starting authentication with custom scopes...\n');
  } else {
    console.log('🔐 Starting authentication...\n');
  }
  const deviceCodeData = await requestDeviceCode({ overrideScopes });
```

---

### A6. `src/utils/error.js` — Add `InsufficientPrivilegesError` suggestion in `handleError()`

**File**: `src/utils/error.js`

After the existing `ConsentRequiredError` suggestion block (line 97), add a similar block for `InsufficientPrivilegesError`:

**Add after line 97** (after the `}` closing the consent error block):
```js
    // Show helpful suggestions for insufficient privileges errors
    if (error instanceof InsufficientPrivilegesError || error.code === 'INSUFFICIENT_PRIVILEGES') {
      console.error('');
      console.error('💡 Suggestions:');
      console.error('');
      console.error('   This command requires additional permissions not included in the default scope set.');
      console.error('');
      console.error('   Re-login with the required scope:');
      console.error('      m365 login --add-scopes Sites.ReadWrite.All');
      console.error('');
      console.error('   Or login with a complete custom scope list:');
      console.error('      m365 login --scopes User.Read,Files.ReadWrite,Sites.ReadWrite.All,offline_access');
      console.error('');
    }
```

---

### A7. Tests — Delete incremental consent tests

**File**: `tests/unit/incremental-consent.test.js` (411 lines)

**Action**: DELETE this entire file. All code paths it tests (getExtraScopes, loginWithScopes, _detectFeature, request retry-with-consent) are being removed.

**Keep untouched**:
- `tests/unit/insufficient-privileges.test.js` — tests `parseGraphError` detection of 403/access denied, which we're keeping
- `tests/unit/token-manager.test.js` — basic token tests, no incremental consent references
- `tests/integration/sharepoint.integration.test.js` — no incremental consent references

---

## Part B: Add `--add-scopes` Option

### B1. `bin/m365.js` — Add `--add-scopes` option to login command

**File**: `bin/m365.js`

**Before** (lines 34-35):
```js
  .option('--scopes <scopes>', 'Comma-separated list of scopes to request (overrides defaults)')
  .option('--exclude <scopes>', 'Comma-separated list of scopes to exclude from defaults')
```

**After**:
```js
  .option('--scopes <scopes>', 'Comma-separated list of scopes to request (overrides defaults)')
  .option('--add-scopes <scopes>', 'Comma-separated list of scopes to add to defaults')
  .option('--exclude <scopes>', 'Comma-separated list of scopes to exclude from defaults')
```

**Update action handler** (line 38):

**Before**:
```js
      await login({ scopes: options.scopes, exclude: options.exclude });
```

**After**:
```js
      await login({ scopes: options.scopes, addScopes: options.addScopes, exclude: options.exclude });
```

---

### B2. `src/auth/token-manager.js` — Add `addScopes` to `login()`

**File**: `src/auth/token-manager.js`

#### B2a. Update `login()` signature and validation (lines 185-198)

**Before**:
```js
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
```

**After**:
```js
/**
 * Perform login (device code flow)
 * @param {Object} [options]
 * @param {string} [options.scopes] - Comma-separated scopes to request (overrides defaults)
 * @param {string} [options.addScopes] - Comma-separated scopes to add to defaults
 * @param {string} [options.exclude] - Comma-separated scopes to exclude from defaults
 */
export async function login({ scopes, addScopes, exclude } = {}) {
  // Resolve final scope list
  let overrideScopes;
  let effectiveScopes;

  // Validate mutually exclusive options
  const optionCount = [scopes, addScopes, exclude].filter(Boolean).length;
  if (optionCount > 1) {
    throw new AuthError('Cannot combine --scopes, --add-scopes, and --exclude. Use only one.');
  }
```

#### B2b. Add `addScopes` handling branch (after `scopes` branch, before `exclude` branch)

**Before** (lines 210-228):
```js
  } else if (exclude) {
    // User wants to exclude specific scopes from defaults
```

**After**:
```js
  } else if (addScopes) {
    // User wants to add extra scopes on top of defaults
    const additionalList = addScopes.split(',').map(s => {
      s = s.trim();
      if (s === 'offline_access' || s.startsWith('https://')) return s;
      return `${GRAPH_PREFIX}${s}`;
    });
    const defaultScopes = config.get('scopes');
    overrideScopes = [...new Set([...defaultScopes, ...additionalList])];
    effectiveScopes = overrideScopes;

    const added = additionalList.filter(s => !defaultScopes.includes(s));
    if (added.length > 0) {
      console.log(`ℹ️  Adding scopes: ${added.map(s => s.replace(GRAPH_PREFIX, '')).join(', ')}\n`);
    }
  } else if (exclude) {
    // User wants to exclude specific scopes from defaults
```

---

### B3. `src/auth/token-manager.js` — Remove `loginWithScopes` from import in `bin/m365.js`

No change needed — `bin/m365.js` imports `{ login, logout }`, not `loginWithScopes`.

---

## Part C: Update README

### C1. Remove incremental consent references

Remove any mention of "on-demand permissions", "incremental consent", or "automatic re-authentication" from:
- Permissions section
- Security section
- Troubleshooting section (if any)

### C2. Document `--add-scopes`

Add to the login command documentation:

```
### Login Options

| Option | Description |
|--------|-------------|
| `--scopes <scopes>` | Override all default scopes with specified list (comma-separated) |
| `--add-scopes <scopes>` | Add extra scopes on top of defaults (comma-separated) |
| `--exclude <scopes>` | Remove specified scopes from defaults (comma-separated) |

**Examples:**
```bash
# Default login (all standard scopes)
m365 login

# Add SharePoint permissions to defaults
m365 login --add-scopes Sites.ReadWrite.All

# Login with only specific scopes (replaces all defaults)
m365 login --scopes User.Read,Files.ReadWrite,offline_access

# Login without calendar permissions
m365 login --exclude Calendars.ReadWrite
```

### C3. Update Troubleshooting / Permissions section

Add guidance for 403 errors:

```
### Insufficient Permissions (403 Error)

If you see "Insufficient privileges" when using SharePoint or other features,
you need to add the required permission scope:

```bash
m365 login --add-scopes Sites.ReadWrite.All
```

This re-authenticates with your existing default scopes PLUS the additional scope.
```

---

## Execution Order

1. **A1** — Remove `extraScopes` from config
2. **A2** — Remove `getExtraScopes()` from config.js
3. **A3** — Clean up client.js (imports, _detectFeature, _retried, auto-retry block)
4. **A4** — Remove `loginWithScopes()` from token-manager.js
5. **A5** — Simplify device-flow.js (remove additionalScopes)
6. **A6** — Add InsufficientPrivilegesError suggestion to handleError()
7. **A7** — Delete incremental-consent.test.js
8. **B1** — Add `--add-scopes` to CLI
9. **B2** — Add `addScopes` to `login()` function
10. **C1-C3** — Update README

## Verification

After all changes:
```bash
# Run unit tests
npm test

# Verify CLI help shows --add-scopes
node bin/m365.js login --help

# Verify no references to removed code
grep -r "loginWithScopes\|getExtraScopes\|_detectFeature\|additionalScopes\|extraScopes" src/
```
