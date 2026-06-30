import fs from "node:fs/promises";
import path from "node:path";
import os from "node:os";
import { PublicClientApplication, LogLevel } from "@azure/msal-node";

/**
 * Delegated scopes requested for mailbox reading.
 *
 * - Mail.Read   → read the signed-in user's own mail
 * - User.Read   → read the signed-in user's profile (/me)
 *
 * `offline_access` is added automatically by MSAL so a refresh token is
 * issued and the user only signs in once. These are DELEGATED scopes: the
 * resulting token can only ever act as the signed-in user, so /me is
 * intrinsically limited to that one person's mailbox.
 */
const DEFAULT_SCOPES = ["Mail.Read", "User.Read"];

/**
 * Build a default per-app token-cache path under the user's home directory.
 * Keyed by clientId so multiple app registrations don't clobber each other.
 *
 * @param {string} clientId
 * @returns {string}
 */
function defaultCachePath(clientId) {
  return path.join(os.homedir(), ".graph-mailbox-reader", `${clientId}.token-cache.json`);
}

/**
 * MSAL cache plugin that persists the token cache (including the refresh
 * token) to disk, so sign-in survives process restarts.
 *
 * @param {string} cacheFilePath
 */
function createFileCachePlugin(cacheFilePath) {
  return {
    async beforeCacheAccess(cacheContext) {
      try {
        const data = await fs.readFile(cacheFilePath, "utf-8");
        cacheContext.tokenCache.deserialize(data);
      } catch (err) {
        // No cache yet (first run) is expected; anything else is a real error.
        if (err.code !== "ENOENT") throw err;
      }
    },
    async afterCacheAccess(cacheContext) {
      if (!cacheContext.cacheHasChanged) return;
      await fs.mkdir(path.dirname(cacheFilePath), { recursive: true });
      await fs.writeFile(cacheFilePath, cacheContext.tokenCache.serialize(), {
        mode: 0o600,
      });
    },
  };
}

/**
 * Default device-code prompt: print Microsoft's instructions to the console.
 *
 * @param {{ message: string }} response
 */
function defaultDeviceCodeCallback(response) {
  // response.message already contains the URL + user code + instructions.
  console.log(response.message);
}

/**
 * Create the delegated-auth token provider.
 *
 * Uses a PublicClientApplication with the OAuth 2.0 device-code flow — no
 * client secret is needed or wanted (public clients must not hold secrets).
 * The Azure app registration must have "Allow public client flows" enabled
 * and the delegated Mail.Read permission granted.
 *
 * @param {Object} config
 * @param {string} config.tenantId - Azure AD tenant ID (or "common"/"organizations")
 * @param {string} config.clientId - App registration (public client) ID
 * @param {string} [config.tokenCachePath] - Where to persist the token cache
 * @param {string[]} [config.scopes] - Override the default delegated scopes
 * @param {(response: { message: string }) => void} [config.deviceCodeCallback] - Called with sign-in instructions
 * @returns {{ getAccessToken: Function, signIn: Function, signOut: Function, getAccount: Function }}
 */
export function createTokenProvider({
  tenantId,
  clientId,
  tokenCachePath,
  scopes = DEFAULT_SCOPES,
  deviceCodeCallback = defaultDeviceCodeCallback,
}) {
  if (!tenantId || !clientId) {
    throw new Error("Missing required config: tenantId, clientId");
  }

  const cacheFilePath = tokenCachePath || defaultCachePath(clientId);

  const pca = new PublicClientApplication({
    auth: {
      clientId,
      authority: `https://login.microsoftonline.com/${tenantId}`,
    },
    cache: {
      cachePlugin: createFileCachePlugin(cacheFilePath),
    },
    system: {
      loggerOptions: {
        loggerCallback() {},
        logLevel: LogLevel.Error,
      },
    },
  });

  /** Return the first cached account, or null if no one is signed in. */
  async function getAccount() {
    const accounts = await pca.getTokenCache().getAllAccounts();
    return accounts[0] ?? null;
  }

  /**
   * Acquire an access token: silently from cache (refreshing as needed),
   * falling back to an interactive device-code sign-in.
   *
   * @param {{ forceDeviceCode?: boolean }} [opts]
   * @returns {Promise<string>} access token
   */
  async function getAccessToken({ forceDeviceCode = false } = {}) {
    if (!forceDeviceCode) {
      const account = await getAccount();
      if (account) {
        try {
          const result = await pca.acquireTokenSilent({ account, scopes });
          if (result?.accessToken) return result.accessToken;
        } catch {
          // Silent refresh failed (e.g. expired refresh token) — fall through
          // to interactive device-code sign-in below.
        }
      }
    }

    const result = await pca.acquireTokenByDeviceCode({
      scopes,
      deviceCodeCallback,
    });
    if (!result?.accessToken) {
      throw new Error("Device-code sign-in did not return an access token");
    }
    return result.accessToken;
  }

  /**
   * Ensure the user is signed in (triggers device-code flow if needed) and
   * return their account.
   *
   * @returns {Promise<import("@azure/msal-node").AccountInfo>}
   */
  async function signIn() {
    await getAccessToken();
    return getAccount();
  }

  /** Remove the cached account, requiring a fresh sign-in next time. */
  async function signOut() {
    const account = await getAccount();
    if (account) {
      await pca.getTokenCache().removeAccount(account);
    }
  }

  return { getAccessToken, signIn, signOut, getAccount };
}
