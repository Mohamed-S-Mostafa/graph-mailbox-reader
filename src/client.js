import axios from "axios";
import { createTokenProvider } from "./auth.js";

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

/**
 * Fields returned for message list/detail by default. Keeping this explicit
 * keeps payloads small and predictable; override per call via `select`.
 */
const DEFAULT_MESSAGE_FIELDS = [
  "id",
  "subject",
  "from",
  "toRecipients",
  "receivedDateTime",
  "isRead",
  "hasAttachments",
  "bodyPreview",
  "webLink",
];

/**
 * Create a mailbox reader bound to the signed-in user via delegated auth.
 *
 * SECURITY: every request targets `/me` — the signed-in user's own mailbox.
 * There is intentionally NO mailbox/userId parameter on any method. With a
 * delegated Mail.Read token it is impossible to read another person's mail,
 * so "which mailbox" is decided entirely by who signs in. To read a different
 * person's mailbox, that person signs in (optionally with their own
 * `tokenCachePath`, or a separate app registration / clientId).
 *
 * @param {Object} config
 * @param {string} config.tenantId - Azure AD tenant ID (or "common"/"organizations")
 * @param {string} config.clientId - App registration (public client) ID
 * @param {string} [config.tokenCachePath] - Where to persist the token cache (defaults to ~/.graph-mailbox-reader/<clientId>.token-cache.json)
 * @param {(response: { message: string }) => void} [config.deviceCodeCallback] - Called with sign-in instructions; defaults to console.log
 * @returns {object} mailbox reader
 *
 * @example
 * import { createMailboxReader } from "graph-mailbox-reader";
 *
 * const reader = createMailboxReader({
 *   tenantId: process.env.MICROSOFT_GRAPH_TENANT_ID,
 *   clientId: process.env.MICROSOFT_GRAPH_CLIENT_ID,
 * });
 *
 * await reader.signIn();                         // device-code login (first run only)
 * const me = await reader.getProfile();          // { mail, displayName, ... }
 * const { value } = await reader.listMessages({ top: 10, folder: "inbox" });
 */
export function createMailboxReader({
  tenantId,
  clientId,
  tokenCachePath,
  deviceCodeCallback,
}) {
  const tokens = createTokenProvider({
    tenantId,
    clientId,
    tokenCachePath,
    deviceCodeCallback,
  });

  /**
   * GET a Graph resource with the current access token.
   *
   * @param {string} pathName - path under graph.microsoft.com/v1.0, must start with "/"
   * @param {Record<string, string|number|boolean>} [query] - query params (without the leading $ — added here for OData keys)
   * @returns {Promise<any>}
   */
  async function graphGet(pathName, query) {
    const token = await tokens.getAccessToken();
    const res = await axios.get(`${GRAPH_BASE}${pathName}`, {
      headers: { Authorization: `Bearer ${token}` },
      params: query,
    });
    return res.data;
  }

  /**
   * Translate friendly list options into Graph OData query params.
   *
   * @param {ListMessagesOptions} [options]
   * @returns {Record<string, string|number>}
   */
  function toMessageQuery({ top = 25, skip, search, filter, select, orderBy } = {}) {
    const query = {};
    query.$top = top;
    if (skip) query.$skip = skip;
    // $search cannot be combined with $filter/$orderby in Graph mail queries.
    if (search) {
      query.$search = `"${search}"`;
    } else {
      if (filter) query.$filter = filter;
      if (orderBy) query.$orderby = orderBy;
    }
    query.$select = (select && select.length ? select : DEFAULT_MESSAGE_FIELDS).join(",");
    return query;
  }

  /**
   * Ensure the user is signed in (device-code flow on first run) and return
   * their account info.
   */
  async function signIn() {
    return tokens.signIn();
  }

  /** Sign the user out by clearing the cached account. */
  async function signOut() {
    return tokens.signOut();
  }

  /** The signed-in user's profile (Graph /me). */
  async function getProfile() {
    return graphGet("/me", {
      $select: "id,displayName,mail,userPrincipalName",
    });
  }

  /**
   * List messages from the signed-in user's mailbox.
   *
   * @param {ListMessagesOptions} [options]
   * @returns {Promise<{ value: object[], "@odata.nextLink"?: string }>}
   */
  async function listMessages(options = {}) {
    const { folder } = options;
    const base = folder
      ? `/me/mailFolders/${encodeURIComponent(folder)}/messages`
      : "/me/messages";
    return graphGet(base, toMessageQuery(options));
  }

  /**
   * Get a single message by id from the signed-in user's mailbox.
   *
   * @param {string} id - message id
   * @param {{ select?: string[] }} [options]
   * @returns {Promise<object>}
   */
  async function getMessage(id, { select } = {}) {
    if (!id) throw new Error("getMessage requires a message id");
    return graphGet(`/me/messages/${encodeURIComponent(id)}`, {
      $select: (select && select.length ? select : DEFAULT_MESSAGE_FIELDS).join(","),
    });
  }

  /**
   * List the signed-in user's mail folders.
   *
   * @param {{ top?: number }} [options]
   * @returns {Promise<{ value: object[] }>}
   */
  async function listFolders({ top = 50 } = {}) {
    return graphGet("/me/mailFolders", {
      $top: top,
      // Note: `wellKnownName` exists only on Graph's beta endpoint, not v1.0 —
      // selecting it here returns 400. Match folders by displayName instead.
      $select: "id,displayName,unreadItemCount,totalItemCount",
    });
  }

  return {
    signIn,
    signOut,
    getProfile,
    listMessages,
    getMessage,
    listFolders,
  };
}
