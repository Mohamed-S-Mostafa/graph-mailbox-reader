export interface MailboxReaderConfig {
  /** Azure AD tenant ID (or "common"/"organizations") */
  tenantId: string;
  /** App registration (public client) ID */
  clientId: string;
  /**
   * Where to persist the MSAL token cache. Defaults to
   * ~/.graph-mailbox-reader/<clientId>.token-cache.json. Use a distinct path
   * per user if several sign in on the same machine.
   */
  tokenCachePath?: string;
  /**
   * Called with the device-code sign-in instructions. Defaults to printing
   * `response.message` to the console.
   */
  deviceCodeCallback?: (response: { message: string }) => void;
}

export interface ListMessagesOptions {
  /** Well-known folder name ("inbox", "sentitems", ...) or folder id. Omit for all messages. */
  folder?: string;
  /** Page size ($top). Default 25. */
  top?: number;
  /** Number of items to skip ($skip). */
  skip?: number;
  /** Free-text search ($search). Cannot be combined with filter/orderBy. */
  search?: string;
  /** OData $filter expression, e.g. "isRead eq false". Ignored when `search` is set. */
  filter?: string;
  /** Fields to return ($select). Defaults to a compact message projection. */
  select?: string[];
  /** OData $orderby, e.g. "receivedDateTime desc". Ignored when `search` is set. */
  orderBy?: string;
}

export interface GraphAccount {
  homeAccountId: string;
  username: string;
  name?: string;
  [key: string]: unknown;
}

export interface MailboxReader {
  /** Sign in (device-code flow on first run) and return the account. */
  signIn(): Promise<GraphAccount | null>;
  /** Clear the cached account, forcing a fresh sign-in next time. */
  signOut(): Promise<void>;
  /** The signed-in user's profile (Graph /me). */
  getProfile(): Promise<{
    id: string;
    displayName: string;
    mail: string | null;
    userPrincipalName: string;
  }>;
  /** List messages from the signed-in user's own mailbox. */
  listMessages(options?: ListMessagesOptions): Promise<{
    value: Record<string, unknown>[];
    "@odata.nextLink"?: string;
  }>;
  /** Get a single message by id from the signed-in user's own mailbox. */
  getMessage(id: string, options?: { select?: string[] }): Promise<Record<string, unknown>>;
  /** List the signed-in user's mail folders. */
  listFolders(options?: { top?: number }): Promise<{ value: Record<string, unknown>[] }>;
}

export function createMailboxReader(config: MailboxReaderConfig): MailboxReader;
