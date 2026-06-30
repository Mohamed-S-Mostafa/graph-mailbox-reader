# graph-mailbox-reader

Delegated-auth **reader** for a Microsoft 365 / Outlook mailbox. Reads **only
the signed-in user's own mailbox** via Microsoft Graph `/me`, using the OAuth
2.0 **device-code (delegated) flow**. No client secret.

```bash
npm install graph-mailbox-reader
```

> **Integrating this into another project with an AI agent?** See
> [AGENTS.md](./AGENTS.md) — a step-by-step playbook (Azure setup, env, code,
> server pattern, troubleshooting) written for Claude/agents. It ships with the
> package, so it's also at `node_modules/graph-mailbox-reader/AGENTS.md`.

## The one guarantee

> The reader can only ever read the mailbox of the person who signed in.

Every request targets `/me`. There is **no** `mailbox` / `userId` parameter on
any method. A delegated `Mail.Read` token acts *as the signed-in user*, so
Graph itself refuses to return anyone else's mail. "Which mailbox" is decided
entirely by **who signs in** — not by config a caller could tamper with.

## Reading a different person's mailbox

There is nothing to reconfigure in code. To read person X's mailbox, **X signs
in**. To run several side by side:

- **Same app registration, per-user cache** — give each person their own
  `tokenCachePath`; each signs in once with their own Microsoft account.
- **Separate app registrations** — create another public-client app and pass
  its `clientId`. Each app still only reads its signed-in user's mailbox.

## Why delegated (not client_credentials)?

For *reading* a single person's mailbox, app-only auth is the wrong tool: the
`Mail.Read` **Application** permission can read **every mailbox in the tenant**,
and narrowing it to one mailbox needs an admin-managed Exchange *Application
Access Policy*. Delegated `Mail.Read` is scoped to the signed-in user **by
construction** — no policy, no secret, no over-privilege. Public clients hold
no secret.

## Azure app registration (one-time)

1. Azure Portal → App registrations → New registration.
2. **Authentication** → Advanced settings → **Allow public client flows = Yes**
   (required for device-code).
3. **API permissions** → Microsoft Graph → **Delegated** → add `Mail.Read` and
   `User.Read`. Grant admin consent if your tenant requires it.
4. No client secret, no redirect URI needed for device-code.

## Config

| Option | Required | Notes |
|---|---|---|
| `tenantId` | yes | Azure AD tenant ID (or `"common"`/`"organizations"`) |
| `clientId` | yes | App registration (**public client**) ID |
| `tokenCachePath` | no | Defaults to `~/.graph-mailbox-reader/<clientId>.token-cache.json` |
| `deviceCodeCallback` | no | Receives sign-in instructions; defaults to `console.log` |

Env vars used by the example: `MICROSOFT_GRAPH_TENANT_ID`,
`MICROSOFT_GRAPH_CLIENT_ID`. **No client secret.**

## API

```js
import { createMailboxReader } from "graph-mailbox-reader";

const reader = createMailboxReader({
  tenantId: process.env.MICROSOFT_GRAPH_TENANT_ID,
  clientId: process.env.MICROSOFT_GRAPH_CLIENT_ID,
});

await reader.signIn();                 // prints device-code prompt on first run only
const me = await reader.getProfile();  // { id, displayName, mail, userPrincipalName }

const { value } = await reader.listMessages({
  folder: "inbox",                     // well-known name or folder id; omit for all mail
  top: 25,
  orderBy: "receivedDateTime desc",    // ignored if `search` is set
  // filter: "isRead eq false",        // ignored if `search` is set
  // search: "invoice",                // free-text; can't combine with filter/orderBy
  // select: ["id", "subject", "from"],
});

const msg = await reader.getMessage(value[0].id);
const { value: folders } = await reader.listFolders();

await reader.signOut();                // clears cached account → fresh sign-in next time
```

### Methods

| Method | Graph endpoint |
|---|---|
| `getProfile()` | `GET /me` |
| `listMessages(options)` | `GET /me/messages` \| `GET /me/mailFolders/{folder}/messages` |
| `getMessage(id, options)` | `GET /me/messages/{id}` |
| `listFolders(options)` | `GET /me/mailFolders` |
| `signIn()` / `signOut()` | device-code sign-in (blocks) / clear cached account |

### Sign-in over HTTP

For a server that hands the device code back to the API caller (E2E testing,
no console access), run `signIn()` in the background at startup and capture the
code via `deviceCodeCallback`. See
[AGENTS.md → "E2E sign-in over HTTP"](./AGENTS.md#e2e-sign-in-over-http-surface-the-device-code-through-the-api),
with a runnable reference in `mailbox-test-server/server.js`.

## Architecture

```
createMailboxReader({ tenantId, clientId, tokenCachePath?, deviceCodeCallback? })
│
├── Auth layer (src/auth.js — MSAL PublicClientApplication)
│   ├── acquireTokenSilent        → reuse/refresh cached token (no prompt)
│   ├── acquireTokenByDeviceCode  → first-run interactive sign-in
│   └── File token cache (mode 0600) → sign-in survives restarts
│
├── HTTP layer (graphGet) → axios GET graph.microsoft.com/v1.0 + Bearer token
│
└── API layer (src/client.js) — all paths hardwired to /me
```

## Test / run

```bash
npm test                               # unit tests (no network)
MICROSOFT_GRAPH_TENANT_ID=... MICROSOFT_GRAPH_CLIENT_ID=... npm run example
```

## License

Proprietary — All Rights Reserved. See [LICENSE](./LICENSE). The source is
public for reference and use of the published package, but it may not be
copied, modified, or redistributed.
