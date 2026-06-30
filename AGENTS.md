# AGENTS.md — integrating `graph-mailbox-reader` into another project

Instructions for an AI agent (e.g. Claude) asked to **use this package to read
a user's Microsoft 365 / Outlook mailbox in some other project.** Read this
fully before writing integration code. The human-facing docs are in
[README.md](./README.md); this file is the operational recipe.

## What this package does (and the one rule you must not break)

It reads **only the signed-in user's own mailbox** via Microsoft Graph `/me`,
using the delegated **device-code** OAuth flow. There is **no** `mailbox` /
`userId` / `email` parameter on any method, by design — a delegated `Mail.Read`
token can only ever act as the person who signed in.

- ❌ Do **not** add a parameter to target another mailbox. It's impossible with
  this token and defeats the security model. To read person X's mail, X signs in.
- ❌ Do **not** switch it to `client_credentials` / a client secret / `Mail.Read`
  *Application* permission. That can read the whole tenant and needs an Exchange
  access policy. If app-only sending is needed, that's a different package
  (`@hwoodgroup/email-client`), not this one.
- ❌ Do **not** commit `.env*` or `*.token-cache.json` (the cache holds refresh
  tokens). They must be gitignored.

## When this package fits — and when it does NOT

| Situation | Use this package? |
|---|---|
| CLI / script / cron that reads **my own** mailbox | ✅ Yes |
| Backend service that reads **one** fixed mailbox (personal dashboard, a bot's own inbox) | ✅ Yes — sign in once at startup |
| Multi-user **web app** where each end-user logs in and reads *their* mail per-request | ❌ No — device-code is interactive and process-wide. Use MSAL's **auth-code flow** with per-request tokens instead. |
| Read another employee's mailbox unattended, server-side, no sign-in | ❌ No — that needs app-only `Mail.Read` + an Exchange Application Access Policy. |

If the request is one of the ❌ rows, say so and propose the right approach
rather than forcing this package.

## Integration recipe

### 1. Install

Published:
```bash
npm install graph-mailbox-reader
```
Local/monorepo (sibling dir): add to `package.json` dependencies
`"graph-mailbox-reader": "file:../<path-to-package>"`, then `npm install`.

Requires Node ≥ 20, ESM (`"type": "module"`).

### 2. Provision the Azure app registration (one-time, per app)

Needs the Azure CLI (`az login` done) **or** do it via the portal (see README).
A **public client** with **delegated** `Mail.Read` + `User.Read`, **no secret**:

```bash
TENANT_ID="$(az account show --query tenantId -o tsv)"
APP_ID="$(az ad app create \
  --display-name "graph-mailbox-reader-<user-slug>" \
  --sign-in-audience AzureADMyOrg \
  --is-fallback-public-client true \
  --query appId -o tsv)"                     # --is-fallback-public-client = "Allow public client flows"
az ad sp create --id "$APP_ID" >/dev/null
az ad app permission add --id "$APP_ID" \
  --api 00000003-0000-0000-c000-000000000000 \
  --api-permissions 570282fd-fa5c-430d-a7fd-fc8dc98a9dca=Scope \
                    e1fe6dd8-ba31-4d61-89e7-88639da4683d=Scope   # Mail.Read, User.Read (Delegated)
# Optional, needs admin rights; otherwise the user consents at first sign-in:
az ad app permission admin-consent --id "$APP_ID" || true
echo "TENANT_ID=$TENANT_ID  CLIENT_ID=$APP_ID"
```

Model choice:
- **One app per user** (recommended) — run the above per person. Each gets a
  distinct `clientId`, so token caches never collide. No `tokenCachePath` needed.
- **One shared app, many users on one host** — same `clientId` for all; you
  **must** give each user a distinct `tokenCachePath` or they overwrite each
  other's cache.

### 3. Configure env

```
MICROSOFT_GRAPH_TENANT_ID=<tenant id from step 2>
MICROSOFT_GRAPH_CLIENT_ID=<app/client id from step 2>
```
No client secret. Gitignore the env file.

### 4. Use it in code

```js
import { createMailboxReader } from "graph-mailbox-reader";

const reader = createMailboxReader({
  tenantId: process.env.MICROSOFT_GRAPH_TENANT_ID,
  clientId: process.env.MICROSOFT_GRAPH_CLIENT_ID,
  // tokenCachePath: "...",      // ONLY for the shared-app case (step 2)
  // deviceCodeCallback: (r) => log(r.message),   // customise the sign-in prompt
});

await reader.signIn();           // device-code prompt on first run; silent (cached) after
const me = await reader.getProfile();

const { value } = await reader.listMessages({
  folder: "inbox",               // well-known name ("inbox","sentitems",...) or id; omit = all mail
  top: 25,
  orderBy: "receivedDateTime desc",
  // filter: "isRead eq false",  // OData $filter
  // search: "invoice",          // free-text; CANNOT combine with filter/orderBy
  // select: ["id","subject","from"],
});

const msg = await reader.getMessage(value[0].id);
const { value: folders } = await reader.listFolders();
await reader.signOut();          // clear cached account
```

### 5. Server pattern (single fixed mailbox)

Sign in **once at startup**, then serve — so requests already have a token:

```js
const reader = createMailboxReader({ tenantId, clientId });
await reader.signIn();           // prints device code to the server console on first boot only
app.get("/messages", async (_req, res) => {
  res.json(await reader.listMessages({ folder: "inbox", top: 20 }));
});
app.listen(PORT);
```

The device code prints to **stdout** on first boot; an operator opens the URL
and signs in once. After that the cached refresh token keeps it silent. This is
correct for a *personal/single-mailbox* service only (see the ❌ table above).

## E2E sign-in over HTTP (surface the device code through the API)

For end-to-end testing you want a tester to authenticate from the **API
itself** — read the device code off an endpoint (or a page), sign in, and have
the service continue — with no access to the server console. The pattern (this
is exactly what `mailbox-test-server` does):

1. **Start listening first**, then call `signIn()` **in the background** (don't
   `await` it before `listen`). `signIn()` blocks until the user authenticates,
   so it must never run inside a request handler — run it once at boot.
2. **`deviceCodeCallback` captures the prompt into shared state** the routes can
   read (`userCode`, `verificationUri`, `expiresIn`, `message`).
3. **A gate (`requireSignIn`) returns the code** to any mailbox request made
   before sign-in completes, so the tester learns the code from the endpoint
   they hit. (It also prevents a second device-code flow being started by a
   request calling `getProfile`/`listMessages` while the boot flow is pending.)
4. After the user signs in, the background `signIn()` resolves, state flips to
   signed-in, and the same requests now return data. The token is cached, so
   restarts are silent.

```js
const auth = { status: "starting", deviceCode: null, username: null, error: null };

const reader = createMailboxReader({
  tenantId, clientId,
  deviceCodeCallback: (response) => {       // response: { userCode, verificationUri, expiresIn, message }
    auth.status = "awaiting-code";
    auth.deviceCode = response;             // share it with the routes
  },
});

// Gate: serve the code instead of the data until sign-in finishes.
const requireSignIn = (_req, res, next) => {
  if (auth.status === "signed-in") return next();
  const body = { error: { message: "Not signed in. Enter the code, then retry." }, auth: { status: auth.status } };
  if (auth.status === "awaiting-code" && auth.deviceCode) {
    body.auth.userCode = auth.deviceCode.userCode;             // e.g. ABCD-EFGH
    body.auth.verificationUri = auth.deviceCode.verificationUri; // e.g. https://microsoft.com/devicelogin
    body.auth.expiresIn = auth.deviceCode.expiresIn;
  }
  res.status(auth.status === "awaiting-code" ? 401 : 503).json(body);
};

app.get("/auth/status", (_req, res) => res.json({ status: auth.status, username: auth.username }));
app.get("/messages", requireSignIn, (req, res) =>
  reader.listMessages({ folder: "inbox", top: 20 }).then((d) => res.json(d)));

// Listen FIRST, then sign in in the background so the code can be served.
app.listen(PORT, () => console.log(`listening on ${PORT}`));
reader.signIn()
  .then((account) => { auth.status = "signed-in"; auth.username = account?.username; auth.deviceCode = null; })
  .catch((err) => { auth.status = "error"; auth.error = err.message; });
```

Tester flow:
1. `GET /messages` → `401` with `userCode` + `verificationUri` in the body (or
   `503` for the brief "starting" window before Microsoft returns the code).
2. Open the URL, enter the code, sign in as the mailbox owner.
3. `GET /messages` again → `200` with the mail. Subsequent calls stay silent.

For a browser UX, render `userCode` + `verificationUri` on a page and poll
`/auth/status`, reloading once `status` becomes `signed-in`. A runnable
reference implementation (HTML sign-in page + JSON endpoints) lives in
`mailbox-test-server/server.js`.

## API reference

| Method | Returns | Graph endpoint |
|---|---|---|
| `signIn()` | account `{ username, ... }` (blocks until done) | device-code → token cache |
| `signOut()` | `void` | removes cached account |
| `getProfile()` | `{ id, displayName, mail, userPrincipalName }` | `GET /me` |
| `listMessages(options?)` | `{ value: Message[], "@odata.nextLink"? }` | `GET /me/messages` or `/me/mailFolders/{folder}/messages` |
| `getMessage(id, { select? })` | `Message` | `GET /me/messages/{id}` |
| `listFolders({ top? })` | `{ value: Folder[] }` | `GET /me/mailFolders` |

`ListMessagesOptions`: `{ folder?, top=25, skip?, search?, filter?, select?, orderBy? }`.
Full types in `src/index.d.ts`.

## Troubleshooting (common failures & fixes)

| Symptom | Cause → fix |
|---|---|
| `AADSTS7000218` on sign-in | App isn't a public client → set **Allow public client flows = Yes** (`--is-fallback-public-client true`). |
| Sign-in says consent required / blocked | Tenant requires admin consent → run `az ad app permission admin-consent`, or have an admin approve, or let the user consent at the prompt. |
| `400` from Graph on `$select` containing `wellKnownName` | That property is **beta-only**, not v1.0. Don't select it; match folders by `displayName`. |
| `400` when combining `search` with `filter`/`orderBy` | Graph forbids it for mail. Use `search` **or** `filter`+`orderBy`, never both. |
| Device-code prompt appears every run | No persisted cache, or wrong `tokenCachePath`, or `clientId` changed. Confirm the cache file at `~/.graph-mailbox-reader/<clientId>.token-cache.json` exists and is reused. |
| Two users clobber each other's session | Same `clientId` shared without distinct `tokenCachePath` → give each its own path, or one app per user. |
| Token cache "disappeared" after a rename | Cache path is keyed by `clientId` under `~/.graph-mailbox-reader/`. Renaming that dir or the `clientId` orphans the old cache. |

## Checklist before declaring the integration done

- [ ] Confirmed the use case is a ✅ row (single signed-in mailbox), not a ❌ row.
- [ ] Public-client app created; delegated `Mail.Read` + `User.Read`; **no secret**.
- [ ] `MICROSOFT_GRAPH_TENANT_ID` / `MICROSOFT_GRAPH_CLIENT_ID` set; env + `*.token-cache.json` gitignored.
- [ ] First-run `signIn()` completed (device-code), subsequent runs are silent.
- [ ] No code added a mailbox/userId parameter or a client secret.
