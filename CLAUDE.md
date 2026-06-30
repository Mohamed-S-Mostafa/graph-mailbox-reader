# graph-mailbox-reader

Delegated-auth **reader** for a Microsoft 365 / Outlook mailbox. Reads **only
the signed-in user's own mailbox** via Microsoft Graph `/me`, using the OAuth
2.0 **device-code (delegated) flow**. No client secret.

See [README.md](./README.md) for full usage, config, and the Azure app
registration steps.

## The one guarantee

Every request targets `/me`. There is **no** `mailbox` / `userId` parameter on
any method — a delegated `Mail.Read` token acts as the signed-in user, so Graph
itself refuses to return anyone else's mail. To read a different person's
mailbox, that person signs in (their own `tokenCachePath` or `clientId`).

## Layout

- `src/auth.js` — MSAL `PublicClientApplication`; device-code sign-in, silent
  refresh, on-disk token cache (mode `0600`).
- `src/client.js` — API surface; all paths hardwired to `/me`.
- `src/index.js` / `src/index.d.ts` — public entry point and types.
- `test/` — unit tests (config validation + OData query construction; no network).
- `example/read-my-mailbox.js` — manual smoke test against a real mailbox.
- `scripts/create-azure-app.sh` — local-only helper to provision a per-user
  Azure AD app registration (gitignored; not published).

## Conventions

- ESM only (`"type": "module"`), Node >= 18.
- Only `src/` is published to npm (`files: ["src"]`).
- Never commit `.env*` files or `*.token-cache.json` (refresh tokens).
- Proprietary license — see [LICENSE](./LICENSE).

## Test

```bash
npm test
```
