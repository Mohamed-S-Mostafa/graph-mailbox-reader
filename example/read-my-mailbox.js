/**
 * Manual smoke test / usage example.
 *
 * Reads the signed-in user's own mailbox. On first run it prints a device-code
 * sign-in prompt; after that the cached token is reused silently.
 *
 *   MICROSOFT_GRAPH_TENANT_ID=... MICROSOFT_GRAPH_CLIENT_ID=... npm run example
 */
import { createMailboxReader } from "../src/index.js";

const tenantId = process.env.MICROSOFT_GRAPH_TENANT_ID;
const clientId = process.env.MICROSOFT_GRAPH_CLIENT_ID;

if (!tenantId || !clientId) {
  console.error(
    "Set MICROSOFT_GRAPH_TENANT_ID and MICROSOFT_GRAPH_CLIENT_ID env vars first.",
  );
  process.exit(1);
}

const reader = createMailboxReader({ tenantId, clientId });

const account = await reader.signIn();
console.log(`\nSigned in as: ${account?.username}\n`);

const me = await reader.getProfile();
console.log(`Mailbox: ${me.mail || me.userPrincipalName}\n`);

const { value: messages } = await reader.listMessages({
  folder: "inbox",
  top: 10,
  orderBy: "receivedDateTime desc",
});

console.log(`Latest ${messages.length} inbox messages:\n`);
for (const m of messages) {
  const flag = m.isRead ? "  " : "• ";
  const from = m.from?.emailAddress?.address ?? "(unknown)";
  console.log(`${flag}${m.receivedDateTime}  ${from}\n   ${m.subject}`);
}
