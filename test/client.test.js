import { describe, it } from "node:test";
import assert from "node:assert/strict";
import { createMailboxReader } from "../src/index.js";

/**
 * Unit tests covering config validation and the Graph query construction that
 * `listMessages` / `getMessage` rely on. Network calls are not exercised here.
 */

describe("createMailboxReader() — config validation", () => {
  it("throws when tenantId is missing", () => {
    assert.throws(() => createMailboxReader({ clientId: "abc" }), /tenantId/);
  });

  it("throws when clientId is missing", () => {
    assert.throws(() => createMailboxReader({ tenantId: "abc" }), /clientId/);
  });

  it("returns a reader with the expected read-only surface", () => {
    const reader = createMailboxReader({ tenantId: "t", clientId: "c" });
    for (const method of ["signIn", "signOut", "getProfile", "listMessages", "getMessage", "listFolders"]) {
      assert.equal(typeof reader[method], "function", `missing ${method}`);
    }
    // The reader must not expose any way to target another mailbox.
    assert.equal(reader.getMailbox, undefined);
    assert.equal(reader.readUserMailbox, undefined);
  });
});

/**
 * Mirror of the private toMessageQuery() so we can assert the OData mapping
 * without exporting internals. Kept in lockstep with src/client.js.
 */
function toMessageQuery({ top = 25, skip, search, filter, select, orderBy } = {}) {
  const DEFAULT_MESSAGE_FIELDS = [
    "id", "subject", "from", "toRecipients", "receivedDateTime",
    "isRead", "hasAttachments", "bodyPreview", "webLink",
  ];
  const query = {};
  query.$top = top;
  if (skip) query.$skip = skip;
  if (search) {
    query.$search = `"${search}"`;
  } else {
    if (filter) query.$filter = filter;
    if (orderBy) query.$orderby = orderBy;
  }
  query.$select = (select && select.length ? select : DEFAULT_MESSAGE_FIELDS).join(",");
  return query;
}

describe("message query construction", () => {
  it("defaults to top=25 and the compact field projection", () => {
    const q = toMessageQuery();
    assert.equal(q.$top, 25);
    assert.match(q.$select, /^id,subject,from,/);
    assert.equal(q.$skip, undefined);
  });

  it("quotes the search term and drops filter/orderBy when searching", () => {
    const q = toMessageQuery({ search: "invoice", filter: "isRead eq false", orderBy: "receivedDateTime desc" });
    assert.equal(q.$search, '"invoice"');
    assert.equal(q.$filter, undefined);
    assert.equal(q.$orderby, undefined);
  });

  it("passes filter and orderBy through when not searching", () => {
    const q = toMessageQuery({ filter: "isRead eq false", orderBy: "receivedDateTime desc" });
    assert.equal(q.$filter, "isRead eq false");
    assert.equal(q.$orderby, "receivedDateTime desc");
    assert.equal(q.$search, undefined);
  });

  it("honours an explicit select list", () => {
    const q = toMessageQuery({ select: ["id", "subject"] });
    assert.equal(q.$select, "id,subject");
  });
});
