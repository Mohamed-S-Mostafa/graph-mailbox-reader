import { describe, it } from "node:test";
import assert from "node:assert/strict";
import { createMailboxReader } from "../src/index.js";
import { toMessageQuery, scrubGraphError } from "../src/client.js";

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

describe("scrubGraphError() — never leaks the access token", () => {
  // Shape of a real AxiosError: the bearer token lives in config.headers.
  const axiosLikeError = {
    message: "Request failed with status code 401",
    config: { headers: { Authorization: "Bearer SUPER-SECRET-TOKEN" } },
    response: {
      status: 401,
      data: { error: { code: "InvalidAuthenticationToken", message: "Access token expired." } },
    },
  };

  it("strips the request config (and thus the bearer token)", () => {
    const clean = scrubGraphError(axiosLikeError, "/me/messages");
    assert.equal(clean.config, undefined);
    assert.equal(JSON.stringify(clean).includes("SUPER-SECRET-TOKEN"), false);
    assert.equal(clean.message.includes("SUPER-SECRET-TOKEN"), false);
  });

  it("preserves the useful Graph diagnostics", () => {
    const clean = scrubGraphError(axiosLikeError, "/me/messages");
    assert.equal(clean.status, 401);
    assert.equal(clean.code, "InvalidAuthenticationToken");
    assert.match(clean.message, /Graph 401 on \/me\/messages: Access token expired\./);
  });

  it("falls back gracefully when there is no response body", () => {
    const clean = scrubGraphError({ message: "timeout of 30000ms exceeded" }, "/me");
    assert.match(clean.message, /Graph request error on \/me: timeout of 30000ms exceeded/);
    assert.equal(clean.status, undefined);
  });
});
