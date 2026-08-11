import { describe, it, expect, beforeAll, afterAll } from "vitest";

/**
 * taskpane/api.js parses the chat SSE stream by hand. This pins the seam the read
 * bridge depends on: a `read_request` frame must reach the onReadRequest handler, and
 * answering it must POST to /api/chat/read with the token header.
 */

const FRAMES =
  'event: read_request\ndata: {"id":"r1","sheet":"Sales","address":"A2:B3"}\n\n' +
  'event: token\ndata: {"text":"hello"}\n\n' +
  'event: done\ndata: {"text":"hello","actions":[],"stats":{}}\n\n';

let api;
const calls = [];
const realFetch = globalThis.fetch;

beforeAll(async () => {
  globalThis.window = { __EAL_TOKEN__: "tok-123" };
  globalThis.fetch = async (path, opts = {}) => {
    calls.push({ path, opts });
    if (path === "/api/chat") {
      const bytes = new TextEncoder().encode(FRAMES);
      let sent = false;
      return {
        ok: true,
        body: {
          getReader: () => ({
            read: async () => (sent ? { done: true } : ((sent = true), { done: false, value: bytes })),
          }),
        },
      };
    }
    return { ok: true, status: 200, json: async () => ({ ok: true }) };
  };
  api = await import("../../taskpane/api.js");
});

afterAll(() => {
  globalThis.fetch = realFetch;
  delete globalThis.window;
});

describe("the pane side of the read bridge", () => {
  it("dispatches read_request frames and posts the answer back with the token", async () => {
    const seen = [];
    let done = null;
    api.streamChat(
      { message: "hi" },
      {
        onReadRequest: (payload) => {
          seen.push(payload);
          api.postReadResult({ id: payload.id, ok: true, rows: [["x"]] });
        },
        onDone: (payload) => { done = payload; },
      },
    );

    // The stream is synchronous fake data; a couple of microtask turns settles it.
    await new Promise((r) => setTimeout(r, 20));

    expect(seen).toEqual([{ id: "r1", sheet: "Sales", address: "A2:B3" }]);
    expect(done?.text).toBe("hello");

    const post = calls.find((c) => c.path === "/api/chat/read");
    expect(post).toBeDefined();
    expect(post.opts.method).toBe("POST");
    expect(post.opts.headers["x-eal-token"]).toBe("tok-123");
    expect(JSON.parse(post.opts.body)).toMatchObject({ id: "r1", ok: true });
  });
});
