/**
 * Conversation history, persisted to disk.
 *
 * Conversations are small (a few KB of text), so each is one JSON file and an index
 * holds the summaries the list view needs. Writes go to a temp file and are renamed into
 * place, so a crash mid-save can never leave a half-written conversation behind.
 */
import { readFile, writeFile, rename, rm, mkdir, readdir } from "node:fs/promises";
import { randomUUID } from "node:crypto";
import path from "node:path";

/** Keep history bounded; the oldest conversations are pruned past this. */
export const MAX_CONVERSATIONS = 100;
const MAX_TITLE = 60;

/** Derive a readable title from the first thing the user said. */
export function titleFrom(messages) {
  const first = messages?.find((m) => m.role === "user" && m.text?.trim());
  if (!first) return "New conversation";
  const text = first.text.trim().replace(/\s+/g, " ");
  return text.length > MAX_TITLE ? `${text.slice(0, MAX_TITLE - 1)}…` : text;
}

/** Strip a message down to what is worth persisting. */
function sanitizeMessage(raw) {
  if (!raw || typeof raw !== "object") return null;
  const role = raw.role === "user" || raw.role === "assistant" || raw.role === "error" ? raw.role : null;
  if (!role) return null;

  const message = {
    id: typeof raw.id === "string" ? raw.id.slice(0, 64) : randomUUID(),
    role,
    text: typeof raw.text === "string" ? raw.text.slice(0, 100_000) : "",
    at: Number.isFinite(raw.at) ? raw.at : Date.now(),
  };

  if (Array.isArray(raw.actions)) {
    message.actions = raw.actions.slice(0, 20).map((a) => ({
      ...a,
      // A restored action can be inspected but not re-applied: the workbook has moved on
      // and the snapshot needed to undo it died with the session.
      status: a.status === "applied" || a.status === "dismissed" || a.status === "undone" ? a.status : "expired",
    }));
  }
  if (raw.stats && typeof raw.stats === "object") message.stats = raw.stats;
  return message;
}

export class ConversationStore {
  #dir;
  #indexFile;

  constructor(dataDir) {
    this.#dir = path.join(dataDir, "conversations");
    this.#indexFile = path.join(this.#dir, "index.json");
  }

  get directory() {
    return this.#dir;
  }

  async #readIndex() {
    try {
      const parsed = JSON.parse(await readFile(this.#indexFile, "utf8"));
      return Array.isArray(parsed) ? parsed : [];
    } catch {
      return [];
    }
  }

  async #writeAtomic(file, data) {
    await mkdir(this.#dir, { recursive: true });
    const tmp = `${file}.${process.pid}.tmp`;
    await writeFile(tmp, JSON.stringify(data, null, 2), "utf8");
    await rename(tmp, file);
  }

  #fileFor(id) {
    // Ids are generated here, but never trust one arriving from a request.
    if (!/^[a-f0-9-]{36}$/i.test(id)) throw new Error("Invalid conversation id");
    return path.join(this.#dir, `${id}.json`);
  }

  /** Summaries, newest first. */
  async list() {
    const index = await this.#readIndex();
    return index.sort((a, b) => b.updatedAt - a.updatedAt);
  }

  async get(id) {
    try {
      return JSON.parse(await readFile(this.#fileFor(id), "utf8"));
    } catch {
      return null;
    }
  }

  /**
   * Create or update a conversation. Returns its summary.
   * @param {{id?:string, title?:string, messages:Array}} input
   */
  async save({ id, title, messages }) {
    const clean = (Array.isArray(messages) ? messages : []).map(sanitizeMessage).filter(Boolean);
    if (!clean.length) throw new Error("A conversation needs at least one message");

    const now = Date.now();
    const index = await this.#readIndex();
    const existing = id ? index.find((c) => c.id === id) : null;

    const conversation = {
      id: existing?.id ?? (id && /^[a-f0-9-]{36}$/i.test(id) ? id : randomUUID()),
      title: (typeof title === "string" && title.trim() ? title.trim() : titleFrom(clean)).slice(0, MAX_TITLE),
      createdAt: existing?.createdAt ?? now,
      updatedAt: now,
      messages: clean,
    };

    await this.#writeAtomic(this.#fileFor(conversation.id), conversation);

    const summary = {
      id: conversation.id,
      title: conversation.title,
      createdAt: conversation.createdAt,
      updatedAt: conversation.updatedAt,
      messageCount: clean.length,
      preview: clean.find((m) => m.role === "assistant")?.text?.slice(0, 120) ?? "",
    };

    const next = [summary, ...index.filter((c) => c.id !== conversation.id)]
      .sort((a, b) => b.updatedAt - a.updatedAt);

    // Prune the tail, removing the files too so the directory does not grow forever.
    const pruned = next.slice(MAX_CONVERSATIONS);
    await Promise.all(pruned.map((c) => rm(this.#fileFor(c.id), { force: true }).catch(() => {})));

    await this.#writeAtomic(this.#indexFile, next.slice(0, MAX_CONVERSATIONS));
    return summary;
  }

  async remove(id) {
    const index = await this.#readIndex();
    await rm(this.#fileFor(id), { force: true });
    await this.#writeAtomic(this.#indexFile, index.filter((c) => c.id !== id));
  }

  async clear() {
    try {
      for (const file of await readdir(this.#dir)) {
        if (file.endsWith(".json")) await rm(path.join(this.#dir, file), { force: true });
      }
    } catch { /* nothing to clear */ }
  }
}
