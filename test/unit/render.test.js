/**
 * @vitest-environment jsdom
 */
import { describe, it, expect, vi } from "vitest";
import { buildProposal, buildProposalGroup, armProposals } from "../../taskpane/ui/render.js";

const action = {
  id: "1",
  type: "write_formula",
  address: "E2:E6",
  formula: "=C2*D2",
  summary: "Fill 5 cells",
};

const tick = () => new Promise((r) => setTimeout(r, 0));

describe("proposal cards during a streaming reply", () => {
  // Regression: cards appear while the reply is still streaming, and the engine is
  // single-threaded — clicking Apply mid-stream always failed with "the model is busy".
  it("starts disarmed and only goes live when the turn finishes", () => {
    const card = buildProposal(action, { onApply: vi.fn() }, { deferActions: true });
    document.body.appendChild(card);

    const [apply, dismiss] = card.querySelectorAll(".proposal__actions .btn");
    expect(apply.disabled).toBe(true);
    expect(dismiss.disabled).toBe(true);
    expect(card.querySelector(".proposal__wait")).toBeTruthy();

    armProposals(document.body);
    expect(apply.disabled).toBe(false);
    expect(dismiss.disabled).toBe(false);
    expect(card.querySelector(".proposal__wait")).toBeNull();
    document.body.textContent = "";
  });

  it("is live immediately when built outside a stream", () => {
    const card = buildProposal(action, { onApply: vi.fn() });
    expect(card.querySelector(".btn--primary").disabled).toBe(false);
    expect(card.querySelector(".proposal__wait")).toBeNull();
  });

  it("refuses an apply while the engine is busy, and says why", async () => {
    const onApply = vi.fn();
    const card = buildProposal(action, { onApply, canApply: () => "Wait for the current reply to finish." });
    document.body.appendChild(card);

    card.querySelector(".btn--primary").click();
    await tick();

    expect(onApply).not.toHaveBeenCalled();
    expect(card.querySelector(".proposal__error").textContent).toMatch(/Wait for the current reply/);
    // The buttons stay usable — the refusal is momentary, not a dead end.
    expect(card.querySelector(".btn--primary").disabled).toBe(false);
    document.body.textContent = "";
  });

  it("applies normally once the engine is free", async () => {
    const onApply = vi.fn(async () => ({ undo: async () => {} }));
    const card = buildProposal(action, { onApply, canApply: () => true });
    document.body.appendChild(card);

    card.querySelector(".btn--primary").click();
    await tick();

    expect(onApply).toHaveBeenCalledOnce();
    expect(card.dataset.status).toBe("applied");
    document.body.textContent = "";
  });
});

describe("grouped proposals: one reply, one Apply", () => {
  const header = { id: "h1", type: "insert_column", address: "D1", header: "Stock", summary: 'Header "Stock"' };
  const fill = { id: "f1", type: "write_formula", address: "D2:D6", formula: "=Inventory!B2", summary: "Fill 5 cells" };

  it("absorbs several instant changes into one card with one Apply", () => {
    const group = buildProposalGroup({ onApply: vi.fn() }, { deferActions: true });
    group.add(header);
    expect(group.card.dataset.multi).toBe("false");
    expect(group.card.querySelector(".proposal__label").textContent).toBe("Proposed change");

    group.add(fill);
    expect(group.card.dataset.multi).toBe("true");
    expect(group.card.querySelector(".proposal__label").textContent).toBe("Proposed changes (2)");
    expect(group.card.querySelectorAll(".proposal__section")).toHaveLength(2);
    // One Apply, one Dismiss — never a button per change.
    expect(group.card.querySelectorAll(".proposal__actions .btn")).toHaveLength(2);
    expect(group.card.querySelector(".btn--primary").textContent).toBe("Apply all 2");
  });

  it("applies the changes in order and undoes them in reverse with one click", async () => {
    const log = [];
    const onApply = vi.fn(async (action) => {
      log.push(`apply:${action.id}`);
      return { undo: async () => log.push(`undo:${action.id}`) };
    });
    const group = buildProposalGroup({ onApply, canApply: () => true });
    document.body.appendChild(group.card);
    group.add(header);
    group.add(fill);

    group.card.querySelector(".btn--primary").click();
    await tick();
    expect(log).toEqual(["apply:h1", "apply:f1"]);
    expect(group.card.dataset.status).toBe("applied");
    expect(group.card.querySelector(".proposal__receipt").textContent).toMatch(/2 changes written/);

    group.card.querySelector(".proposal__receipt .link").click();
    await tick();
    expect(log).toEqual(["apply:h1", "apply:f1", "undo:f1", "undo:h1"]);
    document.body.textContent = "";
  });

  it("rolls back what already landed when a later change fails", async () => {
    const log = [];
    const onApply = vi.fn(async (action) => {
      if (action.id === "f1") throw new Error("range is protected");
      log.push(`apply:${action.id}`);
      return { undo: async () => log.push(`undo:${action.id}`) };
    });
    const group = buildProposalGroup({ onApply, canApply: () => true });
    document.body.appendChild(group.card);
    group.add(header);
    group.add(fill);

    group.card.querySelector(".btn--primary").click();
    await tick();
    // The header applied, the fill failed, the header was unwound.
    expect(log).toEqual(["apply:h1", "undo:h1"]);
    expect(group.card.dataset.status).toBe("pending");
    expect(group.card.querySelector(".proposal__error").textContent).toMatch(/rolled back/);
    // Still usable: the user can fix the sheet and try again.
    expect(group.card.querySelector(".btn--primary").disabled).toBe(false);
    expect(group.card.querySelector(".btn--primary").textContent).toBe("Apply all 2");
    document.body.textContent = "";
  });

  it("starts disarmed while the reply streams, like single cards", () => {
    const group = buildProposalGroup({ onApply: vi.fn() }, { deferActions: true });
    document.body.appendChild(group.card);
    group.add(header);
    expect(group.card.querySelector(".btn--primary").disabled).toBe(true);
    armProposals(document.body);
    expect(group.card.querySelector(".btn--primary").disabled).toBe(false);
    document.body.textContent = "";
  });
});

describe("the unified card: long-running jobs live inside the group", () => {
  const split = {
    id: "s1", type: "split_column", address: "D2:F10", source: "A2:A10", rows: 9,
    columns: ["Street", "City", "Zip"], instruction: "Split the address", summary: "Split 9 rows",
  };
  const rewrite = {
    id: "r1", type: "transform_column", address: "G2:G10", source: "B2:B10", rows: 9,
    instruction: "USPS format", summary: "Rewrite 9 rows from B2:B10 into G2:G10",
  };

  it("labels a lone job like its old card did, and a mix as one Apply", () => {
    const alone = buildProposalGroup({ onApply: vi.fn() });
    alone.add(split);
    expect(alone.card.querySelector(".btn--primary").textContent).toBe("Split 9 rows");
    expect(alone.card.querySelector(".transform__progress").hidden).toBe(true);

    const mixed = buildProposalGroup({ onApply: vi.fn() });
    mixed.add(split);
    mixed.add(rewrite);
    expect(mixed.card.querySelector(".btn--primary").textContent).toBe("Apply all 2");
    expect(mixed.card.querySelectorAll(".proposal__section")).toHaveLength(2);
    expect(mixed.card.querySelectorAll(".proposal__actions .btn")).toHaveLength(2);
  });

  it("runs jobs in order, streams progress into the right section, and settles results", async () => {
    const order = [];
    const onApply = vi.fn(async (action, handlers) => {
      order.push(action.id);
      expect(typeof handlers.onCancelHandle).toBe("function");
      handlers.onPhase?.("reading");
      handlers.onProgress?.({ done: 5, total: 9, sample: ["1 MAIN ST"] });
      return { undo: async () => order.push(`undo:${action.id}`), results: ["a", "b"], failed: [], lossy: [] };
    });
    const group = buildProposalGroup({ onApply, canApply: () => true });
    document.body.appendChild(group.card);
    group.add(split);
    group.add(rewrite);

    group.card.querySelector(".btn--primary").click();
    await tick();

    expect(order).toEqual(["s1", "r1"]);
    expect(group.card.dataset.status).toBe("applied");
    // Each section settled with its own results.
    const samples = group.card.querySelectorAll(".transform__sample");
    expect(samples).toHaveLength(2);
    expect(samples[0].textContent).toContain("a");
    document.body.textContent = "";
  });

  it("turns Dismiss into Stop while a job runs, wired to that job's cancel", async () => {
    let cancelled = false;
    let release;
    const onApply = vi.fn(
      (action, handlers) =>
        new Promise((resolve, reject) => {
          handlers.onCancelHandle(() => { cancelled = true; reject(new Error("Transform cancelled")); });
          release = resolve;
        }),
    );
    const group = buildProposalGroup({ onApply, canApply: () => true });
    document.body.appendChild(group.card);
    group.add(split);

    group.card.querySelector(".btn--primary").click();
    await tick();

    const dismiss = group.card.querySelectorAll(".proposal__actions .btn")[1];
    expect(dismiss.textContent).toBe("Stop");
    dismiss.click();
    await tick();

    expect(cancelled).toBe(true);
    expect(group.card.querySelector(".proposal__error").textContent).toMatch(/Transform cancelled/);
    // Back to a usable pending card.
    expect(group.card.querySelector(".btn--primary").disabled).toBe(false);
    document.body.textContent = "";
    void release;
  });

  it("reports a partial undo honestly when a later change was overwritten", async () => {
    const onApply = vi.fn(async (action) => ({
      undo: action.id === "s1"
        ? async () => { throw new Error("the sheet has changed here since this was applied"); }
        : async () => {},
    }));
    const group = buildProposalGroup({ onApply, canApply: () => true });
    document.body.appendChild(group.card);
    group.add(split);
    group.add(rewrite);

    group.card.querySelector(".btn--primary").click();
    await tick();
    group.card.querySelector(".proposal__receipt .link").click();
    await tick();

    // r1 undid, s1 refused — the receipt says exactly that and Undo stays clickable.
    expect(group.card.querySelector(".proposal__receipt").textContent)
      .toMatch(/Undid 1 of 2, then stopped: the sheet has changed/);
    expect(group.card.querySelector(".proposal__receipt .link").disabled).toBe(false);
    document.body.textContent = "";
  });
});
