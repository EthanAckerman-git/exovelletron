/**
 * @vitest-environment jsdom
 */
import { describe, it, expect, vi } from "vitest";
import { buildProposal, armProposals } from "../../taskpane/ui/render.js";

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
