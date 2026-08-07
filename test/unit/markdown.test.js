/**
 * @vitest-environment jsdom
 */
import { describe, it, expect, beforeEach } from "vitest";
import { renderMarkdown, renderInline } from "../../taskpane/ui/markdown.js";

let root;
beforeEach(() => {
  root = document.createElement("div");
});

const render = (md) => {
  renderMarkdown(root, md);
  return root;
};

describe("inline formatting", () => {
  it("renders bold, italic, code, and strikethrough", () => {
    const p = document.createElement("p");
    renderInline(p, "plain **bold** *italic* `code` ~~gone~~");
    expect(p.querySelector("strong").textContent).toBe("bold");
    expect(p.querySelector("em").textContent).toBe("italic");
    expect(p.querySelector("code").textContent).toBe("code");
    expect(p.querySelector("s").textContent).toBe("gone");
    expect(p.textContent).toBe("plain bold italic code gone");
  });

  it("handles underscore emphasis without breaking snake_case", () => {
    const p = document.createElement("p");
    renderInline(p, "__strong__ and some_variable_name");
    expect(p.querySelector("strong").textContent).toBe("strong");
    expect(p.textContent).toContain("some_variable_name");
  });

  it("nests bold italic", () => {
    const p = document.createElement("p");
    renderInline(p, "***both***");
    expect(p.querySelector("strong em").textContent).toBe("both");
  });

  it("leaves cell addresses and formulas alone", () => {
    const p = document.createElement("p");
    renderInline(p, "Fill E2:E501 with =C2*D2");
    expect(p.textContent).toBe("Fill E2:E501 with =C2*D2");
    expect(p.querySelector("em")).toBeNull();
  });

  // Model output is untrusted; it must never become markup.
  it("never interprets HTML", () => {
    const out = render('<img src=x onerror="alert(1)"> and <b>bold</b>');
    expect(out.querySelector("img")).toBeNull();
    expect(out.querySelector("b")).toBeNull();
    expect(out.textContent).toContain("<img src=x");
  });
});

describe("block formatting", () => {
  it("splits paragraphs on blank lines", () => {
    const out = render("First para.\nSame para.\n\nSecond para.");
    const paras = out.querySelectorAll("p");
    expect(paras).toHaveLength(2);
    expect(paras[0].textContent).toBe("First para. Same para.");
  });

  it("renders bullet lists", () => {
    const out = render("- one\n- two\n- three");
    expect(out.querySelectorAll("ul li")).toHaveLength(3);
    expect(out.querySelector("ul li").textContent).toBe("one");
  });

  it("renders numbered lists and keeps a non-1 start", () => {
    const out = render("3. third\n4. fourth");
    expect(out.querySelectorAll("ol li")).toHaveLength(2);
    expect(out.querySelector("ol").getAttribute("start")).toBe("3");
  });

  it("folds indented continuation lines into the item above", () => {
    const out = render("- first line\n  continued here\n- second");
    const items = out.querySelectorAll("li");
    expect(items).toHaveLength(2);
    expect(items[0].textContent).toBe("first line continued here");
  });

  it("renders headings below the page level", () => {
    const out = render("# Big\n## Smaller");
    expect(out.querySelector("h4").textContent).toBe("Big");
    expect(out.querySelector("h5").textContent).toBe("Smaller");
  });

  it("renders fenced code verbatim with its language", () => {
    const out = render("```excel\n=SUM(A1:A9)\n=AVERAGE(B:B)\n```");
    const code = out.querySelector("pre code");
    expect(code.textContent).toBe("=SUM(A1:A9)\n=AVERAGE(B:B)");
    expect(code.className).toBe("lang-excel");
  });

  it("does not format inside a code block", () => {
    const out = render("```\n**not bold**\n```");
    expect(out.querySelector("strong")).toBeNull();
    expect(out.querySelector("pre code").textContent).toBe("**not bold**");
  });

  it("renders an unterminated fence rather than swallowing the reply", () => {
    const out = render("```\nstill shown");
    expect(out.querySelector("pre code").textContent).toBe("still shown");
  });

  it("renders tables", () => {
    const out = render("| Col | Meaning |\n| --- | --- |\n| A | Region |\n| B | Rep |");
    expect(out.querySelectorAll("th")).toHaveLength(2);
    expect(out.querySelectorAll("tbody tr")).toHaveLength(2);
    expect(out.querySelector("tbody td").textContent).toBe("A");
    // Wide tables scroll inside their own container rather than the pane.
    expect(out.querySelector(".md-table")).toBeTruthy();
  });

  it("treats a pipe line with no divider as ordinary text", () => {
    const out = render("this | is | not a table");
    expect(out.querySelector("table")).toBeNull();
    expect(out.querySelector("p").textContent).toBe("this | is | not a table");
  });

  it("renders blockquotes and dividers", () => {
    const out = render("> quoted\n> more\n\n---");
    expect(out.querySelector("blockquote").textContent).toBe("quoted more");
    expect(out.querySelector("hr")).toBeTruthy();
  });

  it("handles an empty reply", () => {
    expect(render("").children).toHaveLength(0);
    expect(render(null).children).toHaveLength(0);
  });

  it("renders a realistic mixed reply", () => {
    const out = render([
      "Here's what I found:",
      "",
      "- **Revenue** is blank for 12 rows",
      "- `Unit Price` has text in 3 cells",
      "",
      "Suggested fix:",
      "",
      "```excel",
      "=IFERROR(C2*D2,\"\")",
      "```",
    ].join("\n"));

    expect(out.querySelectorAll("p")).toHaveLength(2);
    expect(out.querySelectorAll("li")).toHaveLength(2);
    expect(out.querySelector("li strong").textContent).toBe("Revenue");
    expect(out.querySelector("pre code").textContent).toContain("IFERROR");
  });
});
