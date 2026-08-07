/**
 * A small, safe Markdown renderer for assistant replies.
 *
 * Model output is untrusted text, so this builds DOM nodes directly and never assigns
 * innerHTML. It covers what the model actually produces — paragraphs, lists, headings,
 * emphasis, code, tables, quotes — and deliberately ignores everything else. Links in
 * particular are rendered as plain text: this is an offline tool, and a clickable URL
 * invented by a language model is not something to hand a user.
 *
 * Pure DOM construction, so it can be unit-tested against a document stub.
 */

/* ------------------------------------------------------------------ inline parsing */

/** Ordered so the longest delimiters win: `***` before `**` before `*`. */
const INLINE_PATTERN = /(`[^`\n]+`|\*\*\*[^*\n]+\*\*\*|\*\*[^*\n]+\*\*|__[^_\n]+__|~~[^~\n]+~~|\*[^*\n]+\*|(?<![A-Za-z0-9])_[^_\n]+_(?![A-Za-z0-9]))/g;

/**
 * Append inline-formatted text to a node.
 * @param {Node} parent
 * @param {string} text
 */
export function renderInline(parent, text) {
  const doc = parent.ownerDocument;
  for (const part of String(text).split(INLINE_PATTERN)) {
    if (!part) continue;

    const wrap = (tag, inner) => {
      const node = doc.createElement(tag);
      // Emphasis can contain code and vice versa is not worth supporting; one level is.
      node.textContent = inner;
      parent.appendChild(node);
    };

    if (part.length > 2 && part.startsWith("`") && part.endsWith("`")) wrap("code", part.slice(1, -1));
    else if (part.length > 6 && part.startsWith("***") && part.endsWith("***")) {
      const strong = doc.createElement("strong");
      const em = doc.createElement("em");
      em.textContent = part.slice(3, -3);
      strong.appendChild(em);
      parent.appendChild(strong);
    } else if (part.length > 4 && part.startsWith("**") && part.endsWith("**")) wrap("strong", part.slice(2, -2));
    else if (part.length > 4 && part.startsWith("__") && part.endsWith("__")) wrap("strong", part.slice(2, -2));
    else if (part.length > 4 && part.startsWith("~~") && part.endsWith("~~")) wrap("s", part.slice(2, -2));
    else if (part.length > 2 && part.startsWith("*") && part.endsWith("*")) wrap("em", part.slice(1, -1));
    else if (part.length > 2 && part.startsWith("_") && part.endsWith("_")) wrap("em", part.slice(1, -1));
    else parent.appendChild(doc.createTextNode(part));
  }
}

/* ------------------------------------------------------------------- block parsing */

const RULES = {
  fence: /^\s{0,3}```(.*)$/,
  heading: /^\s{0,3}(#{1,4})\s+(.*)$/,
  bullet: /^\s{0,3}[-*+]\s+(.*)$/,
  ordered: /^\s{0,3}(\d{1,3})[.)]\s+(.*)$/,
  quote: /^\s{0,3}>\s?(.*)$/,
  divider: /^\s{0,3}([-*_])\s*(\1\s*){2,}$/,
  tableRow: /^\s*\|(.+)\|\s*$/,
  tableDivider: /^\s*\|[\s:|-]+\|\s*$/,
};

const splitRow = (line) => line.replace(/^\s*\|/, "").replace(/\|\s*$/, "").split("|").map((c) => c.trim());

/**
 * Render Markdown into `container`, replacing its contents.
 * @param {Element} container
 * @param {string} markdown
 */
export function renderMarkdown(container, markdown) {
  const doc = container.ownerDocument;
  container.textContent = "";

  const lines = String(markdown ?? "").replace(/\r\n/g, "\n").split("\n");
  let i = 0;

  const paragraph = [];
  const flushParagraph = () => {
    if (!paragraph.length) return;
    const p = doc.createElement("p");
    renderInline(p, paragraph.join(" ").trim());
    container.appendChild(p);
    paragraph.length = 0;
  };

  while (i < lines.length) {
    const line = lines[i];

    // Fenced code. An unterminated fence still renders, rather than eating the reply.
    const fence = RULES.fence.exec(line);
    if (fence) {
      const body = [];
      i++;
      while (i < lines.length && !RULES.fence.test(lines[i])) body.push(lines[i++]);
      i++;
      flushParagraph();
      const pre = doc.createElement("pre");
      const code = doc.createElement("code");
      if (fence[1].trim()) code.className = `lang-${fence[1].trim().toLowerCase()}`;
      code.textContent = body.join("\n");
      pre.appendChild(code);
      container.appendChild(pre);
      continue;
    }

    if (!line.trim()) {
      flushParagraph();
      i++;
      continue;
    }

    if (RULES.divider.test(line)) {
      flushParagraph();
      container.appendChild(doc.createElement("hr"));
      i++;
      continue;
    }

    const heading = RULES.heading.exec(line);
    if (heading) {
      flushParagraph();
      // Headings inside a chat turn should not outrank the page; start at h4.
      const level = Math.min(6, 3 + heading[1].length);
      const node = doc.createElement(`h${level}`);
      renderInline(node, heading[2]);
      container.appendChild(node);
      i++;
      continue;
    }

    // Table: a pipe row followed by a divider row.
    if (RULES.tableRow.test(line) && RULES.tableDivider.test(lines[i + 1] ?? "")) {
      flushParagraph();
      const table = doc.createElement("table");
      const thead = doc.createElement("thead");
      const headRow = doc.createElement("tr");
      for (const cell of splitRow(line)) {
        const th = doc.createElement("th");
        renderInline(th, cell);
        headRow.appendChild(th);
      }
      thead.appendChild(headRow);
      table.appendChild(thead);

      const tbody = doc.createElement("tbody");
      i += 2;
      while (i < lines.length && RULES.tableRow.test(lines[i])) {
        const tr = doc.createElement("tr");
        for (const cell of splitRow(lines[i])) {
          const td = doc.createElement("td");
          renderInline(td, cell);
          tr.appendChild(td);
        }
        tbody.appendChild(tr);
        i++;
      }
      table.appendChild(tbody);

      const wrap = doc.createElement("div");
      wrap.className = "md-table";
      wrap.appendChild(table);
      container.appendChild(wrap);
      continue;
    }

    if (RULES.bullet.test(line) || RULES.ordered.test(line)) {
      flushParagraph();
      const ordered = RULES.ordered.test(line);
      const list = doc.createElement(ordered ? "ol" : "ul");
      if (ordered) {
        const start = Number(RULES.ordered.exec(line)[1]);
        if (start !== 1) list.setAttribute("start", String(start));
      }

      while (i < lines.length) {
        const match = ordered ? RULES.ordered.exec(lines[i]) : RULES.bullet.exec(lines[i]);
        if (!match) break;
        const li = doc.createElement("li");
        const parts = [ordered ? match[2] : match[1]];
        i++;
        // Continuation lines are indented and belong to the item above.
        while (i < lines.length && /^\s{2,}\S/.test(lines[i]) && !RULES.bullet.test(lines[i]) && !RULES.ordered.test(lines[i])) {
          parts.push(lines[i].trim());
          i++;
        }
        renderInline(li, parts.join(" "));
        list.appendChild(li);
      }
      container.appendChild(list);
      continue;
    }

    const quote = RULES.quote.exec(line);
    if (quote) {
      flushParagraph();
      const block = doc.createElement("blockquote");
      const parts = [quote[1]];
      i++;
      while (i < lines.length && RULES.quote.test(lines[i])) {
        parts.push(RULES.quote.exec(lines[i])[1]);
        i++;
      }
      renderInline(block, parts.join(" ").trim());
      container.appendChild(block);
      continue;
    }

    paragraph.push(line.trim());
    i++;
  }

  flushParagraph();
  return container;
}
