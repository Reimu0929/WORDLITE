// markdown.js
// WordPad Lite — Markdown <-> HTML 簡易変換

const escapeMap = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' };

export function escapeHtml(text) {
  return String(text).replace(/[&<>"']/g, (c) => escapeMap[c]);
}

/* ---------------- Inline parsing ---------------- */

const inlineRules = [
  {
    // ![alt](src)
    re: /(!\[[^\]]*\]\([^)]*\))/,
    toHtml: (match) => {
      const alt = match.slice(2, match.indexOf('(')).replace(/[\[\]]/g, '');
      const src = match.slice(match.indexOf('(') + 1, -1);
      return `<img src="${src}" alt="${alt}" />`;
    },
  },
  {
    // [text](url)
    re: /(\[[^\]]*\]\([^)]*\))/,
    toHtml: (match) => {
      const text = match.slice(1, match.indexOf('('));
      const href = match.slice(match.indexOf('(') + 1, -1);
      return `<a href="${href}">${text}</a>`;
    },
  },
  {
    // **bold**
    re: /(\*\*[^*]+\*\*)/,
    toHtml: (match) => `<strong>${match.slice(2, -2)}</strong>`,
  },
  {
    // *italic*
    re: /(\*[^*]+\*)/,
    toHtml: (match) => `<em>${match.slice(1, -1)}</em>`,
  },
  {
    // `code`
    re: /(`[^`]+`)/,
    toHtml: (match) => `<code>${match.slice(1, -1)}</code>`,
  },
];

function renderInline(text) {
  let html = escapeHtml(text);
  for (const rule of inlineRules) {
    html = html.replace(rule.re, rule.toHtml);
  }
  return html;
}

/* ---------------- Markdown -> HTML ---------------- */

export function parseMarkdown(text) {
  const lines = String(text).replace(/\r\n/g, '\n').split('\n');
  const out = [];
  let paragraph = [];

  function pushParagraph() {
    if (paragraph.length) {
      out.push(`<p>${paragraph.map(renderInline).join('<br>')}</p>`);
      paragraph = [];
    }
  }

  let listType = null; // ul | ol
  let listItems = [];

  function pushList() {
    if (listType && listItems.length) {
      out.push(`<${listType}>${listItems.map((item) => `<li>${renderInline(item)}</li>`).join('')}</${listType}>`);
      listType = null;
      listItems = [];
    }
  }

  for (const line of lines) {
    if (/^\s*$/.test(line)) {
      pushParagraph();
      pushList();
      continue;
    }

    // Heading
    const heading = line.match(/^(#{1,6})\s+(.*)/);
    if (heading) {
      pushParagraph();
      pushList();
      const level = heading[1].length;
      out.push(`<h${level}>${renderInline(heading[2])}</h${level}>`);
      continue;
    }

    // Blockquote
    if (/^\s*>\s?/.test(line)) {
      pushParagraph();
      pushList();
      out.push(`<blockquote>${renderInline(line.replace(/^\s*>\s?/, ''))}</blockquote>`);
      continue;
    }

    // Lists
    const bullet = line.match(/^\s*[-*+]\s+(.*)/);
    const ordered = line.match(/^\s*\d+[.)]\s+(.*)/);
    if (bullet || ordered) {
      pushParagraph();
      const type = bullet ? 'ul' : 'ol';
      if (listType !== type) pushList();
      listType = type;
      listItems.push(bullet ? bullet[1] : ordered[1]);
      continue;
    }

    paragraph.push(line);
  }

  pushParagraph();
  pushList();

  return out.join('\n') || '<p><br></p>';
}

/* ---------------- HTML -> Markdown ---------------- */

function serializeInline(runs) {
  let md = '';
  for (const node of runs) {
    if (node.nodeType === Node.TEXT_NODE) {
      md += node.nodeValue.replace(/\n/g, ' ');
    } else if (node.nodeType === Node.ELEMENT_NODE) {
      switch (node.tagName.toLowerCase()) {
        case 'strong':
        case 'b': {
          md += `**${serializeInline(node.childNodes).trim()}**`;
          break;
        }
        case 'em':
        case 'i': {
          md += `*${serializeInline(node.childNodes).trim()}*`;
          break;
        }
        case 'a': {
          const href = node.getAttribute('href') || '';
          md += `[${serializeInline(node.childNodes).trim()}](${href})`;
          break;
        }
        case 'img': {
          const src = node.getAttribute('src') || '';
          const alt = node.getAttribute('alt') || '';
          md += `![${alt}](${src})`;
          break;
        }
        case 'br': {
          md += '\n';
          break;
        }
        default:
          md += serializeInline(node.childNodes);
      }
    }
  }
  return md;
}

function serializeBlock(el) {
  switch (el.tagName.toLowerCase()) {
    case 'h1':
    case 'h2':
    case 'h3':
    case 'h4':
    case 'h5':
    case 'h6': {
      const level = Number(el.tagName[1]);
      return `${'#'.repeat(level)} ${serializeInline(el.childNodes).trim()}`;
    }
    case 'p':
    case 'div':
      return serializeInline(el.childNodes).trim();
    case 'ul':
    case 'ol': {
      const marker = el.tagName.toLowerCase();
      const items = [...el.children].map((li) => `${marker === 'ol' ? '1.' : '-'} ${serializeInline(li.childNodes).trim()}`);
      return items.join('\n');
    }
    case 'blockquote': {
      return `> ${serializeInline(el.childNodes).trim()}`;
    }
    case 'code':
      return `\`\`\`${serializeInline(el.childNodes).trim()}\`\`\``;
    case 'img': {
      const src = el.getAttribute('src') || '';
      const alt = el.getAttribute('alt') || '';
      return `![${alt}](${src})`;
    }
    default:
      break;
  }
}

function serializeTable(el) {
  const rows = el.querySelectorAll('tr');
  if (!rows.length) return '';
  const first = [...rows[0].children].map((td) => serializeInline(td.childNodes).trim());
  const header = `| ${first.join(' | ')} |`;
  const sep = `| ${first.map(() => '---').join(' | ')} |`;
  const body = [...rows].slice(1).map((row) => `| ${[...row.children].map((td) => serializeInline(td.childNodes).trim()).join(' | ')} |`);
  return [header, sep, ...body].join('\n');
}

export function serializeMarkdown(html) {
  const div = document.createElement('div');
  div.innerHTML = html;
  const blocks = [];

  for (const child of div.children) {
    if (!child.tagName) continue;
    if (/table/.test(child.tagName.toLowerCase())) blocks.push(serializeTable(child));
    else {
      const md = serializeBlock(child);
      if (md.trim() !== '') blocks.push(md);
    }
  }

  return blocks.join('\n\n');
}
