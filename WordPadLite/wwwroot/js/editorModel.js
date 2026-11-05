// editorModel.js
// WordPad Lite — Document Model (AST) 簡易実装

export const NodeTypes = {
  DOCUMENT: 'document',
  PARAGRAPH: 'paragraph',
  TABLE: 'table',
  TABLE_ROW: 'tr',
  TABLE_CELL: 'td',
  TEXT: 'text',
  IMAGE: 'image',
  LINK: 'link',
  LIST: 'list',
  LIST_ITEM: 'li',
  PAGE_BREAK: 'pageBreak',
};

export const BlockTypes = {
  PARAGRAPH: 'paragraph',
  HEADING: 'heading',
  TABLE: 'table',
  IMAGE: 'image',
  PAGE_BREAK: 'pageBreak',
};

export const InlineTypes = {
  TEXT: 'text',
  IMAGE: 'image',
  LINK: 'link',
};

/* ---------------- Escape / Unescape ---------------- */

export function escapeHtml(text) {
  return text.replace(/[&<>"']/g, (c) => ({
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#39;',
  }[c]));
}

/* ---------------- Inline parsing ---------------- */

function parseInlineElement(node, parentRun) {
  const runs = [];
  const tagName = node.tagName?.toLowerCase();

  for (const child of node.childNodes) {
    if (child.nodeType === Node.TEXT_NODE || child.nodeType === Node.COMMENT_NODE) {
      const text = child.nodeType === Node.TEXT_NODE ? child.nodeValue : '';
      if (text.trim() !== '') {
        runs.push({
          type: InlineTypes.TEXT,
          text,
          bold: parentRun?.bold ?? false,
          italic: parentRun?.italic ?? false,
          underline: parentRun?.underline ?? false,
          strike: parentRun?.strike ?? false,
          color: parentRun?.color ?? null,
          backgroundColor: parentRun?.backgroundColor ?? null,
          fontFamily: parentRun?.fontFamily ?? null,
          fontSize: parentRun?.fontSize ?? null,
          href: parentRun?.href ?? null,
        });
      }
    } else if (child.nodeType === Node.ELEMENT_NODE) {
      switch (tagName) {
        case 'b':
        case 'strong':
          runs.push(...parseInlineElement(child, { ...parentRun, bold: true }));
          break;
        case 'i':
        case 'em':
          runs.push(...parseInlineElement(child, { ...parentRun, italic: true }));
          break;
        case 'u':
          runs.push(...parseInlineElement(child, { ...parentRun, underline: true }));
          break;
        case 's':
        case 'strike':
          runs.push(...parseInlineElement(child, { ...parentRun, strike: true }));
          break;
        case 'a': {
          const href = node.getAttribute('href') || '';
          runs.push(...parseInlineElement(child, { ...parentRun, href }));
          break;
        }
        case 'span':
        case 'font': {
          const styles = node.style;
          const styleProps = {
            color: styles?.color || parentRun?.color || null,
            backgroundColor: styles?.backgroundColor || parentRun?.backgroundColor || null,
            fontFamily: styles?.fontFamily || parentRun?.fontFamily || null,
            fontSize: styles?.fontSize || parentRun?.fontSize || null,
          };
          runs.push(...parseInlineElement(child, { ...parentRun, ...styleProps }));
          break;
        }
        case 'img': {
          runs.push({
            type: InlineTypes.IMAGE,
            src: node.getAttribute('src') || '',
            alt: node.getAttribute('alt') || '',
          });
          break;
        }
        default: {
          // 未知のタグは単葉テキストとして扱う
          runs.push({
            type: InlineTypes.TEXT,
            text: node.textContent || '',
            ...parentRun,
          });
        }
      }
    }
  }

  return runs;
}

function parseBlockChildren(el) {
  const children = [];

  for (const child of el.childNodes) {
    if (child.nodeType === Node.TEXT_NODE) {
      children.push({
        type: InlineTypes.TEXT,
        text: child.nodeValue,
        bold: false,
        italic: false,
        underline: false,
        strike: false,
      });
    } else if (child.nodeType === Node.ELEMENT_NODE) {
      children.push(...parseInlineElement(child, {}));
    }
  }

  return children;
}

/* ---------------- Block parsing ---------------- */

const TAG_INLINE_MAP = {
  p: BlockTypes.PARAGRAPH,
  h1: BlockTypes.HEADING,
  h2: BlockTypes.HEADING,
  h3: BlockTypes.HEADING,
  h4: BlockTypes.HEADING,
};

function parseTable(el) {
  const rows = [];
  for (const row of el.querySelectorAll('tr')) {
    const cells = [...row.children].map((cell) => ({
      type: NodeTypes.TABLE_CELL,
      children: parseBlockChildren(cell),
    }));
    rows.push({ type: NodeTypes.TABLE_ROW, cells });
  }
  return { type: BlockTypes.TABLE, rows };
}

export function parseDocument(rootEl) {
  const blocks = [];

  if (!rootEl) {
    return { type: NodeTypes.DOCUMENT, blocks };
  }

  for (const child of rootEl.children) {
    const tag = child.tagName?.toLowerCase();
    if (!tag) continue;

    if (tag === 'table') {
      blocks.push(parseTable(child));
    } else if (tag === 'ul' || tag === 'ol') {
      const items = [...child.children].map((li) => ({
        type: NodeTypes.LIST_ITEM,
        children: parseBlockChildren(li),
      }));
      blocks.push({ type: BlockTypes.PARAGRAPH, list: { kind: tag, items } });
    } else if (TAG_INLINE_MAP[tag]) {
      const level = Number(tag.slice(1)) || 0;
      blocks.push({
        type: TAG_INLINE_MAP[tag] === BlockTypes.HEADING ? BlockTypes.HEADING : BlockTypes.PARAGRAPH,
        level,
        align: child.style?.textAlign || null,
        children: parseBlockChildren(child),
      });
    } else {
      // div / 未知タグは段落として扱う
      blocks.push({
        type: BlockTypes.PARAGRAPH,
        align: child.style?.textAlign || null,
        children: parseBlockChildren(child),
      });
    }
  }

  return { type: NodeTypes.DOCUMENT, blocks };
}

/* ---------------- Inline serialization ---------------- */

function serializeInlineRun(run) {
  let html = escapeHtml(run.text);

  if (run.bold) html = `<strong>${html}</strong>`;
  if (run.italic) html = `<em>${html}</em>`;
  if (run.underline) html = `<u>${html}</u>`;
  if (run.strike) html = `<s>${html}</s>`;
  if (run.href) html = `<a href="${escapeHtml(run.href)}">${html}</a>`;

  const styles = [];
  if (run.color) styles.push(`color: ${run.color}`);
  if (run.backgroundColor) styles.push(`background-color: ${run.backgroundColor}`);
  if (run.fontFamily) styles.push(`font-family: ${run.fontFamily}`);
  if (run.fontSize) styles.push(`font-size: ${run.fontSize}`);
  if (styles.length) {
    html = `<span style="${styles.join('; ')}">${html}</span>`;
  }

  return html;
}

function serializeInlineRuns(runs) {
  return runs.map(serializeInlineRun).join('');
}

/* ---------------- Block serialization ---------------- */

function serializeTableCell(cell) {
  return `<td>${serializeInlineRuns(cell.children)}</td>`;
}

function serializeTableRow(row) {
  return `<tr>${row.cells.map(serializeTableCell).join('')}</tr>`;
}

export function serializeBlock(block) {
  if (block.type === BlockTypes.HEADING) {
    const tag = `h${block.level ?? 1}`;
    return `<${tag}>${serializeInlineRuns(block.children ?? [])}</${tag}>`;
  }

  if (block.type === BlockTypes.TABLE) {
    return `<table><tbody>${block.rows.map(serializeTableRow).join('')}</tbody></table>`;
  }

  // PARAGRAPH / list
  if (block.list) {
    const tag = block.list.kind === 'ol' ? 'ol' : 'ul';
    return `<${tag}>${block.list.items.map((li) => `<li>${serializeInlineRuns(li.children)}</li>`).join('')}</${tag}>`;
  }

  const alignStyle = block.align ? ` style="text-align: ${block.align}"` : '';
  return `<p${alignStyle}>${serializeInlineRuns(block.children ?? [])}</p>`;
}

export function serializeDocument(doc) {
  return doc.blocks.map(serializeBlock).join('');
}

/* ---------------- AST <-> DOM ASCII preview ---------------- */

export function debugAst(rootEl) {
  const ast = parseDocument(rootEl);
  return JSON.stringify(ast, null, 2);
}
