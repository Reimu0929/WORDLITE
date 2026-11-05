// rtf.js
// WordPad Lite — RTF -> HTML 簡易パーサ

const CONTROL_HEX_PAIRS = /[\da-fA-F]/;

const controlWordMap = {
  b: 'bold',
  i: 'italic',
  u: 'underline',
  ul: 'underline',
  strike: 'strike',
  par: 'par',
  pard: 'reset',
  plain: 'reset',
  line: 'line',
  tab: 'tab',
  emdash: emDash,
  endash: endDash,
  lquote: lQuote,
  rquote: rQuote,
  bullet: bullet,
  pict: 'ignore',
  object: 'ignore',
  fonttbl: 'ignore',
  colortbl: 'ignore',
  stylesheet: 'ignore',
  info: 'ignore',
  filetbl: 'ignore',
  header: 'ignore',
  footer: 'ignore',
};

function emDash() { return '\u2014'; }
function endDash() { return '\u2013'; }
function lQuote() { return '\u2018'; }
function rQuote() { return '\u2019'; }
function bullet() { return '\u2022'; }

function escapeHtml(text) {
  return text.replace(/[&<>"']/g, (c) => ({
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#39;',
  }[c]));
}

function defaultFmt() {
  return {
    bold: false,
    italic: false,
    underline: false,
    strike: false,
    color: null,
    fontFamily: null,
    fontSize: null,
  };
}

function cloneFmt(fmt) {
  return { ...fmt };
}

function runsToHtml(runs) {
  return runs.map((run) => {
    let html = escapeHtml(run.text);
    if (run.bold) html = `<strong>${html}</strong>`;
    if (run.italic) html = `<em>${html}</em>`;
    if (run.underline) html = `<u>${html}</u>`;
    if (run.strike) html = `<s>${html}</s>`;
    const styles = [];
    if (run.color) styles.push(`color: ${run.color}`);
    if (run.fontFamily) styles.push(`font-family: ${run.fontFamily}`);
    if (run.fontSize) styles.push(`font-size: ${run.fontSize}`);
    if (styles.length) html = `<span style="${styles.join('; ')}">${html}</span>`;
    return html;
  }).join('');
}

export function parseRtf(text) {
  const paragraphs = [];
  let currentRuns = [];
  let currentText = '';
  let fmt = defaultFmt();
  let ignored = false;
  const stack = [];

  let i = 0;
  const len = text.length;

  function flushText() {
    if (currentText) {
      currentRuns.push({
        text: currentText,
        ...fmt,
      });
      currentText = '';
    }
  }

  function closeParagraph() {
    flushText();
    if (currentRuns.length) {
      paragraphs.push(currentRuns);
      currentRuns = [];
    }
  }

  function openGroup() {
    stack.push({ fmt: cloneFmt(fmt), ignored });
  }

  function closeGroup() {
    const popped = stack.pop();
    if (popped) {
      fmt = popped.fmt;
      ignored = popped.ignored;
    }
  }

  function handleChar(c) {
    if (ignored) return;
    if (c === '\\') {
      // エスケープ文字は読み飛ばす
      i++;
      return;
    }
    if (c === '{' || c === '}') return; // グループ区切りは無視
    if (c === '\n' || c === '\r') return;
    currentText += c;
  }

  while (i < len) {
    const c = text[i];

    if (c === '{') {
      flushText();
      openGroup();
      i++;
      continue;
    }

    if (c === '}') {
      flushText();
      closeGroup();
      i++;
      continue;
    }

    if (c === '\\') {
      i++;
      const start = i;
      // 制御ワード成形
      let word = '';
      while (i < len && /[a-zA-Z]/.test(text[i])) {
        word += text[i];
        i++;
      }
      let param = '';
      if (text[i] === '-') {
        param += '-';
        i++;
      }
      while (i < len && /[0-9]/.test(text[i])) {
        param += text[i];
        i++;
      }
      if (text[i] === ' ') i++;

      if (!word) {
        // \ に制御ワードが続かない場合
        if (text[i] === '\'') {
          i++;
          if (i + 1 < len) {
            const hex = text.slice(i, i + 2);
            i += 2;
            try {
              const charCode = parseInt(hex, 16);
              if (!isNaN(charCode) && !ignored) {
                currentText += String.fromCharCode(charCode);
              }
            } catch {}
          }
          continue;
        }
        // 不明なエスケープは何もしない
        i++;
        continue;
      }

      switch (word) {
        case 'pict':
        case 'object':
        case 'fonttbl':
        case 'colortbl':
        case 'stylesheet':
        case 'info':
        case 'filetbl':
        case 'header':
        case 'footer':
          ignored = true;
          break;
        case 'b':
          fmt.bold = param !== '0';
          break;
        case 'i':
          fmt.italic = param !== '0';
          break;
        case 'u':
        case 'ul':
          fmt.underline = param !== '0';
          break;
        case 'strike':
          fmt.strike = param !== '0';
          break;
        case 'par':
          closeParagraph();
          break;
        case 'pard':
          fmt = defaultFmt();
          break;
        case 'plain':
          fmt = defaultFmt();
          break;
        case 'line':
        case 'tab':
          if (!ignored) currentText += ' ';
          break;
        default:
          break;
      }

      // `\\u` は特別扱い
      if (word === 'u' && param !== '0') {
        flushText();
        const code = parseInt(param, 10);
        if (!Number.isNaN(code) && !ignored) {
          currentText += String.fromCharCode(code);
        }
      }

      continue;
    }

    handleChar(c);
    i++;
  }

  closeParagraph();

  // 空ドキュメント対策
  if (!paragraphs.length) return '<p><br></p>';

  return paragraphs.map((runs) => `<p>${runsToHtml(runs)}</p>`).join('');
}

export function serializeRtf(html) {
  // 簡易 HTML -> RTF 変換（MVP）
  // HTML タグを剥がし、rich encoding は今後強化
  const div = document.createElement('div');
  div.innerHTML = html;
  const paragraphs = [...div.children];

  const lines = [];
  for (const para of paragraphs) {
    lines.push(para.innerText.replace(/\n/g, ' '));
  }
  const plain = lines.join('\n');

  const rtf = `{\\rtf1\\ansi ${plain.replace(/\r?\n/g, '\\par ')}\\par}`;
  return rtf;
}
