import { UndoStack } from './undoStack.js';
import { parseRtf } from './rtf.js';
import { parseMarkdown, serializeMarkdown } from './markdown.js';

const $ = (sel) => document.querySelector(sel);
const $$ = (sel) => document.querySelectorAll(sel);
const byId = (id) => document.getElementById(id);

const undoStack = new UndoStack(1000);
let isRestoring = false;
let lastPushAt = 0;

const STORAGE_KEY = 'wpl-settings';
const PAGE_SIZES = {
  A4: [210, 297],
  B5: [182, 257],
  Letter: [215.9, 279.4],
  Legal: [215.9, 355.6],
};

const state = {
  zoom: 100,
  theme: 'system',
  fileName: '無題',
  modified: false,
  showRuler: true,
  showMarks: false,
  defaultFormat: 'docx',
  pageSize: 'A4',
  orientation: 'portrait',
  margins: { top: 25, right: 25, bottom: 25, left: 25 },
  headerText: '',
  footerText: '',
  showPageNumbers: false,
  readMode: false,
  ribbonCollapsed: false,
};

const editor = byId('editor');
const page = byId('page');
const titleText = $('.title-bar__text');
let lastEditorRange = null;

/* ---------- Selection helpers ---------- */
function getSelectionOffsets() {
  const sel = window.getSelection();
  if (!sel || !sel.rangeCount) return null;
  const range = sel.getRangeAt(0);
  return {
    start: getGlobalOffset(range.startContainer, range.startOffset),
    end: getGlobalOffset(range.endContainer, range.endOffset),
  };
}

function getGlobalOffset(node, offsetInNode) {
  if (!node || !editor.contains(node)) return 0;
  const walker = document.createTreeWalker(editor, NodeFilter.SHOW_TEXT);
  let total = 0;
  let current;
  while ((current = walker.nextNode())) {
    if (current === node) return total + offsetInNode;
    total += current.length;
  }
  return total;
}

function restoreSelection(offsets) {
  if (!offsets) return;
  const ee = offsetAt(offsets.end ?? offsets.start);
  const ss = offsetAt(offsets.start);
  if (!ss || !ee) return;

  const range = document.createRange();
  range.setStart(ss.node, ss.offset);
  range.setEnd(ee.node, ee.offset);
  const sel = window.getSelection();
  sel.removeAllRanges();
  sel.addRange(range);
}

function offsetAt(globalOffset) {
  const walker = document.createTreeWalker(editor, NodeFilter.SHOW_TEXT);
  let total = 0;
  let current;
  while ((current = walker.nextNode())) {
    const len = current.length;
    if (globalOffset <= total + len) {
      return { node: current, offset: Math.max(0, globalOffset - total) };
    }
    total += len;
  }
  return current ? { node: current, offset: current.length } : null;
}

/* ---------- Undo / Redo state ---------- */
function updateHistoryButtons() {
  const undoBtn = byId('qat-undo');
  const redoBtn = byId('qat-redo');
  if (undoBtn) undoBtn.disabled = !undoStack.canUndo;
  if (redoBtn) redoBtn.disabled = !undoStack.canRedo;
}

function pushState(coalesce = false) {
  if (isRestoring) return;
  const now = Date.now();
  const snapshot = {
    html: editor.innerHTML,
    selection: getSelectionOffsets(),
  };

  if (coalesce && lastPushAt > 0 && now - lastPushAt < 700 && undoStack.canUndo) {
    undoStack.replaceTop(snapshot);
  } else {
    undoStack.push(snapshot);
  }

  lastPushAt = coalesce ? now : 0;
  updateHistoryButtons();
}

function applyState(snapshot) {
  isRestoring = true;
  editor.innerHTML = snapshot.html;
  restoreSelection(snapshot.selection);
  isRestoring = false;
  updateStatus();
  updateRibbonState();
  updateHistoryButtons();
}

function undo() {
  if (!undoStack.undo(applyState)) return;
  lastPushAt = 0;
  editor.focus();
  updateHistoryButtons();
}

function redo() {
  if (!undoStack.redo(applyState)) return;
  lastPushAt = 0;
  editor.focus();
  updateHistoryButtons();
}

function resetUndoStack() {
  isRestoring = true;
  undoStack.reset({ html: editor.innerHTML, selection: getSelectionOffsets() });
  isRestoring = false;
  lastPushAt = 0;
  updateHistoryButtons();
}

/* ---------- Commands ---------- */
const commands = [
  { id: 'bold', label: '太字', run: () => exec('bold'), shortcut: 'Ctrl+B' },
  { id: 'italic', label: '斜体', run: () => exec('italic'), shortcut: 'Ctrl+I' },
  { id: 'underline', label: '下線', run: () => exec('underline'), shortcut: 'Ctrl+U' },
  { id: 'strikeThrough', label: '取り消し線', run: () => exec('strikeThrough'), shortcut: '' },
  { id: 'justifyLeft', label: '左揃え', run: () => exec('justifyLeft'), shortcut: 'Ctrl+L' },
  { id: 'justifyCenter', label: '中央揃え', run: () => exec('justifyCenter'), shortcut: 'Ctrl+E' },
  { id: 'justifyRight', label: '右揃え', run: () => exec('justifyRight'), shortcut: 'Ctrl+R' },
  { id: 'justifyFull', label: '両端揃え', run: () => exec('justifyFull'), shortcut: 'Ctrl+J' },
  { id: 'insertUnorderedList', label: '箇条書き', run: () => exec('insertUnorderedList'), shortcut: '' },
  { id: 'insertOrderedList', label: '番号付け', run: () => exec('insertOrderedList'), shortcut: '' },
  { id: 'cut', label: '切り取り', run: () => exec('cut'), shortcut: 'Ctrl+X' },
  { id: 'copy', label: 'コピー', run: () => exec('copy'), shortcut: 'Ctrl+C' },
  { id: 'paste', label: 'ペースト', run: () => exec('paste'), shortcut: 'Ctrl+V' },
  { id: 'undo', label: '取消', run: () => exec('undo'), shortcut: 'Ctrl+Z' },
  { id: 'redo', label: 'やり直し', run: () => exec('redo'), shortcut: 'Ctrl+Y' },
  { id: 'selectAll', label: 'すべて選択', run: () => exec('selectAll'), shortcut: 'Ctrl+A' },
  { id: 'removeFormat', label: '書式クリア', run: () => exec('removeFormat'), shortcut: '' },
  { id: 'new', label: '新規作成', run: newDocument, shortcut: 'Ctrl+N' },
  { id: 'open', label: '開く', run: openDocument, shortcut: 'Ctrl+O' },
  { id: 'save', label: '保存', run: saveDocument, shortcut: 'Ctrl+S' },
  { id: 'saveAs', label: '名前を付けて保存', run: saveAs, shortcut: 'Ctrl+Shift+S' },
  { id: 'print', label: '印刷', run: () => window.print(), shortcut: 'Ctrl+P' },
  { id: 'find', label: '検索', run: () => showDialog('find-dialog'), shortcut: 'Ctrl+F' },
  { id: 'replace', label: '置換', run: () => showDialog('find-dialog'), shortcut: 'Ctrl+H' },
  { id: 'insertImage', label: '画像挿入', run: insertImage, shortcut: 'Insert' },
  { id: 'insertTable', label: 'テーブル挿入', run: () => byId('table-picker').focus(), shortcut: '' },
  { id: 'insertLink', label: 'リンク挿入', run: showLinkDialog, shortcut: 'Ctrl+K' },
  { id: 'zoomIn', label: '拡大', run: () => setZoom(state.zoom + 10), shortcut: '' },
  { id: 'zoomOut', label: '縮小', run: () => setZoom(state.zoom - 10), shortcut: '' },
  { id: 'zoomReset', label: '100% に戻す', run: () => setZoom(100), shortcut: '' },
  { id: 'toggleRuler', label: '定規の表示切替', run: toggleRuler, shortcut: '' },
  { id: 'toggleDark', label: 'ダークモード', run: () => setTheme('dark'), shortcut: '' },
  { id: 'toggleLight', label: 'ライトモード', run: () => setTheme('light'), shortcut: '' },
];

function exec(command, value = null) {
  editor.focus();
  if (command === 'undo') { undo(); return; }
  if (command === 'redo') { redo(); return; }
  try {
    document.execCommand('styleWithCSS', false, 'true');
    document.execCommand(command, false, value);
    pushState();
  } catch (e) {
    console.warn('execCommand failed', command, e);
  }
  updateStatus();
  updateRibbonState();
}

function applyInlineStyle(prop, value) {
  const sel = window.getSelection();
  if (!sel || !sel.rangeCount) return;
  const range = sel.getRangeAt(0);
  if (!editor.contains(range.commonAncestorContainer)) return;

  const span = document.createElement('span');
  span.style[prop] = value;

  if (range.collapsed) {
    // 選択なし: これから入力する文字に書式を適用する
    span.appendChild(document.createTextNode('\u200B'));
    range.insertNode(span);
    const caret = document.createRange();
    caret.setStart(span.firstChild, 0);
    caret.collapse(true);
    sel.removeAllRanges();
    sel.addRange(caret);
  } else {
    // 複数ノードにまたがる選択でも安全にラップする
    try {
      const fragment = range.extractContents();
      span.appendChild(fragment);
      range.insertNode(span);
      const selected = document.createRange();
      selected.selectNodeContents(span);
      sel.removeAllRanges();
      sel.addRange(selected);
    } catch (e) {
      console.warn('inline style failed', e);
      exec('styleWithCSS', false);
      return;
    }
  }

  editor.focus();
  pushState();
  updateRibbonState();
}

function getSelectedBlocks() {
  const sel = window.getSelection();
  let range = sel && sel.rangeCount ? sel.getRangeAt(0) : null;

  if (!range || !editor.contains(range.commonAncestorContainer)) {
    range = lastEditorRange;
  }
  if (!range || !editor.contains(range.commonAncestorContainer)) return [];

  const blocks = [];
  editor.querySelectorAll('p, div, h1, h2, h3, h4, h5, h6, li, blockquote').forEach((block) => {
    if (range.intersectsNode(block)) blocks.push(block);
  });

  if (blocks.length === 0) {
    const node = range.commonAncestorContainer.nodeType === Node.ELEMENT_NODE
      ? range.commonAncestorContainer
      : range.commonAncestorContainer.parentElement;
    const closest = node?.closest('p, div, h1, h2, h3, h4, h5, h6, li, blockquote');
    if (closest && editor.contains(closest)) blocks.push(closest);
  }

  return blocks;
}

function applyParagraphSpacing(side, rawValue) {
  const points = Math.max(0, Math.min(72, Number(rawValue) || 0));
  const blocks = getSelectedBlocks();
  if (blocks.length === 0) return;

  const property = side === 'before' ? 'marginTop' : 'marginBottom';
  for (const block of blocks) {
    block.style[property] = points > 0 ? `${points}pt` : '';
  }

  const selection = window.getSelection();
  if (selection && lastEditorRange) {
    selection.removeAllRanges();
    selection.addRange(lastEditorRange);
  }
  editor.focus();
  pushState();
  updateRibbonState();
  updateStatus();
}

/* ---------- Ribbon tabs ---------- */
function initTabs() {
  $$('.ribbon-tab').forEach((tab) => {
    tab.addEventListener('click', () => {
      $$('.ribbon-tab').forEach((t) => t.classList.remove('active'));
      tab.classList.add('active');
      const pageName = tab.dataset.tab;
      $$('.ribbon-page').forEach((p) => p.classList.toggle('active', p.dataset.page === pageName));
    });
  });
}

/* ---------- Ribbon command delegation ---------- */
function initRibbon() {
  $('.ribbon-panel').addEventListener('click', (e) => {
    const btn = e.target.closest('[data-cmd]');
    if (!btn) return;
    e.preventDefault();
    exec(btn.dataset.cmd);
  });

  byId('font-family').addEventListener('change', (e) => {
    applyInlineStyle('fontFamily', e.target.value);
  });

  byId('font-size').addEventListener('change', (e) => {
    applyInlineStyle('fontSize', `${e.target.value}pt`);
  });

  byId('paragraph-before').addEventListener('change', (e) => {
    applyParagraphSpacing('before', e.target.value);
  });
  byId('paragraph-after').addEventListener('change', (e) => {
    applyParagraphSpacing('after', e.target.value);
  });

  byId('font-color').addEventListener('change', (e) => {
    applyInlineStyle('color', e.target.value);
  });

  byId('back-color').addEventListener('change', (e) => {
    applyInlineStyle('backgroundColor', e.target.value);
  });

  $$('.line-spacing-select').forEach((select) => {
    select.addEventListener('change', (e) => {
      editor.style.lineHeight = e.target.value;
      $$('.line-spacing-select').forEach((other) => { other.value = e.target.value; });
      pushState();
    });
  });

  $$('.style-btn').forEach((btn) => {
    btn.addEventListener('click', () => {
      const tag = btn.dataset.style === 'p' ? 'P' : btn.dataset.style.toUpperCase();
      exec('formatBlock', tag);
    });
  });

  byId('zoom-in').addEventListener('click', () => setZoom(state.zoom + 10));
  byId('zoom-out').addEventListener('click', () => setZoom(state.zoom - 10));
  byId('zoom-reset').addEventListener('click', () => setZoom(100));
  byId('zoom-slider').addEventListener('input', (e) => setZoom(parseInt(e.target.value, 10)));

  byId('status-zoom-in').addEventListener('click', () => setZoom(state.zoom + 10));
  byId('status-zoom-out').addEventListener('click', () => setZoom(state.zoom - 10));
  byId('status-zoom-slider').addEventListener('input', (e) => setZoom(parseInt(e.target.value, 10)));

  byId('show-ruler').addEventListener('change', (e) => setRuler(e.target.checked));
  byId('show-marks').addEventListener('change', (e) => {
    state.showMarks = e.target.checked;
    editor.classList.toggle('show-marks', state.showMarks);
  });

  $$('input[name="theme"]').forEach((radio) => {
    radio.addEventListener('change', (e) => setTheme(e.target.value));
  });

  byId('btn-insert-image').addEventListener('click', insertImage);
  byId('btn-insert-image-file').addEventListener('click', () => byId('image-input').click());
  byId('image-input').addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (file) insertImageFromFile(file);
    e.target.value = '';
  });
  byId('btn-insert-link').addEventListener('click', showLinkDialog);
  byId('btn-insert-pagebreak').addEventListener('click', insertPageBreak);
  byId('btn-insert-blankpage').addEventListener('click', insertBlankPage);

  byId('file-new').addEventListener('click', newDocument);
  byId('file-open').addEventListener('click', openDocument);
  byId('file-save').addEventListener('click', saveDocument);
  byId('file-save-as').addEventListener('click', saveAs);
  byId('file-export-docx').addEventListener('click', async () => {
    try {
      await exportDocx();
    } catch (error) {
      alert(error.message || 'DOCX への書き出しに失敗しました。');
    }
  });
  byId('file-export-pdf').addEventListener('click', async () => {
    try {
      await exportPdf();
    } catch (error) {
      alert(error.message || 'PDF への書き出しに失敗しました。');
    }
  });
  byId('file-print').addEventListener('click', () => window.print());
  byId('file-properties').addEventListener('click', showProperties);
  byId('btn-search').addEventListener('click', () => showDialog('find-dialog'));
  byId('btn-find').addEventListener('click', () => showDialog('find-dialog'));
  byId('btn-replace').addEventListener('click', () => showDialog('find-dialog'));

  byId('qat-save').addEventListener('click', saveDocument);
  byId('qat-undo').addEventListener('click', undo);
  byId('qat-redo').addEventListener('click', redo);
  byId('btn-command-palette').addEventListener('click', () => window.showCommandPalette());
  byId('btn-ribbon-collapse').addEventListener('click', toggleRibbon);
  byId('btn-read-mode').addEventListener('click', toggleReadMode);
  byId('btn-exit-read').addEventListener('click', () => toggleReadMode(false));
  byId('btn-share').addEventListener('click', shareDocument);

  byId('btn-maximize').addEventListener('click', async () => {
    try {
      if (!document.fullscreenElement) await document.documentElement.requestFullscreen();
      else await document.exitFullscreen();
    } catch {
      // フルスクリーンが許可されない環境では何もしない
    }
  });
  byId('btn-close').addEventListener('click', () => {
    if (!state.modified || confirm('変更を破棄して閉じますか？')) window.close();
  });
}

function updateRibbonState() {
  $$('.ribbon-btn[data-cmd]').forEach((btn) => {
    try {
      const active = document.queryCommandState(btn.dataset.cmd);
      btn.classList.toggle('active', active);
    } catch {
      btn.classList.remove('active');
    }
  });
  try {
    const font = document.queryCommandValue('fontName').replace(/["']/g, '');
    const opt = Array.from(byId('font-family').options).find((o) => o.value.toLowerCase() === font.toLowerCase());
    if (opt) byId('font-family').value = opt.value;
  } catch {}

  const activeId = document.activeElement?.id;
  if (activeId !== 'paragraph-before' && activeId !== 'paragraph-after') {
    const block = getSelectedBlocks()[0];
    if (block) {
      byId('paragraph-before').value = Math.round(parseFloat(block.style.marginTop) || 0);
      byId('paragraph-after').value = Math.round(parseFloat(block.style.marginBottom) || 0);
    }
  }
}

/* ---------- Zoom / Ruler / Theme ---------- */
function setZoom(value) {
  state.zoom = Math.max(50, Math.min(200, value));
  page.style.zoom = `${state.zoom}%`;
  byId('zoom-slider').value = state.zoom;
  byId('status-zoom-slider').value = state.zoom;
  byId('zoom-reset').textContent = `${state.zoom}%`;
  byId('status-zoom-text').textContent = `${state.zoom}%`;
  saveSettings();
}

function setRuler(show) {
  state.showRuler = show;
  byId('ruler-h').classList.toggle('hidden', !show);
  byId('ruler-v').classList.toggle('hidden', !show);
  byId('show-ruler').checked = show;
  saveSettings();
}

function setTheme(theme) {
  state.theme = theme;
  document.body.dataset.theme = theme === 'system' ? (matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light') : theme;
  const radio = $(`input[name="theme"][value="${theme}"]`);
  if (radio) radio.checked = true;
  saveSettings();
}

function applySystemTheme() {
  if (state.theme === 'system') {
    document.body.dataset.theme = matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light';
  }
}

function toggleRuler() {
  setRuler(!state.showRuler);
}

/* ---------- Page layout ---------- */
function pageDimensions() {
  const [w, h] = PAGE_SIZES[state.pageSize] ?? PAGE_SIZES.A4;
  return state.orientation === 'landscape' ? [h, w] : [w, h];
}

function getPageHeightPx() {
  return (pageDimensions()[1] / 25.4) * 96;
}

function applyPageSettings() {
  const [w, h] = pageDimensions();
  const root = document.documentElement.style;
  root.setProperty('--page-width', `${w}mm`);
  root.setProperty('--page-height', `${h}mm`);
  root.setProperty('--page-margin-top', `${state.margins.top}mm`);
  root.setProperty('--page-margin-right', `${state.margins.right}mm`);
  root.setProperty('--page-margin-bottom', `${state.margins.bottom}mm`);
  root.setProperty('--page-margin-left', `${state.margins.left}mm`);

  if (byId('page-size')) byId('page-size').value = state.pageSize;
  if (byId('page-orientation')) byId('page-orientation').value = state.orientation;
  for (const side of ['top', 'right', 'bottom', 'left']) {
    const input = byId(`margin-${side}`);
    if (input) input.value = state.margins[side];
  }
  applyHeaderFooter();
  saveSettings();
  updateStatus();
}

function applyHeaderFooter() {
  const header = byId('page-header');
  const footer = byId('page-footer');
  const footerText = byId('page-footer-text');
  const headerInput = byId('header-text');
  const footerInput = byId('footer-text');
  const pageNumber = byId('footer-page-number');

  if (headerInput) headerInput.value = state.headerText;
  if (footerInput) footerInput.value = state.footerText;
  if (pageNumber) pageNumber.checked = state.showPageNumbers;

  if (header) {
    header.textContent = state.headerText;
    header.classList.toggle('visible', state.headerText.trim() !== '');
  }
  if (footer && footerText) {
    footerText.textContent = state.footerText;
    footer.classList.toggle('visible', state.footerText.trim() !== '' || state.showPageNumbers);
  }
  updateFooterPageNumber();
}

function updateFooterPageNumber() {
  const number = byId('page-footer-number');
  if (!number) return;
  if (!state.showPageNumbers) {
    number.textContent = '';
    return;
  }
  const pages = Math.max(1, Math.ceil((page.scrollHeight - 1) / getPageHeightPx()));
  number.textContent = `1 / ${pages}`;
}

function initPageSettings() {
  byId('page-size').addEventListener('change', (e) => {
    state.pageSize = e.target.value;
    applyPageSettings();
  });
  byId('page-orientation').addEventListener('change', (e) => {
    state.orientation = e.target.value;
    applyPageSettings();
  });

  for (const side of ['top', 'right', 'bottom', 'left']) {
    byId(`margin-${side}`).addEventListener('change', (e) => {
      const value = Math.max(1, Math.min(50, Number(e.target.value) || 25));
      state.margins[side] = value;
      applyPageSettings();
    });
  }

  byId('header-text').addEventListener('input', (e) => {
    state.headerText = e.target.value;
    applyHeaderFooter();
    saveSettings();
  });
  byId('footer-text').addEventListener('input', (e) => {
    state.footerText = e.target.value;
    applyHeaderFooter();
    saveSettings();
  });
  byId('footer-page-number').addEventListener('change', (e) => {
    state.showPageNumbers = e.target.checked;
    applyHeaderFooter();
    saveSettings();
  });
}

/* ---------- View modes ---------- */
function toggleRibbon(force) {
  state.ribbonCollapsed = typeof force === 'boolean' ? force : !state.ribbonCollapsed;
  byId('ribbon').classList.toggle('collapsed', state.ribbonCollapsed);
  const use = byId('btn-ribbon-collapse').querySelector('use');
  if (use) use.setAttribute('href', state.ribbonCollapsed ? '#i-chevron-down' : '#i-chevron-up');
  saveSettings();
}

function toggleReadMode(force) {
  state.readMode = typeof force === 'boolean' ? force : !state.readMode;
  document.body.classList.toggle('read-mode', state.readMode);
  if (state.readMode) {
    editor.focus();
  }
  saveSettings();
}

async function shareDocument() {
  const text = editor.innerText || '';
  try {
    if (navigator.share) {
      await navigator.share({ title: state.fileName, text });
    } else if (navigator.clipboard) {
      await navigator.clipboard.writeText(text);
      alert('本文をクリップボードにコピーしました');
    } else {
      alert('このブラウザでは共有機能を利用できません');
    }
  } catch {
    // キャンセル時は何もしない
  }
}

/* ---------- Editor events ---------- */
function initEditor() {
  editor.addEventListener('input', () => {
    state.modified = true;
    updateTitle();
    updateStatus();
    pushState(true);
  });

  document.addEventListener('selectionchange', () => {
    const selection = window.getSelection();
    if (selection && selection.rangeCount && editor.contains(selection.anchorNode)) {
      lastEditorRange = selection.getRangeAt(0).cloneRange();
    }
    updateRibbonState();
    updateStatus();
  });

  editor.addEventListener('compositionstart', () => {
    byId('ime-status').textContent = '変換中';
  });
  editor.addEventListener('compositionend', () => {
    byId('ime-status').textContent = '日本語';
    updateStatus();
  });

  editor.addEventListener('keydown', (e) => {
    if (e.key === 'Tab') {
      e.preventDefault();
      exec(e.shiftKey ? 'outdent' : 'indent');
    }
  });

  // Drag & drop files
  document.addEventListener('dragover', (e) => e.preventDefault());
  document.addEventListener('drop', (e) => {
    e.preventDefault();
    const file = e.dataTransfer.files[0];
    if (!file) return;
    if (file.type.startsWith('image/')) insertImageFromFile(file);
    else loadFile(file);
  });
}

function updateStatus() {
  const text = editor.innerText || '';
  const chars = text.replace(/\n/g, '').length;
  const lines = text.trim() === '' ? 1 : text.split(/\n/).length;
  const sel = window.getSelection();
  const selChars = sel ? sel.toString().length : 0;
  const pages = Math.max(1, Math.ceil((page.scrollHeight - 1) / getPageHeightPx()));

  byId('status-page').textContent = `ページ 1 / ${pages}`;
  byId('status-line').textContent = `行 ${lines}`;
  byId('status-char').textContent = `文字数 ${chars}`;
  byId('status-sel').textContent = `選択 ${selChars}文字`;
  updateFooterPageNumber();
}

function updateTitle() {
  const marker = state.modified ? '● ' : '';
  titleText.textContent = `${marker}${state.fileName} — WordPad Lite`;
}

/* ---------- File I/O ---------- */
const fileInput = byId('file-input');
fileInput.addEventListener('change', (e) => {
  const file = e.target.files[0];
  if (file) loadFile(file);
  fileInput.value = '';
});

function loadFile(file) {
  if (state.modified && !confirm('変更を破棄してファイルを開きますか？')) return;
  const ext = file.name.split('.').pop()?.toLowerCase();
  if (ext === 'docx') {
    importDocx(file);
    return;
  }
  if (ext === 'pdf') {
    importPdf(file);
    return;
  }

  const reader = new FileReader();
  reader.onload = () => {
    if (ext === 'html' || ext === 'htm') {
      editor.innerHTML = reader.result;
    } else if (ext === 'rtf') {
      editor.innerHTML = parseRtf(reader.result);
    } else if (ext === 'md' || ext === 'markdown') {
      editor.innerHTML = parseMarkdown(reader.result);
    } else {
      const escaped = escapeHtml(reader.result);
      editor.innerHTML = escaped.split(/\n\n+/).map((p) => `<p>${p.replace(/\n/g, '<br>')}</p>`).join('');
    }
    finishFileLoad(file);
  };
  reader.readAsText(file);
}

async function importDocx(file) {
  try {
    const response = await fetch('/api/docx/import', {
      method: 'POST',
      headers: { 'Content-Type': 'application/octet-stream' },
      body: file,
    });
    if (!response.ok) {
      let message = 'DOCX を読み込めませんでした。';
      try {
        const data = await response.json();
        if (data.error) message = data.error;
      } catch {}
      throw new Error(message);
    }
    const data = await response.json();
    editor.innerHTML = data.html || '<p><br></p>';
    finishFileLoad(file);
  } catch (error) {
    alert(error.message || 'DOCX を読み込めませんでした。');
  }
}

async function importPdf(file) {
  try {
    const response = await fetch('/api/pdf/import', {
      method: 'POST',
      headers: { 'Content-Type': 'application/octet-stream' },
      body: file,
    });
    if (!response.ok) {
      let message = 'PDF を読み込めませんでした。';
      try {
        const data = await response.json();
        if (data.error) message = data.error;
      } catch {}
      throw new Error(message);
    }
    const data = await response.json();
    editor.innerHTML = data.html || '<p><br></p>';
    finishFileLoad(file);
  } catch (error) {
    alert(error.message || 'PDF を読み込めませんでした。');
  }
}

function finishFileLoad(file) {
  state.fileName = file.name.replace(/\.[^/.]+$/, '');
  state.modified = false;
  lastEditorRange = null;
  resetHeaderFooter();
  updateTitle();
  updateStatus();
  resetUndoStack();
}

function newDocument() {
  if (state.modified && !confirm('変更を破棄して新規作成しますか？')) return;
  editor.innerHTML = '<p><br></p>';
  state.fileName = '無題';
  state.modified = false;
  lastEditorRange = null;
  resetHeaderFooter();
  updateTitle();
  updateStatus();
  resetUndoStack();
}

function resetHeaderFooter() {
  state.headerText = '';
  state.footerText = '';
  state.showPageNumbers = false;
  applyHeaderFooter();
}

function openDocument() {
  fileInput.click();
}

async function saveDocument() {
  try {
    await downloadFile(state.defaultFormat);
    state.modified = false;
    updateTitle();
  } catch (error) {
    alert(error.message || '保存に失敗しました。');
  }
}

function saveAs() {
  const fmt = (prompt('保存形式を入力してください (docx / pdf / html / rtf / md / txt):', state.defaultFormat) || '').trim().toLowerCase();
  if (!fmt) return;
  state.defaultFormat = ['docx', 'pdf', 'html', 'rtf', 'md', 'txt'].includes(fmt) ? fmt : 'docx';
  saveSettings();
  saveDocument();
}

async function downloadFile(format) {
  if (format === 'docx') {
    await exportDocx();
    return;
  }
  if (format === 'pdf') {
    await exportPdf();
    return;
  }

  const isHtml = format === 'html';
  const isRtf = format === 'rtf';
  const isMarkdown = format === 'md';
  const content = isHtml
    ? `<!DOCTYPE html><html><head><meta charset="utf-8"><title>${state.fileName}</title></head><body>${editor.innerHTML}</body></html>`
    : isRtf
      ? `{\rtf1\ansi ${plainFromHtml(editor.innerHTML).replace(/\r?\n/g, '\par ')}\par}`
      : isMarkdown
        ? serializeMarkdown(editor.innerHTML)
        : editor.innerText;
  const mimeType = isHtml ? 'text/html;charset=utf-8'
    : isRtf ? 'text/rtf;charset=utf-8'
    : isMarkdown ? 'text/markdown;charset=utf-8'
    : 'text/plain;charset=utf-8';
  const extension = isHtml ? 'html' : isRtf ? 'rtf' : isMarkdown ? 'md' : 'txt';
  const blob = new Blob([content], { type: mimeType });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = `${state.fileName}.${extension}`;
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 1000);
}

async function exportDocx() {
  const blob = await exportServerDocument('/api/docx/export', 'DOCX への書き出しに失敗しました。');
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = `${state.fileName}.docx`;
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 1000);
}

async function exportPdf() {
  const blob = await exportServerDocument('/api/pdf/export', 'PDF への書き出しに失敗しました。');
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = `${state.fileName}.pdf`;
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 1000);
}

async function exportServerDocument(endpoint, fallbackMessage) {
  const response = await fetch(endpoint, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      html: editor.innerHTML,
      title: state.fileName,
      layout: {
        pageSize: state.pageSize,
        orientation: state.orientation,
        margins: state.margins,
        defaultFont: byId('font-family').value || 'Meiryo UI',
        defaultFontSize: parseFloat(byId('font-size').value) || 11,
        headerText: state.headerText,
        footerText: state.footerText,
        showPageNumbers: state.showPageNumbers,
      },
    }),
  });

  if (!response.ok) {
    let message = fallbackMessage;
    try {
      const data = await response.json();
      if (data.error) message = data.error;
    } catch {}
    throw new Error(message);
  }

  return response.blob();
}

function showProperties() {
  const text = editor.innerText || '';
  const chars = text.replace(/\n/g, '').length;
  const words = text.trim() === '' ? 0 : text.split(/\s+/).length;
  alert(`ファイル名: ${state.fileName}\n文字数: ${chars}\n単語数: ${words}\n行数: ${text.split(/\n/).length}`);
}

function escapeHtml(str) {
  return str.replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

function plainFromHtml(html) {
  const div = document.createElement('div');
  div.innerHTML = html;
  return div.innerText;
}

/* ---------- Insert objects ---------- */
function insertPageBreak() {
  exec('insertHTML', '<hr class="page-break"><p><br></p>');
}

function insertBlankPage() {
  exec('insertHTML', '<hr class="page-break"><p><br></p><hr class="page-break"><p><br></p>');
}

function insertImage() {
  const url = prompt('画像 URL を入力してください', 'https://');
  if (url && url !== 'https://') insertImageHtml(url, '');
}

function insertImageFromFile(file) {
  if (!file.type.startsWith('image/')) {
    alert('画像ファイルを選択してください。');
    return;
  }
  const reader = new FileReader();
  reader.onload = () => insertImageHtml(reader.result, file.name.replace(/\.[^/.]+$/, ''));
  reader.readAsDataURL(file);
}

function insertImageHtml(src, alt) {
  const html = `<span class="image-frame" contenteditable="false"><img src="${escapeHtml(src)}" alt="${escapeHtml(alt)}" draggable="false"><span class="image-rotate" contenteditable="false" title="90°回転"></span><span class="image-delete" contenteditable="false" title="画像を削除"></span><span class="image-handle" contenteditable="false" title="サイズ変更"></span></span>&nbsp;`;
  exec('insertHTML', html);
  setTimeout(() => {
    const frames = editor.querySelectorAll('.image-frame');
    const last = frames[frames.length - 1];
    if (last) {
      clearImageSelection();
      last.classList.add('selected');
      const img = last.querySelector('img');
      if (img) {
        const fitToPage = () => {
          if (!img.naturalWidth || !img.naturalHeight) return;
          const maxWidth = Math.max(32, editor.clientWidth);
          const width = Math.min(img.naturalWidth, maxWidth);
          img.style.width = `${Math.round(width)}px`;
          img.style.height = 'auto';
          updateStatus();
        };
        if (img.complete) fitToPage();
        else img.addEventListener('load', fitToPage, { once: true });
      }
    }
    editor.focus();
  }, 0);
}

function clearImageSelection() {
  $$('.image-frame.selected').forEach((frame) => frame.classList.remove('selected'));
}

function initImageEditing() {
  let resize = null;

  editor.addEventListener('mousedown', (e) => {
    const handle = e.target.closest('.image-handle');
    if (handle) {
      const frame = handle.closest('.image-frame');
      const img = frame?.querySelector('img');
      if (!frame || !img) return;
      e.preventDefault();
      clearImageSelection();
      frame.classList.add('selected');
      const rect = img.getBoundingClientRect();
      resize = { img, startX: e.clientX, startWidth: rect.width };
      document.body.classList.add('resizing-image');
      return;
    }

    if (!e.target.closest('.image-frame')) clearImageSelection();
  });

  document.addEventListener('mousemove', (e) => {
    if (!resize) return;
    const width = Math.max(32, resize.startWidth + (e.clientX - resize.startX));
    resize.img.style.width = `${Math.round(width)}px`;
    resize.img.style.height = 'auto';
  });

  document.addEventListener('mouseup', () => {
    if (!resize) return;
    resize = null;
    document.body.classList.remove('resizing-image');
    markImageChanged();
  });

  editor.addEventListener('click', (e) => {
    const frame = e.target.closest('.image-frame');
    if (!frame) return;
    clearImageSelection();
    frame.classList.add('selected');

    if (e.target.closest('.image-delete')) {
      frame.remove();
      markImageChanged();
      return;
    }

    if (e.target.closest('.image-rotate')) {
      const img = frame.querySelector('img');
      if (img) {
        const rotation = (Number(img.dataset.rotation || 0) + 90) % 360;
        img.dataset.rotation = String(rotation);
        img.style.transform = rotation ? `rotate(${rotation}deg)` : '';
        markImageChanged();
      }
    }
  });

  editor.addEventListener('keydown', (e) => {
    const selected = editor.querySelector('.image-frame.selected');
    if (selected && (e.key === 'Delete' || e.key === 'Backspace')) {
      e.preventDefault();
      selected.remove();
      markImageChanged();
    }
  });

  function markImageChanged() {
    state.modified = true;
    updateTitle();
    pushState();
    updateStatus();
  }
}

function showLinkDialog() {
  const sel = window.getSelection();
  byId('link-url').value = '';
  showDialog('link-dialog');
  byId('link-ok').onclick = () => {
    const url = byId('link-url').value.trim();
    if (url) {
      if (sel.rangeCount === 0 || sel.toString().trim() === '') {
        exec('insertHTML', `<a href="${escapeHtml(url)}">${escapeHtml(url)}</a>`);
      } else {
        exec('createLink', url);
      }
    }
    hideDialog('link-dialog');
  };
  byId('link-cancel').onclick = () => hideDialog('link-dialog');
}

function initTablePicker() {
  const grid = $('.table-picker-grid');
  const label = $('.table-picker-label');
  grid.innerHTML = '';
  for (let r = 0; r < 8; r++) {
    for (let c = 0; c < 10; c++) {
      const cell = document.createElement('div');
      cell.dataset.r = r + 1;
      cell.dataset.c = c + 1;
      grid.appendChild(cell);
    }
  }

  let rows = 1, cols = 1;
  grid.addEventListener('mouseover', (e) => {
    const cell = e.target.closest('[data-r]');
    if (!cell) return;
    rows = parseInt(cell.dataset.r, 10);
    cols = parseInt(cell.dataset.c, 10);
    label.textContent = `${cols} × ${rows}`;
    $$('.table-picker-grid div').forEach((div) => {
      const dr = parseInt(div.dataset.r, 10);
      const dc = parseInt(div.dataset.c, 10);
      div.classList.toggle('hover', dr <= rows && dc <= cols);
    });
  });

  grid.addEventListener('click', () => {
    insertTable(rows, cols);
    byId('table-picker').classList.remove('open');
  });

  grid.addEventListener('mouseleave', () => {
    rows = 1; cols = 1;
    label.textContent = '1 × 1';
    $$('.table-picker-grid div').forEach((div) => div.classList.remove('hover'));
  });

  const picker = byId('table-picker');
  byId('btn-insert-table').addEventListener('click', (e) => {
    e.stopPropagation();
    picker.classList.toggle('open');
  });
  document.addEventListener('click', (e) => {
    if (!picker.contains(e.target) && e.target !== byId('btn-insert-table')) {
      picker.classList.remove('open');
    }
  });
}

function insertTable(rows, cols) {
  let html = '<table><tbody>';
  for (let r = 0; r < rows; r++) {
    html += '<tr>';
    for (let c = 0; c < cols; c++) html += '<td><br></td>';
    html += '</tr>';
  }
  html += '</tbody></table><p><br></p>';
  exec('insertHTML', html);
}

/* ---------- Find / Replace ---------- */
function initFindDialog() {
  byId('find-close').addEventListener('click', () => hideDialog('find-dialog'));
  byId('find-next').addEventListener('click', findNext);
  byId('replace-one').addEventListener('click', replaceOne);
  byId('replace-all').addEventListener('click', replaceAll);
  byId('find-text').addEventListener('keydown', (e) => { if (e.key === 'Enter') findNext(); });
}

function allEditorTextNodes() {
  const walker = document.createTreeWalker(editor, NodeFilter.SHOW_TEXT);
  const nodes = [];
  let node;
  while ((node = walker.nextNode())) nodes.push(node);
  return nodes;
}

function selectionEndPosition() {
  const sel = window.getSelection();
  if (!sel || !sel.rangeCount) return null;
  const range = sel.getRangeAt(0);
  if (!editor.contains(range.endContainer)) return null;
  if (range.endContainer.nodeType !== Node.TEXT_NODE) return null;
  return { node: range.endContainer, offset: range.endOffset };
}

function findInTextNode(node, query, caseSensitive, fromOffset) {
  const haystack = caseSensitive ? node.nodeValue : node.nodeValue.toLowerCase();
  const needle = caseSensitive ? query : query.toLowerCase();
  const index = haystack.indexOf(needle, fromOffset);
  if (index === -1) return false;

  const range = document.createRange();
  range.setStart(node, index);
  range.setEnd(node, index + query.length);
  const sel = window.getSelection();
  sel.removeAllRanges();
  sel.addRange(range);
  node.parentElement?.scrollIntoView({ block: 'center', behavior: 'smooth' });
  return true;
}

function findNext() {
  const query = byId('find-text').value;
  if (!query) return false;
  const caseSensitive = byId('find-case').checked;
  const nodes = allEditorTextNodes();
  if (!nodes.length) {
    alert('見つかりませんでした');
    return false;
  }

  const position = selectionEndPosition();
  let startIndex = 0;
  let startOffset = 0;
  if (position) {
    const index = nodes.indexOf(position.node);
    if (index !== -1) {
      startIndex = index;
      startOffset = position.offset;
    }
  }

  // カーソル位置から末尾まで検索し、見つからなければ先頭から折り返す
  for (let i = startIndex; i < nodes.length; i++) {
    if (findInTextNode(nodes[i], query, caseSensitive, i === startIndex ? startOffset : 0)) return true;
  }
  for (let i = 0; i <= startIndex; i++) {
    if (findInTextNode(nodes[i], query, caseSensitive, 0)) return true;
  }

  alert('見つかりませんでした');
  return false;
}

function replaceOne() {
  const query = byId('find-text').value;
  const replacement = byId('replace-text').value;
  if (!query) return;
  const sel = window.getSelection();
  const selected = sel ? sel.toString() : '';
  const caseSensitive = byId('find-case').checked;
  const matches = caseSensitive ? selected === query : selected.toLowerCase() === query.toLowerCase();

  if (matches || findNext()) {
    exec('insertText', replacement);
  } else {
    alert('見つかりませんでした');
  }
}

function escapeRegExp(text) {
  return text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function replaceAll() {
  const query = byId('find-text').value;
  const replacement = byId('replace-text').value;
  if (!query) return;
  const caseSensitive = byId('find-case').checked;
  let count = 0;

  for (const node of allEditorTextNodes()) {
    const text = node.nodeValue;
    let next = text;
    if (caseSensitive) {
      const parts = text.split(query);
      count += parts.length - 1;
      next = parts.join(replacement);
    } else {
      const regex = new RegExp(escapeRegExp(query), 'gi');
      count += (text.match(regex) || []).length;
      next = text.replace(regex, replacement);
    }
    if (next !== text) node.nodeValue = next;
  }

  if (count > 0) {
    pushState();
    updateStatus();
    updateTitle();
    state.modified = true;
  }
  alert(`${count} 件置換しました`);
}

/* ---------- Command Palette ---------- */
function initCommandPalette() {
  const overlay = byId('command-palette');
  const input = byId('palette-input');
  const list = byId('palette-list');

  function render(items, selectedIndex = 0) {
    list.innerHTML = '';
    items.forEach((cmd, i) => {
      const li = document.createElement('li');
      li.classList.toggle('selected', i === selectedIndex);
      li.innerHTML = `<span>${escapeHtml(cmd.label)}</span><kbd>${escapeHtml(cmd.shortcut)}</kbd>`;
      li.addEventListener('click', () => run(cmd));
      list.appendChild(li);
    });
  }

  function run(cmd) {
    hideDialog('command-palette');
    cmd.run();
  }

  function open() {
    overlay.hidden = false;
    input.value = '';
    render(commands);
    input.focus();
  }

  input.addEventListener('input', () => {
    const q = input.value.toLowerCase();
    const filtered = commands.filter((c) => c.label.toLowerCase().includes(q) || c.id.toLowerCase().includes(q));
    render(filtered);
  });

  input.addEventListener('keydown', (e) => {
    const items = Array.from(list.children);
    let idx = items.findIndex((li) => li.classList.contains('selected'));
    if (e.key === 'ArrowDown') {
      e.preventDefault();
      idx = (idx + 1) % items.length;
      render(commands.filter((c) => c.label.toLowerCase().includes(input.value.toLowerCase())), idx);
    } else if (e.key === 'ArrowUp') {
      e.preventDefault();
      idx = (idx - 1 + items.length) % items.length;
      render(commands.filter((c) => c.label.toLowerCase().includes(input.value.toLowerCase())), idx);
    } else if (e.key === 'Enter') {
      e.preventDefault();
      const filtered = commands.filter((c) => c.label.toLowerCase().includes(input.value.toLowerCase()));
      if (filtered[idx]) run(filtered[idx]);
    } else if (e.key === 'Escape') {
      hideDialog('command-palette');
    }
  });

  overlay.addEventListener('click', (e) => { if (e.target === overlay) hideDialog('command-palette'); });

  window.showCommandPalette = open;
}

function showDialog(id) {
  const el = byId(id);
  el.hidden = false;
  const input = el.querySelector('input[type="text"]');
  if (input) setTimeout(() => input.focus(), 0);
}

function hideDialog(id) {
  byId(id).hidden = true;
  editor.focus();
}

/* ---------- Settings ---------- */
function loadSettings() {
  try {
    const saved = JSON.parse(localStorage.getItem(STORAGE_KEY) || '{}');
    if (saved.zoom) state.zoom = saved.zoom;
    if (saved.theme) state.theme = saved.theme;
    if (saved.showRuler !== undefined) state.showRuler = saved.showRuler;
    if (saved.showMarks !== undefined) state.showMarks = saved.showMarks;
    if (saved.defaultFormat) state.defaultFormat = saved.defaultFormat;
    if (saved.pageSize && PAGE_SIZES[saved.pageSize]) state.pageSize = saved.pageSize;
    if (saved.orientation) state.orientation = saved.orientation;
    if (saved.margins) state.margins = { ...state.margins, ...saved.margins };
    if (typeof saved.headerText === 'string') state.headerText = saved.headerText;
    if (typeof saved.footerText === 'string') state.footerText = saved.footerText;
    if (saved.showPageNumbers !== undefined) state.showPageNumbers = saved.showPageNumbers;
    if (saved.readMode !== undefined) state.readMode = saved.readMode;
    if (saved.ribbonCollapsed !== undefined) state.ribbonCollapsed = saved.ribbonCollapsed;
    if (saved.defaultFont) byId('font-family').value = saved.defaultFont;
    if (saved.defaultFontSize) byId('font-size').value = saved.defaultFontSize;
  } catch {}
}

function saveSettings() {
  const data = {
    zoom: state.zoom,
    theme: state.theme,
    showRuler: state.showRuler,
    showMarks: state.showMarks,
    defaultFormat: state.defaultFormat,
    pageSize: state.pageSize,
    orientation: state.orientation,
    margins: state.margins,
    headerText: state.headerText,
    footerText: state.footerText,
    showPageNumbers: state.showPageNumbers,
    readMode: state.readMode,
    ribbonCollapsed: state.ribbonCollapsed,
    defaultFont: byId('font-family').value,
    defaultFontSize: byId('font-size').value,
  };
  try { localStorage.setItem(STORAGE_KEY, JSON.stringify(data)); } catch {}
}

/* ---------- Global shortcuts ---------- */
function initShortcuts() {
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') {
      if (state.readMode) {
        e.preventDefault();
        toggleReadMode(false);
        return;
      }
      const openOverlay = document.querySelector('.overlay:not([hidden])');
      if (openOverlay) {
        e.preventDefault();
        hideDialog(openOverlay.id);
        return;
      }
    }

    const k = e.key.toLowerCase();
    if (e.ctrlKey || e.metaKey) {
      if (k === 'k') {
        e.preventDefault();
        window.showCommandPalette();
      } else if (k === 's') {
        e.preventDefault();
        if (e.shiftKey) saveAs(); else saveDocument();
      } else if (k === 'o') {
        e.preventDefault();
        openDocument();
      } else if (k === 'n') {
        e.preventDefault();
        newDocument();
      } else if (k === 'p') {
        e.preventDefault();
        window.print();
      } else if (k === 'f') {
        e.preventDefault();
        showDialog('find-dialog');
      } else if (k === 'h') {
        e.preventDefault();
        showDialog('find-dialog');
      }
    }
    if (e.key === 'F12') {
      e.preventDefault();
      window.print();
    }
  });

  window.addEventListener('beforeunload', (e) => {
    if (state.modified) {
      e.preventDefault();
      e.returnValue = '';
    }
  });
}

/* ---------- Init ---------- */
function init() {
  loadSettings();
  setTheme(state.theme);
  initTabs();
  initRibbon();
  initPageSettings();
  initEditor();
  initImageEditing();
  initTablePicker();
  initFindDialog();
  initCommandPalette();
  initShortcuts();
  applyPageSettings();
  setZoom(state.zoom);
  setRuler(state.showRuler);
  byId('show-marks').checked = state.showMarks;
  editor.classList.toggle('show-marks', state.showMarks);
  document.body.classList.toggle('read-mode', state.readMode);
  byId('ribbon').classList.toggle('collapsed', state.ribbonCollapsed);
  const collapseUse = byId('btn-ribbon-collapse').querySelector('use');
  if (collapseUse) collapseUse.setAttribute('href', state.ribbonCollapsed ? '#i-chevron-down' : '#i-chevron-up');
  updateTitle();
  updateStatus();
  updateRibbonState();
  updateHistoryButtons();

  matchMedia('(prefers-color-scheme: dark)').addEventListener('change', applySystemTheme);
  resetUndoStack();
}

init();
