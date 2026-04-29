/* =============================================================
   MY EMAIL TEMPLATES — Outlook Web Add-in
   taskpane.js — All template management and email insertion logic
   =============================================================
   Storage:  OfficeRuntime.storage (persistent, per-user)
   Insertion: body.prependAsync → inserts ABOVE the signature
   ============================================================= */

'use strict';

// ── Constants ──────────────────────────────────────────────────────────────────
const STORAGE_KEY  = 'MyEmailTemplates_v1';     // main storage key
const MAX_STORAGE  = 500;                        // max templates (practical limit)

// ── State ──────────────────────────────────────────────────────────────────────
let templates  = {};       // { [id]: { id, name, content, created, modified } }
let selectedId = null;     // currently selected template ID
let isDirty    = false;    // unsaved changes in editor?

// ── Office.onReady ─────────────────────────────────────────────────────────────
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    bindEvents();
    loadTemplates();
  }
});

// ── Event Binding ──────────────────────────────────────────────────────────────
function bindEvents() {
  id('btnInsert').addEventListener('click', onInsert);
  id('btnSave').addEventListener('click', onSave);
  id('btnNew').addEventListener('click', onNew);
  id('btnDelete').addEventListener('click', onDelete);
  id('searchInput').addEventListener('input', onSearch);
  id('btnClearSearch').addEventListener('click', clearSearch);
  id('txtContent').addEventListener('input', onEditorChange);
  id('txtName').addEventListener('keydown', (e) => { if (e.key === 'Enter') onSave(); });
  id('btnOpenFolder').addEventListener('click', () => showModal());
  id('btnCloseModal').addEventListener('click', () => hideModal());
  id('btnExportTemplates').addEventListener('click', exportTemplates);
  id('infoModal').addEventListener('click', (e) => { if (e.target === id('infoModal')) hideModal(); });
}

// ── Storage: Load ──────────────────────────────────────────────────────────────
async function loadTemplates() {
  showLoading(true);
  try {
    const raw = await OfficeRuntime.storage.getItem(STORAGE_KEY);
    templates = raw ? JSON.parse(raw) : {};
  } catch (_) {
    templates = {};
  }
  showLoading(false);
  renderAll();
}

// ── Storage: Persist ──────────────────────────────────────────────────────────
async function persist() {
  try {
    await OfficeRuntime.storage.setItem(STORAGE_KEY, JSON.stringify(templates));
  } catch (err) {
    showStatus('Storage error: ' + err.message, 'error');
  }
}

// ── CRUD ──────────────────────────────────────────────────────────────────────
function mkId() {
  return 't' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
}

async function upsertTemplate(name, content, id) {
  const tid  = id || mkId();
  const prev = templates[tid];
  templates[tid] = {
    id:       tid,
    name:     name.trim(),
    content:  content,
    created:  prev ? prev.created : new Date().toISOString(),
    modified: new Date().toISOString()
  };
  await persist();
  return tid;
}

async function removeTemplate(tid) {
  delete templates[tid];
  await persist();
}

function getTemplatesSorted() {
  return Object.values(templates).sort((a, b) => a.name.localeCompare(b.name));
}

// ── Render ─────────────────────────────────────────────────────────────────────
function renderAll() {
  const query = id('searchInput').value.trim().toLowerCase();
  renderList(query);
  updateCountBadge();
}

function renderList(query) {
  const list   = id('templateList');
  const empty  = id('emptyState');
  const noRes  = id('noResults');
  const all    = getTemplatesSorted();
  const shown  = query ? all.filter(t => t.name.toLowerCase().includes(query) || t.content.toLowerCase().includes(query)) : all;

  list.innerHTML = '';

  if (all.length === 0) {
    empty.style.display  = 'flex';
    noRes.style.display  = 'none';
  } else if (shown.length === 0) {
    empty.style.display  = 'none';
    noRes.style.display  = 'flex';
    id('noResultsText').textContent = 'No match for "' + query + '"';
  } else {
    empty.style.display = 'none';
    noRes.style.display = 'none';

    shown.forEach(t => {
      const item = document.createElement('div');
      item.className = 'template-item' + (t.id === selectedId ? ' selected' : '');
      item.dataset.id = t.id;

      const preview = t.content.replace(/\s+/g, ' ').trim().slice(0, 70);
      item.innerHTML =
        '<div class="ti-name">' + esc(t.name) + '</div>' +
        '<div class="ti-preview">' + esc(preview) + (t.content.length > 70 ? '…' : '') + '</div>';

      item.addEventListener('click', () => selectTemplate(t.id));
      list.appendChild(item);
    });
  }
}

function selectTemplate(tid) {
  if (isDirty && selectedId !== tid) {
    // Warn only if there's a different template selected with unsaved edits
    const cur = id('txtName').value.trim();
    const same = Object.values(templates).find(t => t.name === cur && t.id === selectedId);
    if (same && id('txtContent').value !== same.content) {
      if (!confirm('You have unsaved changes. Switch template anyway?')) return;
    }
  }

  selectedId = tid;
  const t = templates[tid];
  if (!t) return;

  id('txtContent').value = t.content;
  id('txtName').value    = t.name;
  id('btnInsert').disabled = false;
  id('btnDelete').disabled = false;
  isDirty = false;

  renderAll();

  // Scroll selected item into view
  const sel = id('templateList').querySelector('.selected');
  if (sel) sel.scrollIntoView({ block: 'nearest' });
}

function updateCountBadge() {
  const count = Object.keys(templates).length;
  id('templateCount').textContent = count;
  id('infoCount').textContent = count;
}

// ── Button Handlers ───────────────────────────────────────────────────────────

// INSERT
async function onInsert() {
  const content = id('txtContent').value;
  if (!content.trim()) {
    showStatus('No content to insert. Select or write a template first.', 'warn');
    return;
  }

  const btn = id('btnInsert');
  btn.disabled = true;
  btn.innerHTML = '<span class="spin-sm"></span> Inserting…';

  try {
    await insertIntoEmail(content);
    showStatus('Template inserted! Your signature is preserved above.', 'success');
  } catch (err) {
    showStatus('Could not insert: ' + err.message, 'error');
  } finally {
    btn.disabled = false;
    btn.innerHTML =
      '<svg viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg" width="14" height="14">' +
      '<path d="M8 2v9M4 8l4 4 4-4" stroke="white" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/>' +
      '<line x1="2" y1="14" x2="14" y2="14" stroke="white" stroke-width="1.5" stroke-linecap="round"/></svg>' +
      ' Insert Template into Email';
    id('btnInsert').disabled = !id('txtContent').value.trim();
  }
}

// SAVE
async function onSave() {
  const name    = id('txtName').value.trim();
  const content = id('txtContent').value;

  if (!name)         { showStatus('Please enter a template name.', 'warn'); id('txtName').focus(); return; }
  if (!content.trim()) { showStatus('Template content cannot be empty.', 'warn'); id('txtContent').focus(); return; }

  // Duplicate name guard (different template)
  const dup = getTemplatesSorted().find(t => t.name === name && t.id !== selectedId);
  if (dup) {
    showStatus('"' + name + '" already exists. Choose a different name.', 'warn');
    id('txtName').focus();
    return;
  }

  const savedId = await upsertTemplate(name, content, selectedId || undefined);
  selectedId = savedId;
  isDirty = false;
  renderAll();
  showStatus('✓ "' + name + '" saved successfully.', 'success');
  id('btnDelete').disabled = false;
}

// NEW
function onNew() {
  if (isDirty) {
    if (!confirm('Discard unsaved changes and start a new template?')) return;
  }
  selectedId = null;
  isDirty    = false;
  id('txtContent').value = '';
  id('txtName').value    = '';
  id('btnInsert').disabled = true;
  id('btnDelete').disabled = true;
  renderAll();
  id('txtName').focus();
}

// DELETE
async function onDelete() {
  if (!selectedId) return;
  const t = templates[selectedId];
  if (!t) return;
  if (!confirm('Delete "' + t.name + '"?\n\nThis cannot be undone.')) return;

  await removeTemplate(selectedId);
  selectedId = null;
  isDirty    = false;
  id('txtContent').value = '';
  id('txtName').value    = '';
  id('btnInsert').disabled = true;
  id('btnDelete').disabled = true;
  renderAll();
  showStatus('Template deleted.', 'info');
}

// SEARCH
function onSearch(e) {
  const q = e.target.value;
  id('btnClearSearch').style.display = q ? 'flex' : 'none';
  renderList(q.trim().toLowerCase());
}

function clearSearch() {
  id('searchInput').value = '';
  id('btnClearSearch').style.display = 'none';
  renderAll();
}

// EDITOR CHANGE — track dirty state
function onEditorChange() {
  isDirty = true;
  const hasContent = id('txtContent').value.trim().length > 0;
  // Enable insert even for unsaved content
  id('btnInsert').disabled = !hasContent;
}

// ── Email Insertion ───────────────────────────────────────────────────────────
/**
 * Uses body.prependAsync to insert the template at the TOP of the email body.
 * This places the template ABOVE any existing signature, which Outlook
 * automatically placed in the body when the compose window was opened.
 */
function insertIntoEmail(plainText) {
  return new Promise((resolve, reject) => {
    const item = Office.context.mailbox.item;
    if (!item) {
      reject(new Error('No active email found. Open a compose or reply window first.'));
      return;
    }

    const html = textToHtml(plainText);

    item.body.prependAsync(
      html,
      { coercionType: Office.CoercionType.Html },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve();
        } else {
          const msg = result.error ? result.error.message : 'Insert failed';
          reject(new Error(msg));
        }
      }
    );
  });
}

/**
 * Converts plain-text template content to Outlook-safe HTML.
 * Preserves blank lines as paragraph spacing.
 * Escapes HTML special characters.
 */
function textToHtml(text) {
  const lines = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');
  let html = '<div style="font-family:Calibri,Arial,Helvetica,sans-serif;' +
             'font-size:11pt;color:#000000;line-height:1.4;">';

  lines.forEach(line => {
    const safe = esc(line);
    if (safe.trim() === '') {
      html += '<p style="margin:0;min-height:1em;">&nbsp;</p>';
    } else {
      html += '<p style="margin:0;">' + safe + '</p>';
    }
  });

  html += '</div><p style="margin:0;">&nbsp;</p>';
  return html;
}

// ── Export ────────────────────────────────────────────────────────────────────
function exportTemplates() {
  const data = JSON.stringify(Object.values(templates), null, 2);
  const blob = new Blob([data], { type: 'application/json' });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  a.href     = url;
  a.download = 'MyEmailTemplates_export_' + new Date().toISOString().slice(0,10) + '.json';
  a.click();
  URL.revokeObjectURL(url);
}

// ── UI Helpers ────────────────────────────────────────────────────────────────
function showLoading(on) {
  id('loadingPane').style.display = on ? 'flex'  : 'none';
  id('mainPane').style.display    = on ? 'none'  : 'flex';
}

let statusTimer = null;
function showStatus(msg, type) {
  const bar = id('statusBar');
  const el  = id('statusMsg');
  el.textContent = msg;
  el.className   = 'status-' + type;
  bar.style.display = 'flex';
  clearTimeout(statusTimer);
  statusTimer = setTimeout(() => { bar.style.display = 'none'; }, 4500);
}

function showModal() {
  updateCountBadge();
  id('infoModal').style.display = 'flex';
}

function hideModal() {
  id('infoModal').style.display = 'none';
}

function id(i) { return document.getElementById(i); }

function esc(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
