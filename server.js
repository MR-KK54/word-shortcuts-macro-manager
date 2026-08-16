const express = require('express');
const cors = require('cors');
const fs = require('fs');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3000;
const DATA_DIR = process.env.DATA_DIR || path.join(__dirname, 'data');
const DB_FILE = path.join(DATA_DIR, 'db.json');

// Middleware
app.use(cors());
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ extended: true, limit: '50mb' }));
app.use(express.static(path.join(__dirname, 'public')));

// Ensure data directory and db file exist
if (!fs.existsSync(DATA_DIR)) {
  fs.mkdirSync(DATA_DIR, { recursive: true });
}

function getDefaultDB() {
  return {
    macros: [
      {
        id: 'macro-1',
        group: 'Formatting Tools',
        name: 'CleanFormatting',
        type: 'bas',
        code: `Attribute VB_Name = "CleanFormatting"\nOption Explicit\n\nSub RemoveExtraSpaces()\n    With Selection.Find\n        .ClearFormatting\n        .Replacement.ClearFormatting\n        .Text = "  "\n        .Replacement.Text = " "\n        .Forward = True\n        .Wrap = wdFindContinue\n        .Execute Replace:=wdReplaceAll\n    End With\n    MsgBox "Extra spaces removed!", vbInformation, "Clean Formatting"\nEnd Sub`,
        updatedAt: new Date().toISOString()
      },
      {
        id: 'macro-2',
        group: 'Document Utilities',
        name: 'WordCountReport',
        type: 'bas',
        code: `Attribute VB_Name = "WordCountReport"\nOption Explicit\n\nSub ShowDocumentStats()\n    Dim words As Long, paras As Long\n    words = ActiveDocument.Words.Count\n    paras = ActiveDocument.Paragraphs.Count\n    MsgBox "Words: " & words & vbCrLf & "Paragraphs: " & paras, vbInformation, "Document Statistics"\nEnd Sub`,
        updatedAt: new Date().toISOString()
      }
    ],
    shortcuts: [
      {
        id: 'shortcut-1',
        name: 'Editor Master Layout',
        csv: `KeyCategory,Command,KeyCode,KeyCode2,KeyString\n1,CleanFormatting,70,1024,"Ctrl+Shift+F"\n1,WordCountReport,87,1024,"Ctrl+Shift+W"`,
        updatedAt: new Date().toISOString()
      },
      {
        id: 'shortcut-kishore',
        group: 'Kishore',
        name: "Kishore Shortcut's",
        csv: `KeyCategory,Command,KeyCode,KeyCode2,KeyString
1,TableMergeCells,593,0,"Ctrl+Q"
1,TextToTable,596,0,"Ctrl+T"
1,InsertEmSpace,619,0,"Ctrl+Num +"
1,BorderLineStyle,834,0,"Ctrl+Shift+B"
1,CopyFormat,835,0,"Ctrl+Shift+C"
1,FormatParagraph,836,0,"Ctrl+Shift+D"
1,FontColorPicker,838,0,"Ctrl+Shift+F"
1,TableProperties,852,0,"Ctrl+Shift+T"
1,PasteFormat,854,0,"Ctrl+Shift+V"
1,InsertEnSpace,875,0,"Ctrl+Shift+Num +"
1,TableSplitCells,1107,0,"Alt+S"
1,TableSplit,1112,0,"Alt+X"
1,EditPasteSpecial,1622,0,"Alt+Ctrl+V"
1,DecreaseIndent,1829,0,"Alt+Ctrl+Shift+Left"
1,IncreaseIndent,1831,0,"Alt+Ctrl+Shift+Right"
2,Normal.MergeFileTool.UltimateExactMerge,589,0,"Ctrl+M"
2,Normal.setcolumnlevels.reset_space,608,0,"Ctrl+Num 0"
2,Normal.ReFLinkBookmark.AllInsertBookmarks_ToSelection,611,0,"Ctrl+Num 3"
2,Normal.modREY_Bookmarks.InsertBookmark_CurrentParagraph,622,0,"Ctrl+Num ."
2,Normal.PasteOptionSelector.PasteOptionSelector,848,0,"Ctrl+Shift+P"
2,Normal.Replace_images.ReplaceSelectedPicture_SessionFolder,856,0,"Ctrl+Shift+X"
2,Normal.Module3.DecrementLeading,1062,0,"Alt+Up"
2,Normal.Module3.IncrementLeading,1064,0,"Alt+Down"
2,Normal.Module3.InsertAndFormatPicture,1089,0,"Alt+A"
2,Normal.Module3.ShowCleanupTool,1091,0,"Alt+C"
2,Normal.modHideUnhide.ShowHideUnhideGUI,1092,0,"Alt+D"
2,Normal.Module1.ShowPlaceholderPicker,1345,0,"Alt+Shift+A"
2,Normal.Module1.ShowPlaceholderPicker1,1363,0,"Alt+Shift+S"
2,Normal.setcolumnlevels.after_space_DOWN,1573,0,"Alt+Ctrl+Left"
2,Normal.setcolumnlevels.before_space_up,1574,0,"Alt+Ctrl+Up"
2,Normal.setcolumnlevels.after_space_UP,1575,0,"Alt+Ctrl+Right"
2,Normal.setcolumnlevels.before_space_DOWN,1576,0,"Alt+Ctrl+Down"
2,Normal.GroupSelectedTextboxes.GroupTextboxesInSelectedParagraph,1607,0,"Alt+Ctrl+G"
2,Normal.KeyboardshortcutsExToImp.KeyboardShortcutsMenu,1624,0,"Alt+Ctrl+X"
2,Normal.TextboxCleaner.CleanSelectedTextboxes,1757,0,"Alt+Ctrl+]"
2,Normal.indent.RightIndentIncrease,1830,0,"Alt+Ctrl+Shift+Up"
2,Normal.indent.RightIndentDecrease,1832,0,"Alt+Ctrl+Shift+Down"`,
        updatedAt: new Date().toISOString()
      }
    ],
    ribbon: [],
    version: '1.0'
  };
}

function readDB() {
  try {
    if (!fs.existsSync(DB_FILE)) {
      const defaultData = getDefaultDB();
      fs.writeFileSync(DB_FILE, JSON.stringify(defaultData, null, 2), 'utf8');
      return defaultData;
    }
    const raw = fs.readFileSync(DB_FILE, 'utf8');
    const data = JSON.parse(raw);
    if (!data.macros) data.macros = [];
    if (!data.shortcuts) data.shortcuts = [];
    if (!data.ribbon) data.ribbon = [];
    seedMissingDefaults(data);
    if (dedupeShortcuts(data)) writeDB(data);
    return data;
  } catch (err) {
    console.error('Error reading DB, restoring default structure:', err);
    const defaultData = getDefaultDB();
    seedMissingDefaults(defaultData);
    return defaultData;
  }
}

// Add bundled default items (e.g. Kishore shortcut set) to existing databases
// so new defaults ship with every deployed server even if data/ already exists.
function seedMissingDefaults(data) {
  const defaults = getDefaultDB();
  let changed = false;

  (defaults.shortcuts || []).forEach(def => {
    const exists = (data.shortcuts || []).some(s =>
      s.name.toLowerCase() === (def.name || '').toLowerCase()
    );
    if (!exists) {
      data.shortcuts.push({ ...def, id: def.id + '-' + Date.now() });
      changed = true;
    }
  });

  if (changed) writeDB(data);
}

// One set per name: merge duplicates (same name in different groups) keeping
// the most recently updated entry, so bundled saves never show twice.
function dedupeShortcuts(data) {
  const seen = new Map();
  let changed = false;
  (data.shortcuts || []).forEach(s => {
    const key = (s.name || '').toLowerCase();
    const existing = seen.get(key);
    if (!existing || (s.updatedAt || '') > (existing.updatedAt || '')) seen.set(key, s);
  });
  const merged = Array.from(seen.values());
  if (merged.length !== (data.shortcuts || []).length) {
    data.shortcuts = merged;
    changed = true;
  }
  return changed;
}

function writeDB(data) {
  try {
    fs.writeFileSync(DB_FILE, JSON.stringify(data, null, 2), 'utf8');
    return true;
  } catch (err) {
    console.error('Error writing DB:', err);
    return false;
  }
}

// --- API Endpoints ---

// Health Check
app.get('/api/health', (req, res) => {
  res.json({ status: 'ok', server: 'Word Toolkit Render Server', timestamp: new Date().toISOString() });
});

// GET /api/macros - List macros with optional group filter
app.get('/api/macros', (req, res) => {
  const db = readDB();
  let macros = db.macros;

  const { group, search } = req.query;
  if (group) {
    macros = macros.filter(m => (m.group || 'General').toLowerCase() === group.toLowerCase());
  }
  if (search) {
    const q = search.toLowerCase();
    macros = macros.filter(m =>
      (m.name || '').toLowerCase().includes(q) ||
      (m.group || '').toLowerCase().includes(q) ||
      (m.code || '').toLowerCase().includes(q)
    );
  }

  res.json({ success: true, count: macros.length, macros });
});

// GET /api/groups - Distinct group names
app.get('/api/groups', (req, res) => {
  const db = readDB();
  const groupSet = new Set(db.macros.map(m => m.group || 'General'));
  res.json({ success: true, groups: Array.from(groupSet) });
});

// POST /api/macros - Store or update macro (by ID or Group+Name replacement)
app.post('/api/macros', (req, res) => {
  const { id, group, name, type, code, description } = req.body;

  if (!name || !name.trim() || !code || !code.trim()) {
    return res.status(400).json({ success: false, error: 'Module name and code are required.' });
  }

  const cleanName = name.trim().replace(/[^a-zA-Z0-9_]/g, '');
  const cleanGroup = (group && group.trim()) ? group.trim() : 'General';
  const cleanType = (type || 'bas').toLowerCase();

  const db = readDB();
  
  // Look for existing macro by ID or by matching name in the same group (Replace if exists!)
  let existingIndex = -1;
  if (id) {
    existingIndex = db.macros.findIndex(m => m.id === id);
  }
  if (existingIndex === -1) {
    existingIndex = db.macros.findIndex(m =>
      m.name.toLowerCase() === cleanName.toLowerCase() &&
      (m.group || 'General').toLowerCase() === cleanGroup.toLowerCase()
    );
  }

  const now = new Date().toISOString();
  const macroObj = {
    id: existingIndex >= 0 ? db.macros[existingIndex].id : 'macro-' + Date.now() + '-' + Math.random().toString(36).substr(2, 4),
    group: cleanGroup,
    name: cleanName,
    type: cleanType,
    code: code,
    description: description || '',
    updatedAt: now
  };

  let action = 'created';
  if (existingIndex >= 0) {
    db.macros[existingIndex] = macroObj;
    action = 'replaced';
  } else {
    db.macros.push(macroObj);
  }

  if (writeDB(db)) {
    res.json({ success: true, action, macro: macroObj });
  } else {
    res.status(500).json({ success: false, error: 'Failed to write macro to database.' });
  }
});

// DELETE /api/macros/:id
app.delete('/api/macros/:id', (req, res) => {
  const { id } = req.params;
  const db = readDB();
  const initialLength = db.macros.length;
  db.macros = db.macros.filter(m => m.id !== id);

  if (db.macros.length === initialLength) {
    return res.status(404).json({ success: false, error: 'Macro not found.' });
  }

  writeDB(db);
  res.json({ success: true, message: 'Macro deleted.' });
});

// --- SHORTCUTS API ---
app.get('/api/shortcuts', (req, res) => {
  const db = readDB();
  let shortcuts = db.shortcuts;
  const { group } = req.query;
  if (group) {
    shortcuts = shortcuts.filter(s => (s.group || 'General').toLowerCase() === group.toLowerCase());
  }
  res.json({ success: true, shortcuts });
});

app.post('/api/shortcuts', (req, res) => {
  const { id, name, csv, group } = req.body;
  if (!name || !name.trim() || !csv || !csv.trim()) {
    return res.status(400).json({ success: false, error: 'Set name and CSV data are required.' });
  }

  const db = readDB();
  const cleanName = name.trim();
  const cleanGroup = (group && group.trim()) ? group.trim() : 'General';
  let existingIndex = id
    ? db.shortcuts.findIndex(s => s.id === id)
    : db.shortcuts.findIndex(s =>
        s.name.toLowerCase() === cleanName.toLowerCase()
      );

  const shortcutObj = {
    id: existingIndex >= 0 ? db.shortcuts[existingIndex].id : 'sc-' + Date.now(),
    group: cleanGroup,
    name: cleanName,
    csv: csv,
    updatedAt: new Date().toISOString()
  };

  if (existingIndex >= 0) {
    db.shortcuts[existingIndex] = shortcutObj;
  } else {
    db.shortcuts.push(shortcutObj);
  }

  writeDB(db);
  res.json({ success: true, shortcut: shortcutObj });
});

app.delete('/api/shortcuts/:id', (req, res) => {
  const { id } = req.params;
  const db = readDB();
  db.shortcuts = db.shortcuts.filter(s => s.id !== id);
  writeDB(db);
  res.json({ success: true, message: 'Shortcut set deleted.' });
});

// --- RIBBON API ---
app.get('/api/ribbon', (req, res) => {
  const db = readDB();
  let ribbon = db.ribbon;
  const { group } = req.query;
  if (group) {
    ribbon = ribbon.filter(r => (r.group || 'General').toLowerCase() === group.toLowerCase());
  }
  res.json({ success: true, ribbon });
});

app.post('/api/ribbon', (req, res) => {
  const { id, name, filename, base64, group } = req.body;
  if (!name || !name.trim() || !base64) {
    return res.status(400).json({ success: false, error: 'Name and file content are required.' });
  }

  const db = readDB();
  const cleanName = name.trim();
  const cleanGroup = (group && group.trim()) ? group.trim() : 'General';
  let existingIndex = id
    ? db.ribbon.findIndex(r => r.id === id)
    : db.ribbon.findIndex(r =>
        r.name.toLowerCase() === cleanName.toLowerCase() &&
        (r.group || 'General').toLowerCase() === cleanGroup.toLowerCase()
      );

  const ribbonObj = {
    id: existingIndex >= 0 ? db.ribbon[existingIndex].id : 'rb-' + Date.now(),
    group: cleanGroup,
    name: cleanName,
    filename: filename || 'Word.officeUI',
    base64: base64,
    updatedAt: new Date().toISOString()
  };

  if (existingIndex >= 0) {
    db.ribbon[existingIndex] = ribbonObj;
  } else {
    db.ribbon.push(ribbonObj);
  }

  writeDB(db);
  res.json({ success: true, ribbon: ribbonObj });
});

app.delete('/api/ribbon/:id', (req, res) => {
  const { id } = req.params;
  const db = readDB();
  db.ribbon = db.ribbon.filter(r => r.id !== id);
  writeDB(db);
  res.json({ success: true, message: 'Ribbon profile deleted.' });
});

// Connector source files (used by WordToolkit_Setup.vbs to install into Word)
app.get('/api/connector/:file', (req, res) => {
  const { file } = req.params;
  if (!/^[A-Za-z0-9_\-]+\.(bas|vbs)$/.test(file)) {
    return res.status(400).json({ success: false, error: 'Invalid file name.' });
  }
  const fullPath = path.join(__dirname, 'connector', file);
  if (!fs.existsSync(fullPath)) {
    return res.status(404).json({ success: false, error: 'File not found.' });
  }
  res.setHeader('Content-Type', 'text/plain; charset=utf-8');
  let content = fs.readFileSync(fullPath, 'utf8');
  if (file === 'WordToolkit_Setup.vbs') {
    const protocol = req.headers['x-forwarded-proto'] || req.protocol || 'http';
    const host = req.get('host') || 'localhost:3000';
    const serverApiUrl = `${protocol}://${host}/api`;
    content = content.replace('"http://localhost:3000/api"', `"${serverApiUrl}"`);
  }
  res.send(content);
});

// --- SYSTEM IMPORT / EXPORT API ---

// --- DIRECT WORD SYNC BUNDLES (plain text, VBA-friendly) ---
// All macros in a parseable bundle so the VBA connector can install them into Word directly
app.get('/api/sync/macros', (req, res) => {
  const db = readDB();
  const { group } = req.query;
  let list = db.macros || [];
  if (group) list = list.filter(m => (m.group || 'General').toLowerCase() === group.toLowerCase());
  let out = 'WORDTOOLKIT MACRO BUNDLE v1\n';
  list.forEach(m => {
    out += `@group=${m.group || 'General'}\n`;
    out += `@name=${m.name}\n`;
    out += `@type=${(m.type || 'bas').toLowerCase()}\n`;
    out += (m.code || '') + '\n';
    out += '#WTMACRO-END#\n';
  });
  res.setHeader('Content-Type', 'text/plain; charset=utf-8');
  res.send(out);
});

// All shortcut sets combined as CSV so the VBA connector can apply them to Word directly
app.get('/api/sync/shortcuts', (req, res) => {
  const db = readDB();
  const { group } = req.query;
  let list = db.shortcuts || [];
  if (group) list = list.filter(s => (s.group || 'General').toLowerCase() === group.toLowerCase());
  let out = '';
  list.forEach(sc => {
    out += `#SET:${sc.name}${sc.group && sc.group !== 'General' ? ` (${sc.group})` : ''}\n`;
    if (!/^KeyCategory/i.test((sc.csv || '').trim())) {
      out += 'KeyCategory,Command,KeyCode,KeyCode2,KeyString\n';
    }
    out += (sc.csv || '') + '\n';
  });
  res.setHeader('Content-Type', 'text/plain; charset=utf-8');
  res.send(out);
});

// Ribbon profiles as plain .officeUI XML so the VBA connector can write the file
// directly into Word's AppData folder (one profile per machine - last one wins).
app.get('/api/sync/ribbon', (req, res) => {
  const db = readDB();
  const { group } = req.query;
  let list = db.ribbon || [];
  if (group) list = list.filter(r => (r.group || 'General').toLowerCase() === group.toLowerCase());
  let out = '';
  list.forEach(r => {
    out += `#RIBBON:${r.name}${r.group && r.group !== 'General' ? ` (${r.group})` : ''}\n`;
    out += (r.base64 ? Buffer.from(r.base64, 'base64').toString('utf8') : '') + '\n';
    out += '#WTRIBBON-END#\n';
  });
  res.setHeader('Content-Type', 'text/plain; charset=utf-8');
  res.send(out);
});

// Export full backup as JSON
app.get('/api/export', (req, res) => {
  const db = readDB();
  res.setHeader('Content-Type', 'application/json');
  res.setHeader('Content-Disposition', `attachment; filename="word-toolkit-backup-${Date.now()}.json"`);
  res.json({
    appName: 'Word Shortcuts & Macro Manager',
    version: '1.0',
    exportedAt: new Date().toISOString(),
    macros: db.macros,
    shortcuts: db.shortcuts,
    ribbon: db.ribbon
  });
});

// Import full backup JSON (mode: 'merge' or 'replace')
app.post('/api/import', (req, res) => {
  const { mode = 'merge', data } = req.body;

  if (!data || (typeof data !== 'object')) {
    return res.status(400).json({ success: false, error: 'Invalid import package data.' });
  }

  const db = readDB();

  if (mode === 'replace') {
    db.macros = Array.isArray(data.macros) ? data.macros : [];
    db.shortcuts = Array.isArray(data.shortcuts) ? data.shortcuts : [];
    db.ribbon = Array.isArray(data.ribbon) ? data.ribbon : [];
  } else {
    // Merge mode: replace existing macros with matching group+name, append new ones
    if (Array.isArray(data.macros)) {
      data.macros.forEach(newM => {
        const idx = db.macros.findIndex(m =>
          m.name.toLowerCase() === (newM.name || '').toLowerCase() &&
          (m.group || 'General').toLowerCase() === (newM.group || 'General').toLowerCase()
        );
        if (idx >= 0) db.macros[idx] = { ...newM, updatedAt: new Date().toISOString() };
        else db.macros.push({ ...newM, id: newM.id || 'macro-' + Date.now() + '-' + Math.random().toString(36).substr(2,4) });
      });
    }

    if (Array.isArray(data.shortcuts)) {
      data.shortcuts.forEach(newS => {
        const idx = db.shortcuts.findIndex(s =>
          s.name.toLowerCase() === (newS.name || '').toLowerCase()
        );
        if (idx >= 0) db.shortcuts[idx] = newS;
        else db.shortcuts.push(newS);
      });
    }

    if (Array.isArray(data.ribbon)) {
      data.ribbon.forEach(newR => {
        const idx = db.ribbon.findIndex(r =>
          r.name.toLowerCase() === (newR.name || '').toLowerCase() &&
          (r.group || 'General').toLowerCase() === (newR.group || 'General').toLowerCase()
        );
        if (idx >= 0) db.ribbon[idx] = newR;
        else db.ribbon.push(newR);
      });
    }
  }

  writeDB(db);
  res.json({
    success: true,
    message: `Data successfully imported (${mode} mode).`,
    counts: {
      macros: db.macros.length,
      shortcuts: db.shortcuts.length,
      ribbon: db.ribbon.length
    }
  });
});

// Fallback route for SPA
app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// Start server
app.listen(PORT, () => {
  console.log(`====================================================`);
  console.log(`🚀 Word Toolkit Server running on http://localhost:${PORT}`);
  console.log(`====================================================`);
});
