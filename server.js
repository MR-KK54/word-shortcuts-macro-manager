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
    return data;
  } catch (err) {
    console.error('Error reading DB, restoring default structure:', err);
    return getDefaultDB();
  }
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
  res.json({ success: true, shortcuts: db.shortcuts });
});

app.post('/api/shortcuts', (req, res) => {
  const { id, name, csv } = req.body;
  if (!name || !name.trim() || !csv || !csv.trim()) {
    return res.status(400).json({ success: false, error: 'Set name and CSV data are required.' });
  }

  const db = readDB();
  const cleanName = name.trim();
  let existingIndex = id ? db.shortcuts.findIndex(s => s.id === id) : db.shortcuts.findIndex(s => s.name.toLowerCase() === cleanName.toLowerCase());

  const shortcutObj = {
    id: existingIndex >= 0 ? db.shortcuts[existingIndex].id : 'sc-' + Date.now(),
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
  res.json({ success: true, ribbon: db.ribbon });
});

app.post('/api/ribbon', (req, res) => {
  const { id, name, filename, base64 } = req.body;
  if (!name || !name.trim() || !base64) {
    return res.status(400).json({ success: false, error: 'Name and file content are required.' });
  }

  const db = readDB();
  let existingIndex = id ? db.ribbon.findIndex(r => r.id === id) : db.ribbon.findIndex(r => r.name.toLowerCase() === name.trim().toLowerCase());

  const ribbonObj = {
    id: existingIndex >= 0 ? db.ribbon[existingIndex].id : 'rb-' + Date.now(),
    name: name.trim(),
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

// --- SYSTEM IMPORT / EXPORT API ---

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
        const idx = db.shortcuts.findIndex(s => s.name.toLowerCase() === (newS.name || '').toLowerCase());
        if (idx >= 0) db.shortcuts[idx] = newS;
        else db.shortcuts.push(newS);
      });
    }

    if (Array.isArray(data.ribbon)) {
      data.ribbon.forEach(newR => {
        const idx = db.ribbon.findIndex(r => r.name.toLowerCase() === (newR.name || '').toLowerCase());
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
