// Word Toolkit & Macro Manager - Frontend Client Logic

const $ = (sel, el = document) => el.querySelector(sel);
const $$ = (sel, el = document) => Array.from(el.querySelectorAll(sel));

// State Management
let apiBaseUrl = localStorage.getItem('wt_server_url') || '/api';
if (apiBaseUrl.endsWith('/')) apiBaseUrl = apiBaseUrl.slice(0, -1);

let isServerOnline = false;
let currentMacros = [];
let currentShortcuts = [];
let currentRibbon = [];

/* ---------- INIT & CONNECTOR FILES ---------- */
const CONNECTOR_FILES = {
'Toolkit_Helpers.bas': `Attribute VB_Name = "Toolkit_Helpers"
Option Explicit

Function GetFolder(promptText As String) As String
    Dim fd As Object
    On Error Resume Next
    Set fd = Application.FileDialog(4)
    If Not fd Is Nothing Then
        fd.Title = promptText
        If fd.Show = -1 Then
            GetFolder = fd.SelectedItems(1)
            Exit Function
        End If
    End If
    GetFolder = ""
End Function`,

'Toolkit_Macros.bas': `Attribute VB_Name = "Toolkit_Macros"
Option Explicit

Const vbext_ct_StdModule = 1
Const vbext_ct_ClassModule = 2
Const vbext_ct_MSForm = 3

Sub ExportAllMacros()
    Dim vbProj As Object, vbComp As Object
    Dim exportPath As String, ext As String, count As Integer
    On Error GoTo HandleErr

    exportPath = GetFolder("Select a folder to export macro modules to")
    If exportPath = "" Then Exit Sub

    Set vbProj = Application.VBE.ActiveVBProject
    For Each vbComp In vbProj.VBComponents
        Select Case vbComp.Type
            Case vbext_ct_StdModule: ext = ".bas"
            Case vbext_ct_ClassModule: ext = ".cls"
            Case vbext_ct_MSForm: ext = ".frm"
            Case Else: ext = ""
        End Select

        If ext <> "" And vbComp.Name <> "ThisDocument" Then
            vbComp.Export exportPath & "\\" & vbComp.Name & ext
            count = count + 1
        End If
    Next vbComp

    MsgBox count & " macro module(s) exported to:" & vbCrLf & exportPath, vbInformation, "Export Macros Successful"
    Exit Sub

HandleErr:
    MsgBox "Could not export macros. Make sure 'Trust access to the VBA project object model' is enabled." & vbCrLf & "Error: " & Err.Description, vbCritical, "Export Macros Error"
End Sub

Sub ImportAllMacros()
    Dim vbProj As Object, importPath As String
    Dim fso As Object, fld As Object, f As Object
    Dim compName As String, ext As String, count As Integer, replacedCount As Integer
    Dim existingComp As Object

    On Error GoTo HandleErr
    importPath = GetFolder("Select the folder containing .bas / .cls / .frm files to import")
    If importPath = "" Then Exit Sub

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set vbProj = Application.VBE.ActiveVBProject
    Set fld = fso.GetFolder(importPath)

    For Each f In fld.Files
        ext = LCase(fso.GetExtensionName(f.Name))
        If ext = "bas" Or ext = "cls" Or ext = "frm" Then
            compName = fso.GetBaseName(f.Name)

            ' Replace existing macro module in Word if already present!
            On Error Resume Next
            Set existingComp = Nothing
            Set existingComp = vbProj.VBComponents(compName)
            If Not existingComp Is Nothing Then
                vbProj.VBComponents.Remove existingComp
                replacedCount = replacedCount + 1
            End If
            On Error GoTo HandleErr

            vbProj.VBComponents.Import f.Path
            count = count + 1
        End If
    Next f

    MsgBox count & " macro module(s) imported (" & replacedCount & " existing module(s) replaced)." & vbCrLf & _
           "Remember to save Normal.dotm or active template.", vbInformation, "Import Macros Successful"
    Exit Sub

HandleErr:
    MsgBox "Could not import macros. Make sure 'Trust access to the VBA project object model' is enabled." & vbCrLf & "Error: " & Err.Description, vbCritical, "Import Macros Error"
End Sub`,

'Toolkit_Shortcuts.bas': `Attribute VB_Name = "Toolkit_Shortcuts"
Option Explicit

Sub ExportKeyboardShortcuts()
    Dim kb As Object, filePath As Variant, fnum As Integer, count As Integer
    On Error GoTo HandleErr

    filePath = Application.GetSaveAsFilename(InitialFileName:="WordKeyboardShortcuts.csv", FileFilter:="CSV Files (*.csv), *.csv")
    If filePath = False Or filePath = "False" Then Exit Sub

    fnum = FreeFile
    Open filePath For Output As #fnum
    Print #fnum, "KeyCategory,Command,KeyCode,KeyCode2,KeyString"

    For Each kb In Application.KeyBindings
        Print #fnum, kb.KeyCategory & "," & kb.Command & "," & kb.KeyCode & "," & kb.KeyCode2 & "," & Chr(34) & kb.KeyString & Chr(34)
        count = count + 1
    Next kb

    Close #fnum
    MsgBox count & " keyboard shortcut(s) exported to:" & vbCrLf & filePath, vbInformation, "Export Shortcuts Successful"
    Exit Sub

HandleErr:
    If fnum <> 0 Then Close #fnum
    MsgBox "Could not export keyboard shortcuts." & vbCrLf & "Error: " & Err.Description, vbCritical, "Export Shortcuts Error"
End Sub

Sub ImportKeyboardShortcuts()
    Dim filePath As Variant, fnum As Integer, ln As String, parts() As String, count As Integer, failCount As Integer
    On Error GoTo HandleErr

    filePath = Application.GetOpenFilename("CSV Files (*.csv), *.csv")
    If filePath = False Or filePath = "False" Then Exit Sub

    fnum = FreeFile
    Open filePath For Input As #fnum
    If Not EOF(fnum) Then Line Input #fnum, ln

    Application.CustomizationContext = NormalTemplate

    Do While Not EOF(fnum)
        Line Input #fnum, ln
        If Len(Trim(ln)) > 0 Then
            parts = Split(ln, ",")
            If UBound(parts) >= 3 Then
                On Error Resume Next
                Err.Clear
                Application.KeyBindings.Add KeyCategory:=CLng(parts(0)), Command:=parts(1), KeyCode:=CLng(parts(2)), KeyCode2:=IIf(parts(3) = "0" Or parts(3) = "", 0, CLng(parts(3)))
                If Err.Number = 0 Then count = count + 1 Else failCount = failCount + 1
                On Error GoTo HandleErr
            End If
        End If
    Loop

    Close #fnum
    MsgBox count & " shortcut(s) imported successfully." & IIf(failCount > 0, vbCrLf & failCount & " could not be applied.", ""), vbInformation, "Import Shortcuts Successful"
    Exit Sub

HandleErr:
    If fnum <> 0 Then Close #fnum
    MsgBox "Could not import keyboard shortcuts." & vbCrLf & "Error: " & Err.Description, vbCritical, "Import Shortcuts Error"
End Sub

Sub ResetAllKeyboardShortcuts()
    If MsgBox("This will reset ALL custom keyboard shortcuts to Word defaults. Continue?", vbYesNo + vbExclamation, "Reset Shortcuts") = vbYes Then
        Application.CustomizationContext = NormalTemplate
        Application.KeyBindings.ClearAll
        MsgBox "Keyboard shortcuts reset to Word defaults.", vbInformation, "Reset Shortcuts"
    End If
End Sub`,

'Toolkit_RibbonQAT.bas': `Attribute VB_Name = "Toolkit_RibbonQAT"
Option Explicit

Private Function OfficeUIPath() As String
    OfficeUIPath = Environ$("LOCALAPPDATA") & "\\Microsoft\\Office\\Word.officeUI"
End Function

Sub ExportRibbonCustomizations()
    Dim srcFile As String, destFolder As String, fso As Object
    srcFile = OfficeUIPath()
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(srcFile) Then
        MsgBox "No custom ribbon/QAT file was found at:" & vbCrLf & srcFile & vbCrLf & vbCrLf & "Use File > Options > Customize Ribbon > Import/Export in Word instead.", vbExclamation, "Export Ribbon/QAT"
        Exit Sub
    End If
    destFolder = GetFolder("Select a folder to export the ribbon/QAT customization to")
    If destFolder = "" Then Exit Sub
    On Error GoTo HandleErr
    fso.CopyFile srcFile, destFolder & "\\Word.officeUI", True
    MsgBox "Ribbon and Quick Access Toolbar customizations exported to:" & vbCrLf & destFolder, vbInformation, "Export Ribbon/QAT"
    Exit Sub
HandleErr:
    MsgBox "Could not export the ribbon/QAT file." & vbCrLf & "Error: " & Err.Description, vbCritical, "Export Ribbon/QAT Error"
End Sub

Sub ImportRibbonCustomizations()
    Dim srcFile As Variant, destFile As String, fso As Object
    srcFile = Application.GetOpenFilename("Office UI Files (*.officeUI), *.officeUI")
    If srcFile = False Or srcFile = "False" Then Exit Sub
    destFile = OfficeUIPath()
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error GoTo HandleErr
    fso.CopyFile srcFile, destFile, True
    MsgBox "Ribbon/QAT customization imported. Close and reopen Word for changes to take effect.", vbInformation, "Import Ribbon/QAT"
    Exit Sub
HandleErr:
    MsgBox "Could not import the ribbon/QAT file." & vbCrLf & "Error: " & Err.Description, vbCritical, "Import Ribbon/QAT Error"
End Sub`,

'Toolkit_Menu.bas': `Attribute VB_Name = "Toolkit_Menu"
Option Explicit

Sub ShowToolkitMenu()
    Dim choice As String, msg As String
    msg = "WORD CUSTOMIZATION TOOLKIT" & vbCrLf & vbCrLf & _
          "  1 - Export Macros (To Files)" & vbCrLf & _
          "  2 - Import Macros (Replace Existing)" & vbCrLf & _
          "  3 - Export Keyboard Shortcuts (.CSV)" & vbCrLf & _
          "  4 - Import Keyboard Shortcuts (.CSV)" & vbCrLf & _
          "  5 - Export Ribbon / QAT Customizations" & vbCrLf & _
          "  6 - Import Ribbon / QAT Customizations" & vbCrLf & vbCrLf & _
          "Enter an option number (1-6):"
    choice = InputBox(msg, "Word Customization Toolkit")
    Select Case choice
        Case "1": ExportAllMacros
        Case "2": ImportAllMacros
        Case "3": ExportKeyboardShortcuts
        Case "4": ImportKeyboardShortcuts
        Case "5": ExportRibbonCustomizations
        Case "6": ImportRibbonCustomizations
        Case ""
        Case Else
            MsgBox "Please enter a number from 1 to 6.", vbExclamation, "Word Customization Toolkit"
    End Select
End Sub`
};

/* ---------- DOM LOAD & EVENT LISTENERS ---------- */
document.addEventListener('DOMContentLoaded', () => {
  initTabs();
  initSettingsModal();
  initConnectorDownloads();
  initBatchDropZone();
  checkServerConnection();
});

function initTabs() {
  $$('#tabs button').forEach(btn => {
    btn.addEventListener('click', () => {
      $$('#tabs button').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      $$('.tab-content').forEach(s => s.style.display = 'none');
      const target = $('#tab-' + btn.dataset.tab);
      if (target) target.style.display = 'block';
    });
  });
}

/* ---------- SERVER CONNECTIVITY & LOCAL FALLBACK ---------- */
async function checkServerConnection() {
  const dot = $('#status-dot');
  const text = $('#status-text');
  try {
    const res = await fetch(`${apiBaseUrl}/health`, { cache: 'no-store' });
    if (res.ok) {
      const data = await res.json();
      isServerOnline = true;
      dot.className = 'status-dot';
      text.textContent = 'Render Server Online';
    } else {
      throw new Error('Server returned error status');
    }
  } catch (e) {
    isServerOnline = false;
    dot.className = 'status-dot offline';
    text.textContent = 'Local Mode (Offline)';
  }
  refreshAllData();
}

function initSettingsModal() {
  const modal = $('#settings-modal');
  const input = $('#server-url-input');
  
  $('#btn-open-settings').onclick = () => {
    input.value = apiBaseUrl;
    modal.style.display = 'flex';
  };
  $('#btn-close-settings').onclick = () => {
    modal.style.display = 'none';
  };
  $('#btn-save-settings').onclick = () => {
    let val = input.value.trim();
    if (!val) val = '/api';
    if (val.endsWith('/')) val = val.slice(0, -1);
    apiBaseUrl = val;
    localStorage.setItem('wt_server_url', apiBaseUrl);
    modal.style.display = 'none';
    showToast('Server settings updated');
    checkServerConnection();
  };
}

/* ---------- REFRESH DATA ---------- */
async function refreshAllData() {
  await Promise.all([refreshMacros(), refreshShortcuts(), refreshRibbon()]);
}

/* ---------- MACRO MANAGEMENT ---------- */
async function refreshMacros() {
  const listEl = $('#m-list');
  try {
    if (isServerOnline) {
      const res = await fetch(`${apiBaseUrl}/macros`);
      const data = await res.json();
      currentMacros = data.macros || [];
    } else {
      const raw = localStorage.getItem('wt_local_macros');
      currentMacros = raw ? JSON.parse(raw) : getSampleMacros();
    }
  } catch (e) {
    console.error('Failed to fetch macros:', e);
    currentMacros = getSampleMacros();
  }

  updateGroupSuggestionsAndFilter();
  renderGroupedMacros();
}

function getSampleMacros() {
  return [
    {
      id: 'm-sample-1',
      group: 'Formatting Tools',
      name: 'CleanFormatting',
      type: 'bas',
      code: `Attribute VB_Name = "CleanFormatting"\nOption Explicit\n\nSub RemoveExtraSpaces()\n    With Selection.Find\n        .ClearFormatting\n        .Replacement.ClearFormatting\n        .Text = "  "\n        .Replacement.Text = " "\n        .Forward = True\n        .Wrap = wdFindContinue\n        .Execute Replace:=wdReplaceAll\n    End With\n    MsgBox "Extra spaces removed!", vbInformation\nEnd Sub`,
      updatedAt: new Date().toISOString()
    }
  ];
}

function updateGroupSuggestionsAndFilter() {
  const datalist = $('#group-suggestions');
  const selectFilter = $('#m-filter-group');
  const selectedGroup = selectFilter.value;

  const groups = Array.from(new Set(currentMacros.map(m => m.group || 'General'))).sort();

  datalist.innerHTML = groups.map(g => `<option value="${escapeHtml(g)}">`).join('');

  selectFilter.innerHTML = '<option value="">All Groups</option>' + 
    groups.map(g => `<option value="${escapeHtml(g)}"${g === selectedGroup ? ' selected' : ''}>${escapeHtml(g)}</option>`).join('');
}

function renderGroupedMacros() {
  const listEl = $('#m-list');
  const searchQuery = ($('#m-search').value || '').toLowerCase().trim();
  const groupFilter = ($('#m-filter-group').value || '').toLowerCase().trim();

  let filtered = currentMacros.filter(m => {
    const groupName = (m.group || 'General').toLowerCase();
    const nameMatch = (m.name || '').toLowerCase().includes(searchQuery) ||
                      groupName.includes(searchQuery) ||
                      (m.code || '').toLowerCase().includes(searchQuery);
    const groupMatch = !groupFilter || groupName === groupFilter;
    return nameMatch && groupMatch;
  });

  if (filtered.length === 0) {
    listEl.innerHTML = `<div class="empty">No macros found. ${currentMacros.length === 0 ? 'Add a new macro above or import a .bas file.' : 'Try adjusting your search filter.'}</div>`;
    return;
  }

  // Group macros by Group Name
  const groupedMap = {};
  filtered.forEach(m => {
    const g = m.group || 'General';
    if (!groupedMap[g]) groupedMap[g] = [];
    groupedMap[g].push(m);
  });

  listEl.innerHTML = '';

  Object.keys(groupedMap).sort().forEach(groupName => {
    const items = groupedMap[groupName];
    const groupShortcuts = currentShortcuts.filter(sc =>
      (sc.group || 'General').toLowerCase() === groupName.toLowerCase()
    );

    const groupSec = document.createElement('div');
    groupSec.className = 'group-section';

    const header = document.createElement('div');
    header.className = 'group-header';
    header.innerHTML = `
      <div class="group-title-area">
        <h3 class="group-title">${escapeHtml(groupName)}</h3>
        <span class="group-count-badge">${items.length} macro${items.length === 1 ? '' : 's'}${groupShortcuts.length > 0 ? ` · ${groupShortcuts.length} shortcut set${groupShortcuts.length === 1 ? '' : 's'}` : ''}</span>
      </div>
      <button class="btn small secondary btn-export-group" data-group="${escapeHtml(groupName)}">Export Group (.bas)</button>
    `;

    const stampsDiv = document.createElement('div');
    stampsDiv.className = 'stamps';

    items.forEach(macro => {
      const stamp = document.createElement('div');
      stamp.className = 'stamp';
      stamp.innerHTML = `
        <div class="meta">
          <span class="badge">.${escapeHtml(macro.type || 'bas')}</span>
          <p class="name">${escapeHtml(macro.name)}</p>
          <span class="sub">Group: ${escapeHtml(macro.group || 'General')} · ${fmtDate(macro.updatedAt)}</span>
        </div>
        <div class="actions">
          <button class="btn small secondary" data-act="copy">Copy Code</button>
          <button class="btn small secondary" data-act="dl">Download</button>
          <button class="btn small secondary" data-act="edit">Edit</button>
          <button class="btn small ghost-danger" data-act="del">Delete</button>
        </div>
      `;

      stamp.querySelector('[data-act=copy]').onclick = () => {
        navigator.clipboard.writeText(macro.code);
        showToast(`Copied ${macro.name} to clipboard!`);
      };
      stamp.querySelector('[data-act=dl]').onclick = () => {
        downloadFile(`${macro.name}.${macro.type || 'bas'}`, macro.code);
      };
      stamp.querySelector('[data-act=edit]').onclick = () => {
        $('#m-group').value = macro.group || '';
        $('#m-name').value = macro.name;
        $('#m-type').value = macro.type || 'bas';
        $('#m-code').value = macro.code;
        window.scrollTo({ top: $('#tab-macros').offsetTop, behavior: 'smooth' });
        showToast(`Loaded ${macro.name} into editor`);
      };
      stamp.querySelector('[data-act=del]').onclick = async () => {
        if (confirm(`Delete macro "${macro.name}" from group "${macro.group}"?`)) {
          await deleteMacro(macro.id);
        }
      };

      stampsDiv.appendChild(stamp);
    });

    // Bundled shortcut sets for this group
    groupShortcuts.forEach(sc => {
      const rows = (sc.csv.match(/\n/g) || []).length;
      const stamp = document.createElement('div');
      stamp.className = 'stamp shortcut-stamp';
      stamp.innerHTML = `
        <div class="meta">
          <span class="badge">.CSV</span>
          <p class="name">${escapeHtml(sc.name)}</p>
          <span class="sub">${rows} shortcut(s) · ${fmtDate(sc.updatedAt)}</span>
        </div>
        <div class="actions">
          <button class="btn small secondary" data-act="sdl">Download CSV</button>
          <button class="btn small ghost-danger" data-act="sdel">Delete</button>
        </div>
      `;
      stamp.querySelector('[data-act=sdl]').onclick = () => downloadFile(`${sc.name}.csv`, sc.csv);
      stamp.querySelector('[data-act=sdel]').onclick = async () => {
        if (confirm(`Delete shortcut set "${sc.name}" from group "${groupName}"?`)) {
          await deleteShortcut(sc.id);
        }
      };
      stampsDiv.appendChild(stamp);
    });

    header.querySelector('.btn-export-group').onclick = (e) => {
      e.stopPropagation();
      exportMacroGroup(groupName, items);
    };

    groupSec.appendChild(header);
    groupSec.appendChild(stampsDiv);
    listEl.appendChild(groupSec);
  });
}

// Search and filter listeners
$('#m-search').addEventListener('input', renderGroupedMacros);
$('#m-filter-group').addEventListener('change', renderGroupedMacros);

// Save Macro Click Event
$('#m-save').addEventListener('click', async () => {
  const group = $('#m-group').value.trim() || 'General';
  const name = $('#m-name').value.trim();
  const type = $('#m-type').value;
  const code = $('#m-code').value;
  const status = $('#m-status');

  if (!name || !code.trim()) {
    status.textContent = 'Module name and code are required.';
    status.className = 'status-msg err';
    return;
  }

  status.textContent = 'Saving macro…';
  status.className = 'status-msg';

  try {
    const payload = { group, name, type, code };

    if (isServerOnline) {
      const res = await fetch(`${apiBaseUrl}/macros`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });
      const data = await res.json();
      if (!data.success) throw new Error(data.error || 'Server save failed');
      showToast(data.action === 'replaced' ? `Replaced existing macro "${name}"` : `Saved macro "${name}"`);

      // Auto-bundle: if a shortcut set is being imported with the same group, save it along with the macro
      await saveGroupShortcutSet(group);
    } else {
      // Offline LocalStorage Save
      const existingIdx = currentMacros.findIndex(m =>
        m.name.toLowerCase() === name.toLowerCase() &&
        (m.group || 'General').toLowerCase() === group.toLowerCase()
      );
      const macroObj = {
        id: existingIdx >= 0 ? currentMacros[existingIdx].id : 'macro-local-' + Date.now(),
        group, name, type, code, updatedAt: new Date().toISOString()
      };
      if (existingIdx >= 0) currentMacros[existingIdx] = macroObj;
      else currentMacros.push(macroObj);

      localStorage.setItem('wt_local_macros', JSON.stringify(currentMacros));
      showToast(existingIdx >= 0 ? `Replaced local macro "${name}"` : `Saved local macro "${name}"`);
      await saveGroupShortcutSet(group);
    }

    status.textContent = 'Saved successfully!';
    status.className = 'status-msg';
    $('#m-name').value = '';
    $('#m-code').value = '';
    refreshMacros();
  } catch (e) {
    status.textContent = 'Save failed: ' + e.message;
    status.className = 'status-msg err';
  }
});

async function deleteMacro(id) {
  try {
    if (isServerOnline) {
      await fetch(`${apiBaseUrl}/macros/${id}`, { method: 'DELETE' });
    } else {
      currentMacros = currentMacros.filter(m => m.id !== id);
      localStorage.setItem('wt_local_macros', JSON.stringify(currentMacros));
    }
    showToast('Macro deleted');
    refreshMacros();
  } catch (e) {
    showToast('Could not delete macro: ' + e.message);
  }
}

function exportMacroGroup(groupName, items) {
  const bundledCode = items.map(m => `' ==========================================\n' Module: ${m.name}.${m.type}\n' Group: ${groupName}\n' ==========================================\n${m.code}`).join('\n\n');
  downloadFile(`${groupName.replace(/[^a-zA-Z0-9_]/g, '_')}_Macros.bas`, bundledCode);
  showToast(`Exported ${items.length} macro(s) from "${groupName}"`);
}

/* ---------- SHORTCUTS MANAGEMENT ---------- */
async function refreshShortcuts() {
  const listEl = $('#s-list');
  try {
    if (isServerOnline) {
      const res = await fetch(`${apiBaseUrl}/shortcuts`);
      const data = await res.json();
      currentShortcuts = data.shortcuts || [];
    } else {
      const raw = localStorage.getItem('wt_local_shortcuts');
      currentShortcuts = raw ? JSON.parse(raw) : [];
    }
  } catch (e) {
    currentShortcuts = [];
  }

  if (currentShortcuts.length === 0) {
    listEl.innerHTML = '<div class="empty">No shortcut sets saved yet.</div>';
    return;
  }

  listEl.innerHTML = '';
  currentShortcuts.forEach(sc => {
    const rows = (sc.csv.match(/\n/g) || []).length;
    const el = document.createElement('div');
    el.className = 'stamp';
    el.innerHTML = `
      <div class="meta">
        <span class="badge">.CSV</span>
        <p class="name">${escapeHtml(sc.name)}</p>
        <span class="sub">Group: ${escapeHtml(sc.group || 'General')} · ${rows} shortcut(s) · ${fmtDate(sc.updatedAt)}</span>
      </div>
      <div class="actions">
        <button class="btn small secondary" data-act="dl">Download</button>
        <button class="btn small ghost-danger" data-act="del">Delete</button>
      </div>`;
    el.querySelector('[data-act=dl]').onclick = () => downloadFile(`${sc.name}.csv`, sc.csv);
    el.querySelector('[data-act=del]').onclick = () => deleteShortcut(sc.id);
    listEl.appendChild(el);
  });
}

$('#s-save').addEventListener('click', async () => {
  const name = $('#s-name').value.trim();
  const group = $('#s-group').value.trim();
  const csv = $('#s-csv').value;
  const status = $('#s-status');

  if (!name || !csv.trim()) {
    status.textContent = 'Set name and CSV data are required.';
    status.className = 'status-msg err';
    return;
  }

  status.textContent = 'Saving shortcut set…';
  try {
    if (isServerOnline) {
      await fetch(`${apiBaseUrl}/shortcuts`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ name, group, csv })
      });
    } else {
      currentShortcuts.push({ id: 'sc-' + Date.now(), name, group: group || 'General', csv, updatedAt: new Date().toISOString() });
      localStorage.setItem('wt_local_shortcuts', JSON.stringify(currentShortcuts));
    }
    status.textContent = 'Saved!';
    status.className = 'status-msg';
    $('#s-name').value = ''; $('#s-csv').value = '';
    showToast('Saved shortcut set');
    refreshShortcuts();
  } catch (e) {
    status.textContent = 'Save failed: ' + e.message;
    status.className = 'status-msg err';
  }
});

async function deleteShortcut(id) {
  if (isServerOnline) {
    await fetch(`${apiBaseUrl}/shortcuts/${id}`, { method: 'DELETE' });
  } else {
    currentShortcuts = currentShortcuts.filter(s => s.id !== id);
    localStorage.setItem('wt_local_shortcuts', JSON.stringify(currentShortcuts));
  }
  showToast('Shortcut set deleted');
  refreshShortcuts();
}

/* ---------- RIBBON MANAGEMENT ---------- */
async function refreshRibbon() {
  const listEl = $('#r-list');
  try {
    if (isServerOnline) {
      const res = await fetch(`${apiBaseUrl}/ribbon`);
      const data = await res.json();
      currentRibbon = data.ribbon || [];
    } else {
      const raw = localStorage.getItem('wt_local_ribbon');
      currentRibbon = raw ? JSON.parse(raw) : [];
    }
  } catch (e) {
    currentRibbon = [];
  }

  if (currentRibbon.length === 0) {
    listEl.innerHTML = '<div class="empty">No ribbon customization profiles saved yet.</div>';
    return;
  }

  listEl.innerHTML = '';
  currentRibbon.forEach(rb => {
    const el = document.createElement('div');
    el.className = 'stamp';
    el.innerHTML = `
      <div class="meta">
        <span class="badge">.officeUI</span>
        <p class="name">${escapeHtml(rb.name)}</p>
        <span class="sub">${escapeHtml(rb.filename || 'Word.officeUI')} · ${fmtDate(rb.updatedAt)}</span>
      </div>
      <div class="actions">
        <button class="btn small secondary" data-act="dl">Download</button>
        <button class="btn small ghost-danger" data-act="del">Delete</button>
      </div>`;
    el.querySelector('[data-act=dl]').onclick = () => downloadBase64(rb.filename || 'Word.officeUI', rb.base64);
    el.querySelector('[data-act=del]').onclick = () => deleteRibbon(rb.id);
    listEl.appendChild(el);
  });
}

$('#r-save').addEventListener('click', () => {
  const name = $('#r-name').value.trim();
  const file = $('#r-file').files[0];
  const status = $('#r-status');

  if (!name || !file) {
    status.textContent = 'Name and file are required.';
    status.className = 'status-msg err';
    return;
  }

  status.textContent = 'Reading & saving file…';
  const reader = new FileReader();
  reader.onload = async () => {
    try {
      const base64 = reader.result.split(',')[1];
      if (isServerOnline) {
        await fetch(`${apiBaseUrl}/ribbon`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ name, filename: file.name, base64 })
        });
      } else {
        currentRibbon.push({ id: 'rb-' + Date.now(), name, filename: file.name, base64, updatedAt: new Date().toISOString() });
        localStorage.setItem('wt_local_ribbon', JSON.stringify(currentRibbon));
      }
      status.textContent = 'Saved!';
      status.className = 'status-msg';
      $('#r-name').value = ''; $('#r-file').value = '';
      showToast('Saved ribbon profile');
      refreshRibbon();
    } catch (e) {
      status.textContent = 'Save failed: ' + e.message;
      status.className = 'status-msg err';
    }
  };
  reader.readAsDataURL(file);
});

async function deleteRibbon(id) {
  if (isServerOnline) {
    await fetch(`${apiBaseUrl}/ribbon/${id}`, { method: 'DELETE' });
  } else {
    currentRibbon = currentRibbon.filter(r => r.id !== id);
    localStorage.setItem('wt_local_ribbon', JSON.stringify(currentRibbon));
  }
  showToast('Ribbon profile deleted');
  refreshRibbon();
}

/* ---------- IMPORT / EXPORT HUB ---------- */

// Export full backup JSON
$('#io-export-all').addEventListener('click', async () => {
  if (isServerOnline) {
    window.location.href = `${apiBaseUrl}/export`;
  } else {
    const backupPkg = {
      appName: 'Word Shortcuts & Macro Manager',
      version: '1.0',
      exportedAt: new Date().toISOString(),
      macros: currentMacros,
      shortcuts: currentShortcuts,
      ribbon: currentRibbon
    };
    downloadFile(`word-toolkit-backup-${Date.now()}.json`, JSON.stringify(backupPkg, null, 2));
    showToast('Exported system backup package');
  }
});

// Restore System Backup JSON
$('#io-btn-restore').addEventListener('click', () => {
  const file = $('#io-restore-file').files[0];
  const mode = $('#io-restore-mode').value;

  if (!file) {
    alert('Please select a system backup .json file first.');
    return;
  }

  const reader = new FileReader();
  reader.onload = async () => {
    try {
      const parsedData = JSON.parse(reader.result);
      if (isServerOnline) {
        const res = await fetch(`${apiBaseUrl}/import`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ mode, data: parsedData })
        });
        const rData = await res.json();
        if (!rData.success) throw new Error(rData.error);
        showToast(rData.message);
      } else {
        // Offline Restore
        if (mode === 'replace') {
          currentMacros = parsedData.macros || [];
          currentShortcuts = parsedData.shortcuts || [];
          currentRibbon = parsedData.ribbon || [];
        } else {
          // Merge
          (parsedData.macros || []).forEach(newM => {
            const idx = currentMacros.findIndex(m => m.name.toLowerCase() === (newM.name || '').toLowerCase());
            if (idx >= 0) currentMacros[idx] = newM;
            else currentMacros.push(newM);
          });
        }
        localStorage.setItem('wt_local_macros', JSON.stringify(currentMacros));
        showToast('Local system backup restored!');
      }
      $('#io-restore-file').value = '';
      refreshAllData();
    } catch (e) {
      alert('Error parsing backup file: ' + e.message);
    }
  };
  reader.readAsText(file);
});

/* ---------- BATCH DRAG & DROP ZONE ---------- */
function initBatchDropZone() {
  const zone = $('#batch-drop-zone');
  const fileInput = $('#batch-file-input');

  zone.onclick = () => fileInput.click();

  zone.ondragover = (e) => { e.preventDefault(); zone.classList.add('dragover'); };
  zone.ondragleave = () => zone.classList.remove('dragover');
  zone.ondrop = (e) => {
    e.preventDefault();
    zone.classList.remove('dragover');
    if (e.dataTransfer.files.length) handleBatchFiles(e.dataTransfer.files);
  };

  fileInput.onchange = () => {
    if (fileInput.files.length) handleBatchFiles(fileInput.files);
  };
}

async function handleBatchFiles(files) {
  let count = 0;
  for (const file of Array.from(files)) {
    const ext = file.name.split('.').pop().toLowerCase();
    if (['bas', 'cls', 'frm'].includes(ext)) {
      const text = await file.text();
      // Parse module name from Attribute VB_Name if present
      let modName = file.name.replace(/\.[^/.]+$/, "");
      const match = text.match(/Attribute\s+VB_Name\s*=\s*"([^"]+)"/i);
      if (match && match[1]) modName = match[1];

      await saveMacroItem({
        group: 'Imported Files',
        name: modName,
        type: ext,
        code: text
      });
      count++;
    } else if (ext === 'csv') {
      const text = await file.text();
      await saveShortcutItem({
        name: file.name.replace(/\.csv$/i, ''),
        csv: text
      });
      count++;
    } else if (ext === 'officeui') {
      const base64 = await readFileAsBase64(file);
      await saveRibbonItem({
        name: file.name.replace(/\.officeui$/i, ''),
        filename: file.name,
        base64: base64
      });
      count++;
    }
  }
  showToast(`Successfully processed ${count} file(s)!`);
  refreshAllData();
}

async function saveMacroItem(item) {
  if (isServerOnline) {
    await fetch(`${apiBaseUrl}/macros`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(item)
    });
  } else {
    const idx = currentMacros.findIndex(m => m.name.toLowerCase() === item.name.toLowerCase());
    const macroObj = { ...item, id: 'macro-local-' + Date.now(), updatedAt: new Date().toISOString() };
    if (idx >= 0) currentMacros[idx] = macroObj;
    else currentMacros.push(macroObj);
    localStorage.setItem('wt_local_macros', JSON.stringify(currentMacros));
  }
}

async function saveShortcutItem(item) {
  if (isServerOnline) {
    await fetch(`${apiBaseUrl}/shortcuts`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(item)
    });
  } else {
    currentShortcuts.push({ ...item, id: 'sc-' + Date.now(), updatedAt: new Date().toISOString() });
    localStorage.setItem('wt_local_shortcuts', JSON.stringify(currentShortcuts));
  }
}

// When a macro is saved to a group, also persist any shortcut set the user
// has prepared for that same group (bundled under the same group name).
async function saveGroupShortcutSet(group) {
  const scName = $('#s-name').value.trim();
  const scGroup = $('#s-group').value.trim();
  const scCsv = $('#s-csv').value;
  if (!scName || !scCsv.trim()) return;

  const targetGroup = scGroup || group || 'General';
  if ((scGroup || 'General').toLowerCase() !== (group || 'General').toLowerCase()) return;

  try {
    if (isServerOnline) {
      await fetch(`${apiBaseUrl}/shortcuts`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ name: scName, group: targetGroup, csv: scCsv })
      });
    } else {
      currentShortcuts.push({ id: 'sc-' + Date.now(), name: scName, group: targetGroup, csv: scCsv, updatedAt: new Date().toISOString() });
      localStorage.setItem('wt_local_shortcuts', JSON.stringify(currentShortcuts));
    }
    $('#s-name').value = ''; $('#s-group').value = ''; $('#s-csv').value = '';
    showToast(`Bundled shortcut set "${scName}" into "${targetGroup}"`);
    refreshShortcuts();
  } catch (e) {
    showToast('Shortcut set bundled, but sync failed: ' + e.message);
  }
}

async function saveRibbonItem(item) {
  if (isServerOnline) {
    await fetch(`${apiBaseUrl}/ribbon`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(item)
    });
  } else {
    currentRibbon.push({ ...item, id: 'rb-' + Date.now(), updatedAt: new Date().toISOString() });
    localStorage.setItem('wt_local_ribbon', JSON.stringify(currentRibbon));
  }
}

/* ---------- CONNECTOR DOWNLOADS ---------- */
function initConnectorDownloads() {
  const grid = $('#connector-downloads');
  grid.innerHTML = '';
  Object.keys(CONNECTOR_FILES).forEach(fname => {
    const btn = document.createElement('button');
    btn.className = 'btn secondary';
    btn.textContent = '↓ Download ' + fname;
    btn.onclick = () => downloadFile(fname, CONNECTOR_FILES[fname]);
    grid.appendChild(btn);
  });
}

/* ---------- HELPERS ---------- */
function downloadFile(filename, content) {
  const blob = new Blob([content], { type: 'text/plain;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click(); a.remove();
  URL.revokeObjectURL(url);
}

function downloadBase64(filename, base64Str) {
  const bin = atob(base64Str);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  const blob = new Blob([bytes], { type: 'application/octet-stream' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click(); a.remove();
  URL.revokeObjectURL(url);
}

function readFileAsBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result.split(',')[1]);
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

function escapeHtml(str) {
  if (!str) return '';
  return String(str).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

function fmtDate(isoStr) {
  if (!isoStr) return '';
  try {
    return new Date(isoStr).toLocaleString(undefined, { dateStyle: 'medium', timeStyle: 'short' });
  } catch (e) {
    return '';
  }
}

function showToast(msg) {
  const existing = $('.toast');
  if (existing) existing.remove();
  const toast = document.createElement('div');
  toast.className = 'toast';
  toast.textContent = msg;
  document.body.appendChild(toast);
  setTimeout(() => {
    toast.style.opacity = '0';
    setTimeout(() => toast.remove(), 300);
  }, 2800);
}
