const RAW = [
  ['Ctrl+Q', 593, 1, 'TableMergeCells'],
  ['Ctrl+T', 596, 1, 'TextToTable'],
  ['Ctrl+Num +', 619, 1, 'InsertEmSpace'],
  ['Ctrl+Shift+B', 834, 1, 'BorderLineStyle'],
  ['Ctrl+Shift+C', 835, 1, 'CopyFormat'],
  ['Ctrl+Shift+D', 836, 1, 'FormatParagraph'],
  ['Ctrl+Shift+F', 838, 1, 'FontColorPicker'],
  ['Ctrl+Shift+T', 852, 1, 'TableProperties'],
  ['Ctrl+Shift+V', 854, 1, 'PasteFormat'],
  ['Ctrl+Shift+Num +', 875, 1, 'InsertEnSpace'],
  ['Alt+S', 1107, 1, 'TableSplitCells'],
  ['Alt+X', 1112, 1, 'TableSplit'],
  ['Alt+Ctrl+V', 1622, 1, 'EditPasteSpecial'],
  ['Alt+Ctrl+Shift+Left', 1829, 1, 'DecreaseIndent'],
  ['Alt+Ctrl+Shift+Right', 1831, 1, 'IncreaseIndent'],
  ['Ctrl+M', 589, 2, 'Normal.MergeFileTool.UltimateExactMerge'],
  ['Ctrl+Num 0', 608, 2, 'Normal.setcolumnlevels.reset_space'],
  ['Ctrl+Num 3', 611, 2, 'Normal.ReFLinkBookmark.AllInsertBookmarks_ToSelection'],
  ['Ctrl+Num .', 622, 2, 'Normal.modREY_Bookmarks.InsertBookmark_CurrentParagraph'],
  ['Ctrl+Shift+P', 848, 2, 'Normal.PasteOptionSelector.PasteOptionSelector'],
  ['Ctrl+Shift+X', 856, 2, 'Normal.Replace_images.ReplaceSelectedPicture_SessionFolder'],
  ['Alt+Up', 1062, 2, 'Normal.Module3.DecrementLeading'],
  ['Alt+Down', 1064, 2, 'Normal.Module3.IncrementLeading'],
  ['Alt+A', 1089, 2, 'Normal.Module3.InsertAndFormatPicture'],
  ['Alt+C', 1091, 2, 'Normal.Module3.ShowCleanupTool'],
  ['Alt+D', 1092, 2, 'Normal.modHideUnhide.ShowHideUnhideGUI'],
  ['Alt+Shift+A', 1345, 2, 'Normal.Module1.ShowPlaceholderPicker'],
  ['Alt+Shift+S', 1363, 2, 'Normal.Module1.ShowPlaceholderPicker1'],
  ['Alt+Ctrl+Left', 1573, 2, 'Normal.setcolumnlevels.after_space_DOWN'],
  ['Alt+Ctrl+Up', 1574, 2, 'Normal.setcolumnlevels.before_space_up'],
  ['Alt+Ctrl+Right', 1575, 2, 'Normal.setcolumnlevels.after_space_UP'],
  ['Alt+Ctrl+Down', 1576, 2, 'Normal.setcolumnlevels.before_space_DOWN'],
  ['Alt+Ctrl+G', 1607, 2, 'Normal.GroupSelectedTextboxes.GroupTextboxesInSelectedParagraph'],
  ['Alt+Ctrl+X', 1624, 2, 'Normal.KeyboardshortcutsExToImp.KeyboardShortcutsMenu'],
  ['Alt+Ctrl+]', 1757, 2, 'Normal.TextboxCleaner.CleanSelectedTextboxes'],
  ['Alt+Ctrl+Shift+Up', 1830, 2, 'Normal.indent.RightIndentIncrease'],
  ['Alt+Ctrl+Shift+Down', 1832, 2, 'Normal.indent.RightIndentDecrease']
];

const lines = ['KeyCategory,Command,KeyCode,KeyCode2,KeyString'];
RAW.forEach(r => lines.push(`${r[2]},${r[3]},${r[1]},0,"${r[0]}"`));
const csv = lines.join('\n');
console.log(JSON.stringify(csv).slice(0, 200));
console.log('TOTAL_LINES:', lines.length);
require('fs').writeFileSync('kishore-seed.txt', csv, 'utf8');