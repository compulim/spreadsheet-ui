import { WorksheetSelection } from '@compulim/spreadsheet-ui';
import { BaseStyles, PageLayout, ThemeProvider } from '@primer/react';
import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import type { ChangeEvent, KeyboardEvent as ReactKeyboardEvent } from 'react';
import { formatValue } from '../formatValue.ts';
import { getColumnName } from '../getColumnName.ts';
import { parseHighlight, stringifyHighlight } from '../highlight.ts';
import { HIGHLIGHTS } from '../highlights.ts';
import { cellsToPNG } from '../printCells.ts';
import { MAX_WORKSHEETS } from '../workbook/constants.ts';
import { Workbook } from '../workbook/Workbook.ts';
import { SheetBar } from './SheetBar.tsx';
import { Toolbar } from './Toolbar.tsx';
import type { BorderPreset, CellStyle, ToolbarActions, ToolbarState } from './Toolbar.tsx';
import type {
  BorderEdges,
  CellAlign,
  CellFormatting,
  CellNumberFormat,
  FormattingUpdater,
  Worksheet,
  WorksheetApp
} from '@compulim/spreadsheet-ui';

/**
 * Number of leading rows and columns that become read-only, and frozen on screen, when the lock button is pressed.
 * Change these to protect a different range, e.g. 2 header rows and 1 ID column.
 */
const LOCKED_ROW_COUNT = 1;
const LOCKED_COLUMN_COUNT = 1;

const ALIGNS: readonly CellAlign[] = ['left', 'center', 'right'];
const FORMATS: readonly CellNumberFormat[] = ['currency', 'percent', 'date', 'datetime', 'duration'];
const STYLES: readonly CellStyle[] = ['bold', 'italic', 'underline', 'strikethrough'];

function download(blob: Blob, fileName: string) {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');

  anchor.href = url;
  anchor.download = fileName;
  document.body.append(anchor);
  anchor.click();
  anchor.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function isTextField(target: EventTarget | null) {
  return (
    target instanceof HTMLInputElement ||
    target instanceof HTMLTextAreaElement ||
    (target instanceof HTMLElement && target.isContentEditable)
  );
}

/**
 * Whether the predicate holds for the formatting of every contentful cell in the selection.
 * If the selection has no contentful cells, checks the active cell instead.
 */
function isFormattingAppliedToSelection(
  worksheet: Worksheet,
  predicate: (formatting: CellFormatting | undefined) => boolean
) {
  const { activeCell, range } = worksheet.selection;
  const width = range.right - range.left + 1;
  const height = range.bottom - range.top + 1;
  const formattings = worksheet.getFormattings(range.left, range.top, width, height);
  let hasContent = false;

  for (let row = range.top; row <= range.bottom; row++) {
    for (let column = range.left; column <= range.right; column++) {
      if ((worksheet.getValue(column, row) ?? '') !== '') {
        hasContent = true;

        if (!predicate(formattings[row - range.top]?.[column - range.left])) {
          return false;
        }
      }
    }
  }

  return hasContent || predicate(worksheet.getFormattings(activeCell.column, activeCell.row, 1, 1)[0]![0]);
}

/** Whether every selected cell is in a merged cell. */
function isSelectionMerged(worksheet: Worksheet) {
  const { left, top, right, bottom } = worksheet.selection.range;
  let covered = 0;

  for (const { x, y, width, height } of worksheet.getMerges(left, top, right - left + 1, bottom - top + 1)) {
    covered +=
      (Math.min(x + width - 1, right) - Math.max(x, left) + 1) *
      (Math.min(y + height - 1, bottom) - Math.max(y, top) + 1);
  }

  // Merged cells never overlap, so all cells are merged if they cover the whole selection.
  return covered === (right - left + 1) * (bottom - top + 1);
}

function createWorkbook() {
  const workbook = new Workbook();

  workbook.addWorksheet();

  return workbook;
}

function App() {
  /**
   * The single `<worksheet-app>`, editing a copy of the active worksheet. Its worksheet is loaded from the active
   * worksheet of the workbook when activated, and every change is saved back to the workbook.
   */
  const elementRef = useRef<WorksheetApp | undefined>(undefined);

  if (!elementRef.current) {
    elementRef.current = document.createElement('worksheet-app') as WorksheetApp;
    elementRef.current.setAttribute('locked-column-count', String(LOCKED_COLUMN_COUNT));
    elementRef.current.setAttribute('locked-row-count', String(LOCKED_ROW_COUNT));
  }

  const element = elementRef.current;

  const [workbook, setWorkbookState] = useState(createWorkbook);
  const [activeWorksheet, setActiveWorksheet] = useState<Worksheet | undefined>(() => workbook.worksheets[0]);
  const [locked, setLocked] = useState(false);
  const [status, setStatus] = useState('');
  /** Bumped when the worksheet of `element`, or its selection, changes. */
  const [revision, setRevision] = useState(0);
  /** Bumped when worksheets of the workbook are added, removed or renamed. */
  const [tabsRevision, setTabsRevision] = useState(0);

  const activeWorksheetRef = useRef(activeWorksheet);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const gridContainerRef = useRef<HTMLDivElement>(null);
  /** Whether `element` is loading the active worksheet, its changes do not need to be saved back. */
  const loadingRef = useRef(false);
  const lockedRef = useRef(locked);
  // Menus return focus to the grid after closing.
  const returnFocusRef = useRef<HTMLElement>(element);
  const statusTimeoutRef = useRef<ReturnType<typeof setTimeout> | undefined>(undefined);
  const workbookRef = useRef(workbook);

  activeWorksheetRef.current = activeWorksheet;
  lockedRef.current = locked;
  workbookRef.current = workbook;

  const refresh = useCallback(() => setRevision(revision => revision + 1), []);

  const showStatus = useCallback((message: string) => {
    clearTimeout(statusTimeoutRef.current);
    setStatus(message);
    statusTimeoutRef.current = setTimeout(() => setStatus(''), 4000);
  }, []);

  useEffect(() => () => clearTimeout(statusTimeoutRef.current), []);

  useEffect(() => {
    gridContainerRef.current?.replaceChildren(element);
  }, [element]);

  useEffect(() => {
    const worksheet = element.worksheet;
    const save = () => {
      const activeWorksheet = activeWorksheetRef.current;

      if (!loadingRef.current && activeWorksheet && workbookRef.current.indexOf(activeWorksheet) !== -1) {
        const json = worksheet.toJSON();

        activeWorksheet.edit(edit => edit.loadJSON(json));
      }
    };
    const handleChange = () => {
      save();
      refresh();
    };
    const handleLockEdit = () => showStatus('Some cells are locked and cannot be edited.');

    worksheet.addEventListener('change', handleChange);
    worksheet.addEventListener('columnwidthchange', save);
    worksheet.addEventListener('selectionchange', refresh);
    worksheet.addEventListener('historychange', refresh);
    element.addEventListener('lockedit', handleLockEdit);

    return () => {
      worksheet.removeEventListener('change', handleChange);
      worksheet.removeEventListener('columnwidthchange', save);
      worksheet.removeEventListener('selectionchange', refresh);
      worksheet.removeEventListener('historychange', refresh);
      element.removeEventListener('lockedit', handleLockEdit);
    };
  }, [element, refresh, showStatus]);

  useEffect(() => {
    const handleChange = () => setTabsRevision(revision => revision + 1);

    workbook.addEventListener('worksheetadd', handleChange);
    workbook.addEventListener('worksheetremove', handleChange);
    workbook.addEventListener('worksheetrename', handleChange);

    // Expose for scripting from the browser console, e.g. `app.workbook.worksheets[0].edit(edit => edit.setValue(0, 0, 'Hello'))`.
    Object.assign(globalThis, { app: { workbook } });

    return () => {
      workbook.removeEventListener('worksheetadd', handleChange);
      workbook.removeEventListener('worksheetremove', handleChange);
      workbook.removeEventListener('worksheetrename', handleChange);
    };
  }, [workbook]);

  // Loads the active worksheet into `element`.
  useEffect(() => {
    if (!activeWorksheet) {
      return;
    }

    const json = activeWorksheet.toJSON();

    loadingRef.current = true;

    try {
      element.worksheet.edit(edit => edit.loadJSON(json));
    } finally {
      loadingRef.current = false;
    }

    // Fires "selectionchange", which refreshes the toolbar.
    element.worksheet.selection = new WorksheetSelection({ column: 0, row: 0 });
    element.focus();
  }, [activeWorksheet, element]);

  useEffect(() => {
    element.toggleAttribute('locked', locked);
  }, [element, locked]);

  const activateWorksheet = useCallback(
    (worksheet: Worksheet) => {
      setActiveWorksheet(worksheet);
      element.focus();
    },
    [element]
  );

  const replaceWorkbook = useCallback((nextWorkbook: Workbook) => {
    if (!nextWorkbook.worksheets.length) {
      nextWorkbook.addWorksheet();
    }

    setWorkbookState(nextWorkbook);
    setActiveWorksheet(nextWorkbook.worksheets[0]);
  }, []);

  const formatSelection = useCallback(
    (updater: FormattingUpdater) => {
      const worksheet = element.worksheet;
      const { left, top, right, bottom } = worksheet.selection.range;

      worksheet.edit(edit => edit.setFormattings(updater, left, top, right - left + 1, bottom - top + 1));
      refresh();
      element.focus();
    },
    [element, refresh]
  );

  const toggleAlign = useCallback(
    (align: CellAlign) => {
      // Toggling a pushed alignment goes back to the default (type-based) alignment.
      const nextAlign = isFormattingAppliedToSelection(element.worksheet, formatting => formatting?.align === align)
        ? undefined
        : align;

      formatSelection(formatting => ({ ...formatting, align: nextAlign }));
    },
    [element, formatSelection]
  );

  const toggleStyle = useCallback(
    (style: CellStyle) => {
      const value = !isFormattingAppliedToSelection(element.worksheet, formatting => !!formatting?.[style]);

      formatSelection(formatting => ({ ...formatting, [style]: value }));
    },
    [element, formatSelection]
  );

  const mergeSelection = useCallback(() => {
    const worksheet = element.worksheet;
    const { left, top, right, bottom } = worksheet.selection.range;
    const isEditable = (column: number, row: number) =>
      !(lockedRef.current && (column < LOCKED_COLUMN_COUNT || row < LOCKED_ROW_COUNT));

    // The selection always covers whole merged cells.
    if (!isEditable(left, top) || !isEditable(right, bottom)) {
      showStatus('Some cells are locked and cannot be merged.');
    } else if (isSelectionMerged(worksheet)) {
      // Unmerges only if all selected cells are merged, otherwise merges all of them.
      worksheet.edit(edit =>
        worksheet
          .getMerges(left, top, right - left + 1, bottom - top + 1)
          .forEach(({ x, y }) => edit.unmergeCells(x, y))
      );

      // Keeps all cells of the unmerged cells selected, the selection may be only the merged cell itself.
      worksheet.selection = new WorksheetSelection({ column: left, row: top }, { left, top, right, bottom });
    } else {
      worksheet.edit(edit => edit.mergeCells(left, top, right - left + 1, bottom - top + 1));
    }

    refresh();
    element.focus();
  }, [element, refresh, showStatus]);

  const undoOrRedo = useCallback(
    (action: 'undo' | 'redo') => {
      element.worksheet[action]() && refresh();
      element.focus();
    },
    [element, refresh]
  );

  const setBorder = useCallback(
    (preset: BorderPreset) => {
      const worksheet = element.worksheet;
      const { left, top, right, bottom } = worksheet.selection.range;
      const range = { x: left, y: top, width: right - left + 1, height: bottom - top + 1 };

      worksheet.edit(edit => {
        if (preset === 'all' || preset === 'none') {
          edit.setBorder(range, new Set(['outside', 'inside']), preset === 'all');
        } else {
          edit.setBorder(range, new Set([preset as BorderEdges]));
        }
      });

      refresh();
      element.focus();
    },
    [element, refresh]
  );

  const saveCSV = useCallback(() => {
    // Byte order mark so Excel recognizes the file as UTF-8.
    download(
      new Blob(['\ufeff', element.worksheet.toCSV()], { type: 'text/csv' }),
      `${workbook.getWorksheetName(activeWorksheet!)}.csv`
    );
  }, [activeWorksheet, element, workbook]);

  const saveZip = useCallback(async () => {
    try {
      download(new Blob([(await workbook.toZip()) as BlobPart], { type: 'application/zip' }), 'workbook.zip');
    } catch (error) {
      alert(`Failed to save workbook: ${(error as Error).message}`);
    }
  }, [workbook]);

  const printSelection = useCallback(async () => {
    const worksheet = element.worksheet;
    const { left, top, right, bottom } = worksheet.selection.range;
    const rows = worksheet.getValues(left, top, right - left + 1, bottom - top + 1);

    try {
      const getBorder = (column: number, row: number) =>
        left + column >= 0 && top + row >= 0 && left + column < worksheet.columnCount && top + row < worksheet.rowCount
          ? worksheet.getBorder(left + column, top + row)
          : undefined;

      const columnWidths: number[] = [];

      for (let column = left; column <= right; column++) {
        columnWidths.push(worksheet.getColumnWidth(column));
      }

      const formattings = worksheet.getFormattings(left, top, right - left + 1, bottom - top + 1);
      /** Column and row spans of merged cells, keyed by `row,column`, 0 for cells not in the first column or row. */
      const spans = new Map<string, { columns: number; rows: number }>();

      // The selection always covers whole merged cells.
      for (const merge of worksheet.getMerges(left, top, right - left + 1, bottom - top + 1)) {
        for (let row = merge.y; row < merge.y + merge.height; row++) {
          for (let column = merge.x; column < merge.x + merge.width; column++) {
            spans.set(`${row - top},${column - left}`, {
              columns: column === merge.x ? merge.width : 0,
              rows: row === merge.y ? merge.height : 0
            });
          }
        }
      }

      const getColumnSpan = (column: number, row: number) => spans.get(`${row},${column}`)?.columns ?? 1;
      const getRowSpan = (column: number, row: number) => spans.get(`${row},${column}`)?.rows ?? 1;
      const getFormatting = (column: number, row: number) => formattings[row]?.[column];

      download(
        await cellsToPNG(rows, { columnWidths, getBorder, getColumnSpan, getFormatting, getRowSpan }),
        `${workbook.getWorksheetName(activeWorksheet!)}.png`
      );
    } catch (error) {
      alert(`Failed to print selected cells: ${(error as Error).message}`);
    }
  }, [activeWorksheet, element, workbook]);

  const handleLockedChange = useCallback(
    (nextLocked: boolean) => {
      setLocked(nextLocked);
      element.focus();
    },
    [element]
  );

  const toolbarActions = useMemo<ToolbarActions>(
    () => ({
      merge: mergeSelection,
      open: () => fileInputRef.current?.click(),
      print: printSelection,
      redo: () => undoOrRedo('redo'),
      saveCSV,
      saveZip,
      setAlign: align => formatSelection(formatting => ({ ...formatting, align })),
      setBorder,
      setFormat: format => formatSelection(formatting => ({ ...formatting, format })),
      setHighlight: highlight =>
        formatSelection(formatting => ({ ...formatting, highlight: parseHighlight(highlight) })),
      setLocked: handleLockedChange,
      toggleStyle,
      undo: () => undoOrRedo('undo'),
      unformat: () => formatSelection(() => null)
    }),
    [
      formatSelection,
      handleLockedChange,
      mergeSelection,
      printSelection,
      saveCSV,
      saveZip,
      setBorder,
      toggleStyle,
      undoOrRedo
    ]
  );

  const handleGridKeyDown = useCallback(
    (event: ReactKeyboardEvent<HTMLDivElement>) => {
      if ((event.ctrlKey || event.metaKey) && !event.altKey && !event.shiftKey && !isTextField(event.target)) {
        const action = {
          b: () => toggleStyle('bold'),
          i: () => toggleStyle('italic'),
          u: () => toggleStyle('underline'),
          l: () => toggleAlign('left'),
          e: () => toggleAlign('center'),
          r: () => toggleAlign('right'),
          m: mergeSelection
        }[event.key.toLowerCase()];

        if (action) {
          event.preventDefault();
          action();
        }
      }
    },
    [mergeSelection, toggleAlign, toggleStyle]
  );

  useEffect(() => {
    const handleKeyDown = (event: KeyboardEvent) => {
      // Let text fields (e.g. the cell editor) handle their own undo.
      if (!(event.ctrlKey || event.metaKey) || event.altKey || isTextField(event.target)) {
        return;
      }

      const key = event.key.toLowerCase();

      if (key === 'z' && !event.shiftKey) {
        event.preventDefault();
        undoOrRedo('undo');
      } else if (key === 'y' || (key === 'z' && event.shiftKey)) {
        event.preventDefault();
        undoOrRedo('redo');
      }
    };

    document.addEventListener('keydown', handleKeyDown);

    return () => document.removeEventListener('keydown', handleKeyDown);
  }, [undoOrRedo]);

  const handleFileChange = useCallback(
    async (event: ChangeEvent<HTMLInputElement>) => {
      const fileInput = event.currentTarget;
      const file = fileInput.files?.[0];

      fileInput.value = '';

      if (!file) {
        return;
      }

      try {
        if (/\.zip$/i.test(file.name) || file.type === 'application/zip') {
          if (!confirm('Opening a ZIP file replaces all sheets in the current workbook. Continue?')) {
            return;
          }

          replaceWorkbook(await Workbook.fromZip(await file.arrayBuffer()));
        } else {
          if (workbook.worksheets.length >= MAX_WORKSHEETS) {
            alert(`A workbook cannot have more than ${MAX_WORKSHEETS} sheets.`);

            return;
          }

          const name = workbook.getUniqueWorksheetName(file.name.replace(/\.csv$/i, ''));
          const text = await file.text();
          const worksheet = workbook.addWorksheet(name);

          worksheet.edit(edit => edit.loadCSV(text));
          activateWorksheet(worksheet);
        }
      } catch (error) {
        alert(`Failed to open "${file.name}": ${(error as Error).message}`);
      }
    },
    [activateWorksheet, replaceWorkbook, workbook]
  );

  const handleAddSheet = useCallback(() => {
    if (workbook.worksheets.length < MAX_WORKSHEETS) {
      activateWorksheet(workbook.addWorksheet());
    }
  }, [activateWorksheet, workbook]);

  const handleDeleteSheet = useCallback(
    (worksheet: Worksheet) => {
      if (workbook.worksheets.length <= 1 || !confirm(`Delete sheet "${workbook.getWorksheetName(worksheet)}"?`)) {
        return;
      }

      const index = workbook.indexOf(worksheet);

      workbook.removeWorksheet(worksheet);

      if (activeWorksheetRef.current === worksheet) {
        activateWorksheet(workbook.worksheets[Math.min(index, workbook.worksheets.length - 1)]!);
      }
    },
    [activateWorksheet, workbook]
  );

  const handleRenameSheet = useCallback(
    (worksheet: Worksheet) => {
      const name = prompt('Rename sheet', workbook.getWorksheetName(worksheet));

      if (name === null) {
        return;
      }

      try {
        workbook.renameWorksheet(worksheet, name.trim());
      } catch (error) {
        alert((error as Error).message);
      }
    },
    [workbook]
  );

  // Re-computed on every render, `tabsRevision` is bumped when worksheets are added, removed or renamed.
  void tabsRevision;

  const tabs = workbook.worksheets.map(worksheet => ({ name: workbook.getWorksheetName(worksheet)!, worksheet }));

  const toolbarState = useMemo<ToolbarState>(() => {
    // `revision` is bumped when the worksheet or its selection changes.
    void revision;

    const worksheet = element.worksheet;
    const { left, top, right, bottom } = worksheet.selection.range;
    const { column, row } = worksheet.selection.activeCell;
    const topLeft = `${getColumnName(left)}${top + 1}`;
    const isApplied = (predicate: (formatting: CellFormatting | undefined) => boolean) =>
      isFormattingAppliedToSelection(worksheet, predicate);

    return {
      align: ALIGNS.find(align => isApplied(formatting => formatting?.align === align)),
      canRedo: worksheet.canRedo,
      canUndo: worksheet.canUndo,
      format: FORMATS.find(format => isApplied(formatting => formatting?.format === format)),
      highlight: HIGHLIGHTS.find(({ value }) =>
        isApplied(
          formatting =>
            !!formatting?.highlight && stringifyHighlight(formatting.highlight).toLowerCase() === value.toLowerCase()
        )
      )?.value,
      locked,
      merged: isSelectionMerged(worksheet),
      nameBox: left === right && top === bottom ? topLeft : `${topLeft}:${getColumnName(right)}${bottom + 1}`,
      status,
      styles: Object.fromEntries(
        STYLES.map(style => [style, isApplied(formatting => !!formatting?.[style])])
      ) as Record<CellStyle, boolean>,
      valueBox: formatValue(worksheet.getValue(column, row))
    };
  }, [element, locked, revision, status]);

  return (
    <ThemeProvider colorMode="day">
      <BaseStyles className="app">
        <input
          accept=".csv,.zip,text/csv,application/zip"
          hidden
          onChange={handleFileChange}
          ref={fileInputRef}
          type="file"
        />
        <PageLayout className="app__layout" columnGap="none" containerWidth="full" padding="none" rowGap="none">
          <PageLayout.Header className="app__header" divider="line" padding="condensed">
            <Toolbar actions={toolbarActions} returnFocusRef={returnFocusRef} state={toolbarState} />
          </PageLayout.Header>
          <PageLayout.Content className="app__content" padding="none" width="full">
            <div className="grid-container" onKeyDown={handleGridKeyDown} ref={gridContainerRef} />
          </PageLayout.Content>
          <PageLayout.Footer className="app__footer" divider="line" padding="none">
            <SheetBar
              activeWorksheet={activeWorksheet}
              canAdd={tabs.length < MAX_WORKSHEETS}
              onActivate={activateWorksheet}
              onAdd={handleAddSheet}
              onDelete={handleDeleteSheet}
              onRename={handleRenameSheet}
              tabs={tabs}
            />
          </PageLayout.Footer>
        </PageLayout>
      </BaseStyles>
    </ThemeProvider>
  );
}

export { App };
