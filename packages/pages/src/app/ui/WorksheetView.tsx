import { WorksheetSelection } from '@compulim/spreadsheet-ui';
import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import type { KeyboardEvent as ReactKeyboardEvent } from 'react';
import { download } from '../download.ts';
import { formatValue } from '../formatValue.ts';
import { getColumnName } from '../getColumnName.ts';
import { parseHighlight, stringifyHighlight } from '../highlight.ts';
import { HIGHLIGHTS } from '../highlights.ts';
import { cellsToPNG } from '../printCells.ts';
import type { Workbook } from '../workbook/Workbook.ts';
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

type WorksheetViewProps = {
  /** Called when the user asks to open a file. */
  onOpen: () => void;
  /** Called when the user asks to save the workbook. */
  onSaveZip: () => void;
  /** The workbook containing `worksheet`. */
  workbook: Workbook;
  /** The worksheet to show and edit. */
  worksheet: Worksheet | undefined;
};

/** A worksheet, with its toolbar, editing a copy of `worksheet` and saving every change back to it. */
function WorksheetView({ onOpen, onSaveZip, workbook, worksheet: activeWorksheet }: WorksheetViewProps) {
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

  const [locked, setLocked] = useState(false);
  const [status, setStatus] = useState('');
  /** Bumped when the worksheet of `element`, or its selection, changes. */
  const [revision, setRevision] = useState(0);

  const activeWorksheetRef = useRef(activeWorksheet);
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
      open: onOpen,
      print: printSelection,
      redo: () => undoOrRedo('redo'),
      saveCSV,
      saveZip: onSaveZip,
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
      onOpen,
      onSaveZip,
      printSelection,
      saveCSV,
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
    <div className="worksheet">
      <div className="worksheet__header">
        <Toolbar actions={toolbarActions} returnFocusRef={returnFocusRef} state={toolbarState} />
      </div>
      <div className="grid-container" onKeyDown={handleGridKeyDown} ref={gridContainerRef} />
    </div>
  );
}

export { WorksheetView, type WorksheetViewProps };
