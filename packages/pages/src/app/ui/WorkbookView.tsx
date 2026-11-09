import { PageLayout } from '@primer/react';
import { useCallback, useEffect, useRef, useState } from 'react';
import type { ChangeEvent } from 'react';
import { download } from '../download.ts';
import { MAX_WORKSHEETS } from '../workbook/constants.ts';
import { Workbook } from '../workbook/Workbook.ts';
import { SheetBar } from './SheetBar.tsx';
import { WorksheetView } from './WorksheetView.tsx';
import type { Worksheet } from '@compulim/spreadsheet-ui';

function createWorkbook() {
  const workbook = new Workbook();

  workbook.addWorksheet();

  return workbook;
}

/** A workbook, showing its active worksheet with toolbar, and a sheet bar to switch between worksheets. */
function WorkbookView() {
  const [workbook, setWorkbookState] = useState(createWorkbook);
  const [activeWorksheet, setActiveWorksheet] = useState<Worksheet | undefined>(() => workbook.worksheets[0]);
  /** Bumped when worksheets of the workbook are added, removed or renamed. */
  const [tabsRevision, setTabsRevision] = useState(0);

  const activeWorksheetRef = useRef(activeWorksheet);
  const fileInputRef = useRef<HTMLInputElement>(null);

  activeWorksheetRef.current = activeWorksheet;

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

  const replaceWorkbook = useCallback((nextWorkbook: Workbook) => {
    if (!nextWorkbook.worksheets.length) {
      nextWorkbook.addWorksheet();
    }

    setWorkbookState(nextWorkbook);
    setActiveWorksheet(nextWorkbook.worksheets[0]);
  }, []);

  const handleOpen = useCallback(() => fileInputRef.current?.click(), []);

  const handleSaveZip = useCallback(async () => {
    try {
      download(new Blob([(await workbook.toZip()) as BlobPart], { type: 'application/zip' }), 'workbook.zip');
    } catch (error) {
      alert(`Failed to save workbook: ${(error as Error).message}`);
    }
  }, [workbook]);

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
          setActiveWorksheet(worksheet);
        }
      } catch (error) {
        alert(`Failed to open "${file.name}": ${(error as Error).message}`);
      }
    },
    [replaceWorkbook, workbook]
  );

  const handleAddSheet = useCallback(() => {
    if (workbook.worksheets.length < MAX_WORKSHEETS) {
      setActiveWorksheet(workbook.addWorksheet());
    }
  }, [workbook]);

  const handleDeleteSheet = useCallback(
    (worksheet: Worksheet) => {
      if (workbook.worksheets.length <= 1 || !confirm(`Delete sheet "${workbook.getWorksheetName(worksheet)}"?`)) {
        return;
      }

      const index = workbook.indexOf(worksheet);

      workbook.removeWorksheet(worksheet);

      if (activeWorksheetRef.current === worksheet) {
        setActiveWorksheet(workbook.worksheets[Math.min(index, workbook.worksheets.length - 1)]!);
      }
    },
    [workbook]
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

  return (
    <>
      <input
        accept=".csv,.zip,text/csv,application/zip"
        hidden
        onChange={handleFileChange}
        ref={fileInputRef}
        type="file"
      />
      <PageLayout className="app__layout" columnGap="none" containerWidth="full" padding="none" rowGap="none">
        <PageLayout.Content className="app__content" padding="none" width="full">
          <WorksheetView
            onOpen={handleOpen}
            onSaveZip={handleSaveZip}
            workbook={workbook}
            worksheet={activeWorksheet}
          />
        </PageLayout.Content>
        <PageLayout.Footer className="app__footer" divider="line" padding="none">
          <SheetBar
            activeWorksheet={activeWorksheet}
            canAdd={tabs.length < MAX_WORKSHEETS}
            onActivate={setActiveWorksheet}
            onAdd={handleAddSheet}
            onDelete={handleDeleteSheet}
            onRename={handleRenameSheet}
            tabs={tabs}
          />
        </PageLayout.Footer>
      </PageLayout>
    </>
  );
}

export { WorkbookView };
