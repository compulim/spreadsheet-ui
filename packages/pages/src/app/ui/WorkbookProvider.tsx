import { createContext, useContext, useMemo, useRef } from 'react';
import type { ReactNode } from 'react';
import type { Worksheet, WorksheetJSON } from '@compulim/spreadsheet-ui';

type WorkbookContextValue = {
  /** Returns the saved JSON of the worksheet, or the JSON of the worksheet itself if never saved. */
  getWorksheetJSON: (worksheet: Worksheet) => WorksheetJSON;
  /** Saves the JSON of the worksheet. */
  setWorksheetJSON: (worksheet: Worksheet, json: WorksheetJSON) => void;
  /** Writes all saved JSON back to their worksheets. */
  flush: () => void;
};

const WorkbookContext = createContext<WorkbookContextValue | undefined>(undefined);

/** Keeps the JSON of every worksheet, saved by worksheet views when they are unmounted. */
function WorkbookProvider({ children }: { readonly children?: ReactNode | undefined }) {
  const jsonsRef = useRef<Map<Worksheet, WorksheetJSON>>(new Map());

  const value = useMemo<WorkbookContextValue>(
    () => ({
      flush: () => {
        for (const [worksheet, json] of jsonsRef.current) {
          worksheet.edit(edit => edit.loadJSON(json));
        }

        jsonsRef.current.clear();
      },
      getWorksheetJSON: worksheet => jsonsRef.current.get(worksheet) ?? worksheet.toJSON(),
      setWorksheetJSON: (worksheet, json) => void jsonsRef.current.set(worksheet, json)
    }),
    []
  );

  return <WorkbookContext.Provider value={value}>{children}</WorkbookContext.Provider>;
}

function useWorkbookContext(): WorkbookContextValue {
  const context = useContext(WorkbookContext);

  if (!context) {
    throw new Error('useWorkbookContext() must be used inside <WorkbookProvider>.');
  }

  return context;
}

export { useWorkbookContext, WorkbookProvider, type WorkbookContextValue };
