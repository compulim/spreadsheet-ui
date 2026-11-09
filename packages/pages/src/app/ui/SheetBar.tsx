import { useEffect, useRef } from 'react';
import type { Worksheet } from '@compulim/spreadsheet-ui';

type SheetBarTab = {
  name: string;
  worksheet: Worksheet;
};

function SheetBar({
  activeWorksheet,
  canAdd,
  onActivate,
  onAdd,
  onDelete,
  onRename,
  tabs
}: {
  activeWorksheet: Worksheet | undefined;
  canAdd: boolean;
  onActivate: (worksheet: Worksheet) => void;
  onAdd: () => void;
  onDelete: (worksheet: Worksheet) => void;
  onRename: (worksheet: Worksheet) => void;
  tabs: readonly SheetBarTab[];
}) {
  const tabListRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    tabListRef.current
      ?.querySelector('[aria-selected="true"]')
      ?.scrollIntoView({ block: 'nearest', inline: 'nearest' });
  }, [activeWorksheet, tabs.length]);

  return (
    <div className="sheet-bar">
      <button
        aria-label="Add sheet"
        className="sheet-bar__add"
        disabled={!canAdd}
        onClick={onAdd}
        title="Add sheet"
        type="button"
      >
        +
      </button>
      <div aria-label="Sheets" className="sheet-bar__tabs" ref={tabListRef} role="tablist">
        {tabs.map(({ name, worksheet }) => (
          <div
            aria-selected={worksheet === activeWorksheet}
            className="sheet-tab"
            key={name}
            onClick={() => onActivate(worksheet)}
            onDoubleClick={() => onRename(worksheet)}
            role="tab"
            title="Click to select, double-click to rename"
          >
            <span>{name}</span>
            {tabs.length > 1 && (
              <button
                aria-label={`Delete sheet ${name}`}
                className="sheet-tab__close"
                onClick={event => {
                  event.stopPropagation();
                  onDelete(worksheet);
                }}
                title={`Delete "${name}"`}
                type="button"
              >
                ×
              </button>
            )}
          </div>
        ))}
      </div>
    </div>
  );
}

export { SheetBar, type SheetBarTab };
