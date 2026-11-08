import type { CellValue } from '@compulim/spreadsheet-ui';

function formatValue(value: CellValue | null | undefined): string {
  if (typeof value === 'boolean') {
    return value ? 'TRUE' : 'FALSE';
  }

  return value === undefined || value === null ? '' : String(value);
}

export { formatValue };
