import { type WorksheetApp } from '@compulim/spreadsheet-ui';
import type * as React from 'react';

declare global {
  const IS_DEVELOPMENT: boolean;

  namespace React {
    namespace JSX {
      interface IntrinsicElements {
        'worksheet-app': React.DetailedHTMLProps<React.HTMLAttributes<WorksheetApp>, WorksheetApp> & {
          style?: React.CSSProperties;
        };
      }
    }
  }
}

export { };
