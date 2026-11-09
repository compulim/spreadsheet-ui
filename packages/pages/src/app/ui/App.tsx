import { BaseStyles, ThemeProvider } from '@primer/react';
import { WorkbookProvider } from './WorkbookProvider.tsx';
import { WorkbookView } from './WorkbookView.tsx';

function App() {
  return (
    <ThemeProvider colorMode="day">
      <BaseStyles className="app">
        <WorkbookProvider>
          <WorkbookView />
        </WorkbookProvider>
      </BaseStyles>
    </ThemeProvider>
  );
}

export { App };
