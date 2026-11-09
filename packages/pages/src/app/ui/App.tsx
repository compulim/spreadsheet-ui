import { BaseStyles, ThemeProvider } from '@primer/react';
import { WorkbookView } from './WorkbookView.tsx';

function App() {
  return (
    <ThemeProvider colorMode="day">
      <BaseStyles className="app">
        <WorkbookView />
      </BaseStyles>
    </ThemeProvider>
  );
}

export { App };
