import '@compulim/spreadsheet-ui';
import '@compulim/spreadsheet-ui/index.css';
import { createRoot } from 'react-dom/client';
import App from './ui/App.tsx';

const mainElement = document.createElement('main');

document.body.append(mainElement);

createRoot(mainElement).render(<App />);

typeof IS_DEVELOPMENT !== 'undefined' &&
  IS_DEVELOPMENT &&
  new EventSource('/esbuild').addEventListener('change', () => location.reload());
