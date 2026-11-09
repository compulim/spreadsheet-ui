import '@primer/primitives/dist/css/primitives.css';
import '@primer/primitives/dist/css/functional/themes/light.css';
import { createRoot } from 'react-dom/client';
import { App } from './ui/App.tsx';

createRoot(document.getElementById('root') as HTMLElement).render(<App />);
