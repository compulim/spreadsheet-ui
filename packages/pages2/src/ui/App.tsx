import { WorksheetApp } from '@compulim/spreadsheet-ui';
import { memo, useEffect, useRef } from 'react';

const App = memo(() => {
  const worksheetAppRef = useRef<WorksheetApp>(null);

  useEffect(() => {
    const { current: worksheetApp } = worksheetAppRef;

    worksheetApp?.worksheet.edit(transaction => {
      transaction.setValue(1, 1, 'Hello, World!');
      transaction.setValue(1, 2, 'Aloha!');
      transaction.setFormattings([[{ bold: true }]], 1, 1);
      transaction.setFormattings([[{ italic: true, highlight: { background: '#f76666', color: '#000' } }]], 1, 2);
    });

    console.log(worksheetApp?.worksheet.toJSON());
  }, [worksheetAppRef]);

  return (
    <div>
      <p>Hello, World!</p>
      <worksheet-app ref={worksheetAppRef} style={{ width: 500, height: 500 }}></worksheet-app>
    </div>
  );
});

App.displayName = 'App';

export default App;
