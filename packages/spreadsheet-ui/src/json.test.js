import assert from 'node:assert/strict';
import { test } from 'node:test';
import { Worksheet } from './Worksheet.ts';

function createWorksheet() {
  const worksheet = new Worksheet({ columnCount: 10, rowCount: 20 });

  worksheet.edit(edit => {
    edit.setValues([
      ['a', 1, true],
      ['b', 2.5, false]
    ]);
    edit.setFormattings([[{ bold: true, align: 'center', highlight: { background: '#c6efce', color: '' } }]]);
    edit.setFormattings([[{ format: 'percent' }]], 1, 1);
    edit.setBorder({ x: 0, y: 0, width: 2, height: 2 }, new Set(['outside']));
    edit.mergeCells(4, 4, 2, 3);
  });
  worksheet.setColumnWidth(2, 123);

  return worksheet;
}

test('toJSON and fromJSON round trip', () => {
  const worksheet = createWorksheet();
  const json = JSON.parse(JSON.stringify(worksheet));
  const restored = Worksheet.fromJSON(json);

  assert.equal(restored.columnCount, 10);
  assert.equal(restored.rowCount, 20);
  assert.equal(restored.getColumnWidth(2), 123);
  assert.deepEqual(restored.getValues(0, 0, 10, 20), worksheet.getValues(0, 0, 10, 20));
  assert.deepEqual(restored.getFormattings(0, 0, 10, 20), worksheet.getFormattings(0, 0, 10, 20));
  assert.deepEqual(restored.getMerges(), [{ x: 4, y: 4, width: 2, height: 3 }]);

  for (let y = 0; y < 20; y++) {
    for (let x = 0; x < 10; x++) {
      assert.deepEqual(restored.getBorder(x, y), worksheet.getBorder(x, y));
    }
  }

  assert.deepEqual(restored.toJSON(), worksheet.toJSON());
  assert.equal(restored.canUndo, false);
});

test('loadJSON replaces existing content', () => {
  const worksheet = new Worksheet({ columnCount: 10, rowCount: 20 });

  worksheet.edit(edit => {
    edit.setValues([['x', 'y', 'z']], 0, 10);
    edit.setBorder({ x: 7, y: 7 }, new Set(['outside']));
    edit.mergeCells(8, 8, 2);
  });

  let load;

  worksheet.addEventListener('change', event => (load = event.load));
  worksheet.edit(edit => edit.loadJSON(createWorksheet().toJSON()));

  assert.equal(load, true);
  assert.deepEqual(worksheet.toJSON(), createWorksheet().toJSON());
});

test('invalid JSON throws and rolls back', () => {
  const worksheet = createWorksheet();
  const before = worksheet.toJSON();

  assert.throws(() => Worksheet.fromJSON({ version: 2 }), TypeError);
  assert.throws(
    () => worksheet.edit(edit => edit.loadJSON({ ...before, cells: [{ x: 0, y: 0, value: Infinity }] })),
    /Invalid worksheet JSON at "cells.0.value"/
  );
  assert.throws(
    () => worksheet.edit(edit => edit.loadJSON({ ...before, cells: [{ x: 99, y: 0, value: 'a' }] })),
    RangeError
  );
  assert.deepEqual(worksheet.toJSON(), before);
});
