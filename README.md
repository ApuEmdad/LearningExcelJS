

## ExcelJS Basic Usage

### Installing ExcelJS

```bash
npm install exceljs
```

### Creating a Workbook and Sheet

```js
const ExcelJS = require('exceljs');

const workbook = new ExcelJS.Workbook();
const sheet = workbook.addWorksheet('Sheet1');
```

---

## Columns

### Defining Columns

```js
sheet.columns = [
  { header: 'ID', key: 'id', width: 10 },
  { header: 'Name', key: 'name', width: 25 },
  { header: 'Score', key: 'score', width: 10 }
];
```

### Accessing Columns

```js
const col1 = sheet.getColumn(1);       // by index
const colScore = sheet.getColumn('score'); // by key
```

### Column Properties and Methods

```js
colScore.header       // get or set the header text
colScore.width        // get or set width
colScore.hidden       // hide/show column
colScore.values       // get/set all values in the column as an array
colScore.eachCell((cell, rowNumber) => {
  console.log(rowNumber, cell.value);
});
```

---

## Rows

### Adding Rows

**Object-style rows**

```js
sheet.addRow({ id: 1, name: 'Apu', score: 90 });
sheet.addRow({ id: 2, name: 'Haque', score: 85 });
```

**Array-style rows**

```js
sheet.addRow([3, 'Test User', 92]);
```

### Accessing Rows

```js
const row1 = sheet.getRow(1); // first row
```

### Row Properties and Methods

```js
row1.values         // get/set entire row as array
row1.getCell(2)     // get cell by column number
row1.getCell('name') // get cell by key (object-style)
row1.eachCell((cell, colNumber) => {
  console.log(colNumber, cell.value);
});
row1.height         // set row height
row1.hidden         // hide/show row
row1.number         // row index
```

---

## Cells

### Accessing Cells

```js
const cellA1 = sheet.getCell('A1');  // by address
const cellR1C1 = sheet.getCell(1,1); // row 1, column 1
```

### Setting and Getting Values

```js
cellA1.value = 'Hello';
console.log(cellA1.value); // read value

cellA2.value = { formula: 'B2+C2' }; // formula
```

### Cell Properties and Methods

```js
cell.value       // get/set value
cell.text        // formatted text
cell.address     // e.g., "A1"
cell.row         // row number
cell.col         // column number
```

### Iterating Cells

```js
sheet.eachRow((row) => {
  row.eachCell((cell) => {
    console.log(cell.address, cell.value);
  });
});
```

---


## Styling Cells

### Font Styling

```js
const cell = sheet.getCell('B1');
cell.font = {
  name: 'Arial',
  size: 12,
  bold: true,
  italic: false,
  color: { argb: 'FF0000' } // red text
};
```

### Alignment

```js
cell.alignment = {
  horizontal: 'center', // left, right
  vertical: 'middle',   // top, bottom
  wrapText: true
};
```

### Borders

```js
cell.border = {
  top: { style: 'thin' },
  left: { style: 'thin' },
  bottom: { style: 'thin' },
  right: { style: 'thin' }
};
```

### Fill / Background Color

```js
cell.fill = {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: 'FFFF00' } // yellow background
};
```

---

## Formulas

### Setting Formulas

```js
sheet.getCell('C2').value = { formula: 'A2+B2' };
```

### Reading Calculated Values

```js
const cell = sheet.getCell('C2');
console.log(cell.value); // shows the formula object
console.log(cell.result); // shows the last calculated result (if workbook was calculated)
```

### Example with Multiple Rows

```js
sheet.addRow([1, 2]); // A3, B3
sheet.getCell('C3').value = { formula: 'A3+B3' };
```

---

## Merging Cells

### Merge and Unmerge

```js
sheet.mergeCells('A1:B1'); // merge A1 and B1
sheet.unMergeCells('A1:B1');
```

### Accessing Merged Cells

```js
const cell = sheet.getCell('A1'); // top-left of merged range
cell.value = 'Merged Cell';
```

---

## Data Validation

### List Validation

```js
sheet.getCell('D2').dataValidation = {
  type: 'list',
  allowBlank: true,
  formulae: ['"Option1,Option2,Option3"'],
  showInputMessage: true,
  promptTitle: 'Choose an option',
  prompt: 'Select one from the list'
};
```

### Whole Number Validation

```js
sheet.getCell('E2').dataValidation = {
  type: 'whole',
  operator: 'between',
  showErrorMessage: true,
  formulae: [1, 100],
  error: 'Value must be between 1 and 100'
};
```

### Date Validation

```js
sheet.getCell('F2').dataValidation = {
  type: 'date',
  operator: 'greaterThan',
  formulae: [new Date(2025, 0, 1)],
  error: 'Date must be after Jan 1, 2025'
};
```

---

This completes the **full set of commonly used ExcelJS methods**: writing, reading, object/array style, styling, formulas, merging, and validation.

---

## ExcelJS Cheat Sheet

### Columns

| Method / Property        | Description               | Example                                                    |
| ------------------------ | ------------------------- | ---------------------------------------------------------- |
| `sheet.columns`          | Define columns            | `sheet.columns = [{ header: 'ID', key: 'id', width: 10 }]` |
| `sheet.getColumn(n)`     | Get column by index       | `const col = sheet.getColumn(1)`                           |
| `sheet.getColumn('key')` | Get column by key         | `const col = sheet.getColumn('id')`                        |
| `column.header`          | Get/set header text       | `col.header = 'New Header'`                                |
| `column.width`           | Get/set width             | `col.width = 15`                                           |
| `column.hidden`          | Hide/show column          | `col.hidden = true`                                        |
| `column.values`          | Get/set all column values | `console.log(col.values)`                                  |
| `column.eachCell(cb)`    | Iterate column cells      | `col.eachCell((cell, row) => console.log(cell.value))`     |

---

### Rows

| Method / Property     | Description               | Example                                                |
| --------------------- | ------------------------- | ------------------------------------------------------ |
| `sheet.addRow(obj)`   | Add object-style row      | `sheet.addRow({ id: 1, name: 'Apu' })`                 |
| `sheet.addRow(array)` | Add array-style row       | `sheet.addRow([1, 'Apu'])`                             |
| `sheet.getRow(n)`     | Access a row              | `const row = sheet.getRow(1)`                          |
| `row.values`          | Get/set row values        | `row.values = [1, 'Name']`                             |
| `row.getCell(n)`      | Get cell by column number | `row.getCell(2)`                                       |
| `row.getCell('key')`  | Get cell by key           | `row.getCell('name')`                                  |
| `row.eachCell(cb)`    | Iterate row cells         | `row.eachCell((cell, col) => console.log(cell.value))` |
| `row.height`          | Set row height            | `row.height = 20`                                      |
| `row.hidden`          | Hide/show row             | `row.hidden = true`                                    |
| `row.number`          | Row index                 | `console.log(row.number)`                              |

---

### Cells

| Method / Property         | Description                    | Example                                                                     |
| ------------------------- | ------------------------------ | --------------------------------------------------------------------------- |
| `sheet.getCell('A1')`     | Get cell by address            | `const cell = sheet.getCell('A1')`                                          |
| `sheet.getCell(row, col)` | Get cell by row/column numbers | `sheet.getCell(1, 1)`                                                       |
| `cell.value`              | Get/set value                  | `cell.value = 'Hello'`                                                      |
| `cell.text`               | Formatted text                 | `console.log(cell.text)`                                                    |
| `cell.address`            | Address string                 | `console.log(cell.address)`                                                 |
| `cell.row`                | Row number                     | `console.log(cell.row)`                                                     |
| `cell.col`                | Column number                  | `console.log(cell.col)`                                                     |
| `cell.font`               | Font styling                   | `cell.font = { bold: true }`                                                |
| `cell.alignment`          | Alignment                      | `cell.alignment = { horizontal: 'center' }`                                 |
| `cell.border`             | Borders                        | `cell.border = { top: { style: 'thin' } }`                                  |
| `cell.fill`               | Background fill                | `cell.fill = { type: 'pattern', pattern:'solid', fgColor:{argb:'FFFF00'} }` |
| `cell.note`               | Add comment                    | `cell.note = 'This is a note'`                                              |
| `cell.dataValidation`     | Validation rules               | `cell.dataValidation = { type:'list', formulae:['"Yes,No"'] }`              |

---

### Iteration

```js
// Iterate all rows and cells
sheet.eachRow((row) => {
  row.eachCell((cell) => {
    console.log(cell.address, cell.value);
  });
});

// Iterate a column
sheet.getColumn('score').eachCell((cell, rowNumber) => {
  console.log(rowNumber, cell.value);
});
```

---

### Formulas, Merging, Validation Examples

```js
// Formula
sheet.getCell('C2').value = { formula: 'A2+B2' };

// Merge
sheet.mergeCells('A1:B1');
sheet.getCell('A1').value = 'Merged';

// Data Validation
sheet.getCell('D2').dataValidation = {
  type: 'list',
  formulae: ['"Option1,Option2,Option3"']
};
```

---








****
