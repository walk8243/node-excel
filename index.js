const XLSX  = require('xlsx'),
      Utils = XLSX.utils;

const book  = XLSX.readFile('book.xlsx');
// console.log(book.Sheets);
for(let sheetName in book.Sheets) {
  // console.log(sheetName);
  let sheet = book.Sheets[sheetName],
      range       = sheet['!ref'],
      decodeRange = Utils.decode_range(range);
  // console.log(sheet);
  // console.log(decodeRange);
  let headerText,
      colsArr     = [],
      valuesArr   = [],
      rowIndex    = decodeRange.s.r;

  for(let colIndex = decodeRange.s.c; colIndex <= decodeRange.e.c; colIndex++) {
    const address = Utils.encode_cell({ r: rowIndex, c:colIndex }),
          cell    = sheet[address];
    // console.log(address);
    // console.log(cell.v);
    colsArr.push(cell.v);
  }
  // console.log(colsArr);
  headerText = `INSERT INTO ${sheetName} (${colsArr.join(', ')}) VALUES`;
  // console.log(headerText);

  for(++rowIndex; rowIndex <= decodeRange.e.r; rowIndex++) {
    // console.log("rowIndex: " + rowIndex);
    // let lineText = "|";
    let linesArr  = [];
    for(let colIndex = decodeRange.s.c; colIndex <= decodeRange.e.c; colIndex++) {
      const address = Utils.encode_cell({ r: rowIndex, c:colIndex }),
            cell    = sheet[address],
            value   = String(cell.v);
      // console.log(address);
      // console.log(cell.v);
      // lineText += `  ${cell.v}  |`;
      if(value.match(/^\d+$/)) {
        linesArr.push(cell.v);
      }else{
        linesArr.push('"' + cell.v + '"');
      }
    }
    // console.log(linesArr);
    valuesArr.push("(" + linesArr.join(', ') + ")");
  }

  console.log(`${headerText}\n${valuesArr.join(',\n')}`);
}
