const fs = require("fs-extra");
const xlsx = require("xlsx");

async function mapSheetFromInputFile(mapperFunction, inputFilePath) {
  const inputBuffer = await fs.readFile(inputFilePath);
  return mapSheetFromBuffer(mapperFunction, inputBuffer);
}

async function mapSheetFromInputFileToOutputFile(
  mapperFunction,
  inputFilePath,
  outputFilePath
) {
  const outputBuffer = await mapSheetFromInputFile(
    mapperFunction,
    inputFilePath
  );

  await fs.writeFile(outputFilePath, outputBuffer);
  return outputFilePath;
}

async function mapSheetFromBuffer(mapperFunction, inputBuffer) {
  const workbook = xlsx.read(inputBuffer, { type: "buffer" });
  const sheet = getFirstSheet(workbook);
  const headers = getHeaders(sheet);

  const destinationColumn = headers.length;

  const headerNameForInsertedColumn = "Output";
  if (headerNameForInsertedColumn) {
    const destinationHeaderCellAddress = xlsx.utils.encode_cell({ r: 0, c: destinationColumn });
    sheet[destinationHeaderCellAddress] = { v: headerNameForInsertedColumn };
  }

  const decodedRange = xlsx.utils.decode_range(sheet["!ref"]);
  const newEndColumn = decodedRange.e.c + 1;
  sheet["!ref"] = xlsx.utils.encode_range({
    s: decodedRange.s,
    e: { c: newEndColumn, r: decodedRange.e.r }
  });

  const startRow = 1; // Assuming first row is a header row
  const endRow = decodedRange.e.r;

  for (let rowIndex = startRow; rowIndex <= endRow; rowIndex++) {
    const inputRowObject = {};
    const inputRowArray = [];
    for (let columnIndex = 0; columnIndex < newEndColumn; columnIndex++) {
      const cellValue = getCellValue(sheet, columnIndex, rowIndex);
      inputRowObject[headers[columnIndex]] = cellValue;
      inputRowArray.push(cellValue);
    }

    const output = await mapperFunction.apply(sheet, [
      inputRowObject,
      inputRowArray
    ]);
    const destnationCellAddress = xlsx.utils.encode_cell({
      c: destinationColumn,
      r: rowIndex
    });
    sheet[destnationCellAddress] = { v: output };
  }

  const outputBuffer = xlsx.write(workbook, { type: "buffer" });
  return outputBuffer;
}

function getCellValue(sheet, columnIndex, rowIndex) {
  const cellAddress = xlsx.utils.encode_cell({ c: columnIndex, r: rowIndex });
  const cell = sheet[cellAddress];
  if (cell) {
    return cell.w;
  }
  return null;
}

function getHeaders(sheet) {
  const ref = sheet["!ref"];
  if (!ref) {
    return [];
  }
  const range = xlsx.utils.decode_range(ref);

  const headers = [];
  for (let c = range.s.c; c <= range.e.c; c++) {
    const cellAddress = xlsx.utils.encode_cell({ r: 0, c });
    const cell = sheet[cellAddress];
    let value;
    if (cell) {
      value = cell.v;
    }
    headers.push(value);
  }
  return headers;
}

function getFirstSheet(workbook) {
  const firstSheetName = workbook.SheetNames[0];
  return workbook.Sheets[firstSheetName];
}

module.exports = {
  mapSheetFromBuffer,
  mapSheetFromInputFile,
  mapSheetFromInputFileToOutputFile
};
