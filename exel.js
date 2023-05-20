const XlsxPopulate = require("xlsx-populate");
const fs = require("fs");
const path = require("path");

const filePaths = [
  "/home/milos/Downloads/output.xlsx",
  "/home/milos/Downloads/output2.xlsx",
  "/home/milos/Downloads/output3.xlsx",
];

function columnNumberToName(columnNumber) {
  let columnName = "";
  while (columnNumber > 0) {
    const remainder = (columnNumber - 1) % 26;
    columnName = String.fromCharCode(65 + remainder) + columnName;
    columnNumber = Math.floor((columnNumber - 1) / 26);
  }
  return columnName;
}

async function extractNamesAndCreateFiles(filePaths) {
  const generatedFileNames = new Set();
  const nameOccurrences = {};

  const firstFilePath = filePaths[0];
  const folderPath = path.dirname(firstFilePath);
  const downloadsPath = "/home/milos/Downloads";
  const newFolderPath = path.join(downloadsPath, "Refundacije");

  fs.mkdirSync(newFolderPath, { recursive: true });

  const firstWorkbook = await XlsxPopulate.fromFileAsync(firstFilePath);
  const firstSheet = firstWorkbook.sheet(0);
  const columnCount = firstSheet.usedRange().endCell().columnNumber();

  for (let i = 0; i < filePaths.length; i++) {
    const filePath = filePaths[i];
    const workbook = await XlsxPopulate.fromFileAsync(filePath);
    const sheet = workbook.sheet(0);

    const rowCount = sheet.usedRange().endCell().rowNumber();
    for (let j = 2; j <= rowCount; j++) {
      const cell = sheet.cell(`A${j}`);
      const name = cell.value();

      generatedFileNames.add(name);

      const newFilePath = path.join(newFolderPath, `${name}.xlsx`);

      if (!nameOccurrences[name]) {
        nameOccurrences[name] = [];
      }
      const rowValues = [];
      for (let k = 1; k <= columnCount; k++) {
        const columnLetter = columnNumberToName(k);
        const cellValue = sheet.cell(`${columnLetter}${j}`).value();
        rowValues.push(cellValue);
      }
      nameOccurrences[name].push(rowValues);

      if (fs.existsSync(newFilePath)) {
        const existingWorkbook = await XlsxPopulate.fromFileAsync(newFilePath);
        const existingSheet = existingWorkbook.sheet(0);

        const existingRowCount = existingSheet.usedRange().endCell().rowNumber();
        const newRow = existingRowCount + 1;

        for (let k = 1; k <= columnCount; k++) {
          const columnLetter = columnNumberToName(k);
          const value = sheet.cell(`${columnLetter}${j}`).value();
          existingSheet.cell(`${columnLetter}${newRow}`).value(value);
        }

        await existingWorkbook.toFileAsync(newFilePath);
      } else {
        const newWorkbook = await XlsxPopulate.fromBlankAsync();
        const newSheet = newWorkbook.sheet(0);

        for (let k = 1; k <= columnCount; k++) {
          const columnLetter = columnNumberToName(k);
          const firstCell = firstSheet.cell(`${columnLetter}1`);
          const value = firstCell.value();
          newSheet.cell(`${columnLetter}1`).value(value);

          const cellValue = sheet.cell(`${columnLetter}${j}`).value();
          newSheet.cell(`${columnLetter}2`).value(cellValue);
        }

        await newWorkbook.toFileAsync(newFilePath);
      }
    }
  }

  return { generatedFileNames, nameOccurrences };
}

extractNamesAndCreateFiles(filePaths)
  .then(({ generatedFileNames, nameOccurrences }) => {
    generatedFileNames.forEach((name) => {
      const occurrences = nameOccurrences[name];
      console.log(`Podaci za ime: ${name}`);
      occurrences.forEach((rowValues, index) => {
        console.log(`Red ${index + 1}: ${rowValues}`);
      });
    });
    console.log("Generisanje fajlova je završeno.");
  })
  .catch((error) => {
    console.error("Došlo je do greške prilikom generisanja fajlova:", error);
  });
