const XlsxPopulate = require("xlsx-populate");
const fs = require("fs");
const path = require("path");

// Putanja do Excel datoteke
const filePaths = [
  "C:/Users/Smeska i Smesko/Downloads/output.xlsx",
  "C:/Users/Smeska i Smesko/Downloads/output2.xlsx",
  "C:/Users/Smeska i Smesko/Downloads/output3.xlsx",
];

// Funkcija za konverziju broja kolone u ime kolone
function columnNumberToName(columnNumber) {
  let columnName = "";
  while (columnNumber > 0) {
    const remainder = (columnNumber - 1) % 26;
    columnName = String.fromCharCode(65 + remainder) + columnName;
    columnNumber = Math.floor((columnNumber - 1) / 26);
  }
  return columnName;
}

// Funkcija za izvlačenje imena iz fajlova i generisanje novih fajlova
async function extractNamesAndCreateFiles(filePaths) {
  const firstFilePath = filePaths[0];
  const folderPath = path.dirname(firstFilePath);

  // Putanja do direktorijuma Downloads
  const downloadsPath = "C:/Users/Smeska i Smesko/Downloads";

  // Kreiranje putanje do direktorijuma Refundacije
  const newFolderPath = path.join(downloadsPath, "Refundacije");

  // Kreiranje novog foldera
  fs.mkdirSync(newFolderPath, { recursive: true });

  // Kopiranje rasporeda kolona iz prvog fajla
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

      const newFilePath = path.join(newFolderPath, `${name}.xlsx`);

      const newWorkbook = await XlsxPopulate.fromBlankAsync();
      const newSheet = newWorkbook.sheet(0);

      // Kopiranje rasporeda kolona iz prvog fajla u novi fajl
      for (let k = 1; k <= columnCount; k++) {
        const columnLetter = columnNumberToName(k);
        const firstCell = firstSheet.cell(`${columnLetter}1`);
        const value = firstCell.value();
        newSheet.cell(`${columnLetter}1`).value(value);
      }
      await newWorkbook.toFileAsync(newFilePath);
    }
  }
}
// Izvlačenje imena iz fajlova i generisanje novih fajlova
extractNamesAndCreateFiles(filePaths)
  .then(() => {
    console.log("Generisanje fajlova je završeno.");
  })
  .catch((error) => {
    console.error("Došlo je do greške prilikom generisanja fajlova:", error);
  });