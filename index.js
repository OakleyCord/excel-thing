const Excel = require("exceljs");
require("dotenv").config();

const WRITE_SHEET = process.env.WRITE_SHEET;
const WRITE_COMPARE = process.env.WRITE_COMPARE;
const WRITE_OFFSET = process.env.WRITE_OFFSET;

const READ_SHEET = process.env.READ_SHEET;
const READ_COMPARE = process.env.READ_COMPARE;
const READ_OFFSET = process.env.READ_OFFSET;

const HEADERS = process.env.HEADERS.split(",");

async function main() {
    const startTime = Date.now();

    console.log("Loading Data...");
    const workbook = new Excel.Workbook();
    await workbook.xlsx.readFile("input.xlsx");

    const writeSheet = workbook.getWorksheet(WRITE_SHEET);
    const readSheet = workbook.getWorksheet(READ_SHEET);

    console.log("Comparing data...");
    compare(writeSheet, readSheet);

    console.log("\nSaving Data...");
    await workbook.xlsx.writeFile("output.xlsx");
    const time = (Date.now() - startTime) / 1000;
    console.log(`\nDone in ${time} seconds`);
}

function compare(writeSheet, readSheet) {
    const writeCol = writeSheet.getColumn(WRITE_COMPARE);
    const readData = readSheet.getColumn(READ_COMPARE);

    let readValues = [...readData.values];
    const writeValues = writeCol.values;
    const writeLength = writeValues.length;

    writeCol.eachCell((cell, writeRow) => {
        if (writeRow <= WRITE_OFFSET) return;

        const id = cell.value;

        for (
            let readRow = READ_OFFSET;
            readRow < readValues.length;
            readRow++
        ) {
            const readCell = readValues[readRow];

            if (id === readCell) {
                copyRows(writeSheet, readSheet, writeRow, readRow);
                readValues.splice(readRow, 1);
                break;
            }
        }

        if (writeRow % 100 === 0 || writeRow + 1 === writeLength)
            printProgress(writeRow + 1, writeLength);
    });
}

async function printProgress(index, length) {
    process.stdout.clearLine();
    process.stdout.cursorTo(0);
    const percentLeft = Math.floor((index / length) * 1000) / 10;
    process.stdout.write(`${index} / ${length} (${percentLeft}%)`);
}

const readLetters = [];
const writeLetters = [];
function copyRows(toSheet, fromSheet, toRow, fromRow) {
    for (let i = 0; i < HEADERS.length; i++) {
        const header = HEADERS[i];

        if (readLetters[i] == undefined) {
            const readColumn = fromSheet.getColumn(
                indexOfHeader(fromSheet, header, READ_OFFSET),
            );
            readLetters[i] = readColumn.letter;
        }

        if (writeLetters[i] == undefined) {
            const writeColumn = toSheet.getColumn(
                indexOfHeader(toSheet, header, WRITE_OFFSET),
            );
            writeLetters[i] = writeColumn.letter;
        }

        const value = fromSheet.getCell(`${readLetters[i]}${fromRow}`).value;

        toSheet.getCell(`${writeLetters[i]}${toRow}`).value = value;
    }
}

const headerValues = [];
function indexOfHeader(sheet, header, offset) {
    headerValues[sheet.id] = headerValues[sheet.id] || sheet.getRow(offset).values;
    return headerValues[sheet.id].indexOf(header);
}

main()
