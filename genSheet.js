const Excel = require("exceljs");
const { v4: uuidv4 } = require('uuid');

const SIZE = 10000;

const names = [
    "Joe",
    "Sally",
    "Bob",
    "Sue",
    "Jim",
    "Fred",
    "Sarah",
]


const workbook = new Excel.Workbook();






const sheet1 = workbook.addWorksheet("Sheet1");
sheet1.getCell("A1").value = "ID";
sheet1.getCell("B1").value = "Name";
sheet1.getCell("C1").value = "Salary";
for(let i = 2; i < SIZE; i++) {
    sheet1.getCell(`A${i}`).value = uuidv4();
}

const sheet2 = workbook.addWorksheet("Sheet2");
sheet2.getCell("A1").value = "ID";
sheet2.getCell("B1").value = "Name";
sheet2.getCell("C1").value = "Salary";
for(let i = 2; i < SIZE; i++) {
    sheet2.getCell(`A${i}`).value = uuidv4();
    sheet2.getCell(`B${i}`).value = names[Math.floor(Math.random() * names.length)];
    sheet2.getCell(`C${i}`).value = Math.floor(Math.random() * 100000);
}



for(let i = 2; i < SIZE; i++) {
    const randomRow = Math.floor(Math.random() * SIZE);
    const randomRow2 = Math.floor(Math.random() * SIZE);

    sheet1.getCell(`A${randomRow2}`).value = sheet2.getCell(`A${randomRow}`).value;
}



workbook.xlsx.writeFile("input.xlsx");