const Excel = require('exceljs');
let workbook = new Excel.Workbook();

const read = async () => {
    let result = new Array();
    await workbook.xlsx.readFile('./AttRep.xlsx')
        .then(function () {
            let worksheet = workbook.getWorksheet('First Sheet');
            let headArr = [];
            for (let i = 1; i < 12; i++)
                headArr.push(worksheet.getRow(1).getCell(i).value);
            headArr.push("date");
            headArr.push("Att Status");
            result.push(headArr);

            for (let i = 12; i < worksheet.columnCount; i++) { //worksheet.columnCount
                if (worksheet.getRow(2).getCell(i).value == "Total Present")
                    break;
                result.push(["", "", "", "", "", "", "", "", "", "", "","",""]);
                for (let r = 3; r < worksheet.rowCount; r++) {
                    let arr = [];
                    if (worksheet.getRow(r).getCell(1).value == null)
                        break;
                    for (let c = 1; c < 12; c++)
                        arr.push(worksheet.getRow(r).getCell(c).value);
                    arr.push(worksheet.getRow(2).getCell(i).value + " " + worksheet.getRow(1).getCell(i).value);
                    arr.push(worksheet.getRow(r).getCell(i).value);
                    result.push(arr);
                }
            }
    });

    workbook = new Excel.Workbook();
    worksheet = workbook.addWorksheet('Result Sheet');
        for (let i = 0; i < result.length; i++) {
            worksheet.addRow(result[i]);
    }

    workbook.xlsx.writeFile('output.xlsx').then(function () {
        console.log('File saved!');
    });
}

read();