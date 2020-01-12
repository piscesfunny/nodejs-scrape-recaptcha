const xlsx = require('xlsx');
let excel = require('excel4node');

const getDataFromExcel = (sourceExcelPath) => {
    const workbook = xlsx.readFile(sourceExcelPath);
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];

    const columnC = [];

    for (let z in worksheet) {
        if(z.toString()[0] === 'C'){
            columnC.push(worksheet[z].v);
        }
    }
    return columnC.slice(1)
}

const getTempDataFromExcel = (sourceExcelPath) => {
    const workbook = xlsx.readFile(sourceExcelPath);
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];

    const columnA = [];
    const columnB = [];

    for (let z in worksheet) {
        if(z.toString()[0] === 'A'){
            columnA.push(worksheet[z].v);
        }

        if(z.toString()[0] === 'B'){
            columnB.push(worksheet[z].v);
        }
    }

    const tempInsuredNumbers = columnA.slice(1)
    const tempUninsuredNumbers = columnB.slice(1)

    return { tempInsuredNumbers, tempUninsuredNumbers }
}

const writeUniqueNumbersToExcel = (uniqueNumbers, resultExcelPath) => {
    const workbook = new excel.Workbook();

    const worksheet = workbook.addWorksheet('Sheet 1');

    uniqueNumbers.map((value, index) => {
        worksheet.cell(index+1, 1).string(value);
    })

    workbook.write(resultExcelPath);
}

const writeResultToExcel = (insuredNumbers, unInsuredNumbers, resultExcelPath) => {
    let workbook = new excel.Workbook();

    let worksheet = workbook.addWorksheet('Sheet 1');

    let styleForTitle = workbook.createStyle({
        font: {
            color: '#000000',
            size: 14
        },
    });

    let styleForInsured = workbook.createStyle({
        font: {
            color: '#08FF00',
            size: 12
        },
    });

    let styleForUninsured = workbook.createStyle({
        font: {
            color: '#FF0800',
            size: 12
        },
    });


    worksheet.cell(1,1).string('InsuredNumbers').style(styleForTitle);
    worksheet.cell(1,2).string('UninsuredNumbers').style(styleForTitle);
    insuredNumbers.map((value, index) => {
        worksheet.cell(index+2, 1).string(value).style(styleForInsured);
    })

    unInsuredNumbers.map((value, index) => {
        worksheet.cell(index+2, 2).string(value).style(styleForUninsured);
    })

    workbook.write(resultExcelPath);

    return true
}

const getUniqueArray = (array) => {
    let uniqueArray = [];

    // Loop through array values
    for(let value of array){
        if(uniqueArray.indexOf(value) === -1){
            uniqueArray.push(value);
        }
    }
    return uniqueArray;
}

module.exports = {
    getDataFromExcel,
    getTempDataFromExcel,
    writeUniqueNumbersToExcel,
    writeResultToExcel,
    getUniqueArray
}
