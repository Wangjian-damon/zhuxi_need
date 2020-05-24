let XLSX = require('xlsx');
let fs = require('fs');
let path = require('path');
let outputDataPath = "./real.json";
/* 自己定义的excel文件 */
let xlsList = ['income.xls'];
let worksheetList = [] ;
let worksheets = [] ; //所有的工作表 
//得到所有文件
xlsList.forEach(e=>{
    let workbook = XLSX.readFile(e);
    let worksheet = workbook.Sheets[workbook.SheetNames[0]] ;
    worksheets.push(worksheet) ;
    worksheetList.push(worksheet);
});
let outputbook = XLSX.readFile('output.xls');
let outputWorksheet = outputbook.Sheets[outputbook.SheetNames[0]];
// 处理output
let outputData = [];
for(let i in outputWorksheet){
    let cell = outputWorksheet[i];
    // 列
    let col = i.charAt(0);
    // 行
    let row = parseInt(i.slice(1));
    if (!/!/.test(i)) {
        if (row >= 3) {
            let item = {
                key: '',
                count: '',
                price: '',
                row: row
            }
            if(outputData.length < row){
                outputData.push(item);
            }
            switch (col) {
                case 'H':
                    outputData[row-3].key = cell.v;
                    break;
                case 'I':
                    outputData[row-3].key += '~' + cell.v;
                    break;
                case 'J':
                    outputData[row-3].key += '~' + cell.v;
                    break;
                case 'K':
                    outputData[row-3].count = cell.v;
                    break;
                case 'L':
                    outputData[row-3].price = cell.v;
                    break;
            }
        }
    }
}
// 去除小计 去除无数据行
outputData = outputData.filter(e => (e.key && e.key.indexOf('小计') == -1));
noRpeatKey(outputData);
console.log(outputData.length);

function noRpeatKey(data){
    
}
// 处理进项
let incomeData = [];
worksheetList.forEach(worksheet=>{
    for(let i in worksheet){
        let cell = worksheet[i];
        // 列
        let col = i.charAt(0);
        // 行
        let row = parseInt(i.slice(1)); 
        if (!/!/.test(i)) {
            if (row >= 3) {
                let item = {
                    key: '',
                    count: '',
                    price: '',
                    row: row
                }
                if(incomeData.length < row){
                    incomeData.push(item);
                }
                // 以行的维度储存数据
                switch (col) {
                    case 'G':
                        incomeData[row-3].key = cell.v;
                        break;
                    case 'H':
                        incomeData[row-3].key += '~' + cell.v;
                        break;
                    case 'I':
                        incomeData[row-3].key += '~' + cell.v;
                        break;
                    case 'J':
                        incomeData[row-3].count = cell.v;
                        break;
                    case 'K':
                        incomeData[row-3].price = cell.v;
                        break;
                }
            }
        }
        
    }
})
// G行 开票项目默认为必有项
// 剔除包含小计的项
incomeData = incomeData.filter(e => (e.key && e.key.indexOf('小计') == -1));
// console.log(incomeData);