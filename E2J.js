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
})
let TotalData = [];
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
                    price: ''
                }
                switch (col) {
                    case 'G':
                        item.key = cell.v;
                        break;
                    case 'H':
                        item.key += '_' + cell.v;
                        break;
                    case 'I':
                        item.key += '_' + cell.v;
                        break;
                    case 'J':
                        item.count = cell.v;
                        break;
                }
                TotalData.push(item);
            }
        }
        
    }
})
console.log(TotalData);