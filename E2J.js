let XLSX = require('xlsx');
let fs = require('fs');
let path = require('path');
let outputDataPath = "./real.json";
/* 自己定义的excel文件 */
let xlsList = ['income.xls', 'output.xls', 'in_out.xls'];
let worksheetList = [] ;
let worksheets = [] ; //所有的工作表 
//得到所有文件
xlsList.forEach(e=>{
    let workbook = XLSX.readFile(e);
    let worksheet = workbook.Sheets[workbook.SheetNames[0]] ;
    worksheets.push(worksheet) ;
    worksheetList.push(worksheet);
});

// 处理进项
let totalData = [[], [], []];
console.log(worksheetList.length);
worksheetList.forEach((worksheet, idx)=>{
    let type = ['income', 'output', 'in_out'][idx];
    for(let i in worksheet){
        let cell = worksheet[i];
        // 列
        let col = i.charAt(0);
        // 行
        let row = parseInt(i.slice(1)); 

        if (!/!/.test(i)) {
            if (idx < 2) {
                if (row >= 3) {
                    let item = {
                        key: '',
                        count: '',
                        price: '',
                        row: row,
                        type: type
                    }
                    if(totalData[idx].length < row){
                        totalData[idx].push(item);
                    }
                    // 以行的维度储存数据
                    if (idx == 0) {
                        switch (col) {
                            case 'G':
                                break;
                            case 'H':
                                totalData[idx][row-3].key += '_' + cell.v;
                                break;
                            case 'I':
                                totalData[idx][row-3].key += '_' + cell.v;
                                break;
                            case 'J':
                                totalData[idx][row-3].count = parseFloat(cell.v);
                                break;
                            case 'K':
                                totalData[idx].price = parseFloat(parseFloat(cell.v).toFixed(2));
                                break;
                        }
                    }
                    if (idx == 1) {
                        switch (col) {
                            case 'H':
                                totalData[idx][row-3].key = cell.v;
                                break;
                            case 'I':
                                totalData[idx][row-3].key += '_' + cell.v;
                                break;
                            case 'J':
                                totalData[idx][row-3].key += '_' + cell.v;
                                break;
                            case 'K':
                                totalData[idx][row-3].count = parseFloat(cell.v);
                                break;
                            case 'L':
                                totalData[idx][row-3].price = parseFloat(parseFloat(cell.v).toFixed(2));
                                break;
                        }
                    }
                }
                
            }
            if (idx == 2) {
                if (row >= 6) {
                    let item = {
                        key: '',
                        count: '',
                        price: '',
                        row: row,
                        type: type
                    }
                    if(totalData[idx].length < row){
                        totalData[idx].push(item);
                    }
                    switch (col) {
                        case 'A':
                            totalData[idx][row - 6].key = cell.v;
                            break;
                        case 'B':
                            totalData[idx][row-6].key += '_' + cell.v;
                            break;
                        case 'C':
                            totalData[idx][row-6].key += '_' + cell.v;
                            break;
                        case 'M':
                            totalData[idx][row-6].count = parseFloat(cell.v);
                            break;
                        case 'O':
                            totalData[idx][row-6].price = parseFloat(parseFloat(cell.v).toFixed(2));
                            break;
                    }
                }
            }
        }
        
    }
})
// G行 开票项目默认为必有项
// 剔除包含小计的项
totalData.forEach(item => {
    item = item.filter(e => (e.key && e.key.indexOf('小计') == -1 && e.key.indexOf('__') == -1));
    noRpeatKey(item);
    console.log(item.length);
})
console.log(totalData.length);
function noRpeatKey(data){
    let arr = [];
    data.forEach(e => {
        if(!arr.find(item => item.key == e.key)){
            arr.push(e);
        }else{
            // console.log('repeat', e);
            // 
            arr[arr.findIndex(item => item.key == e.key)].count += e.count;
            arr[arr.findIndex(item => item.key == e.key)].price += e.price;
        }
    })
}
// 拿到 进项销项做处理
function exportExcel(fileName, data) {
    let wb = XLSX.utils.book_new();
    for (let i in data) {
        let tmp = XLSX.utils.aoa_to_sheet(data[i]);
        XLSX.utils.book_append_sheet(wb, tmp, i);
    }
    XLSX.writeFile(wb, fileName + '.xlsx');
}