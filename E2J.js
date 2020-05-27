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
// console.log(worksheetList.length);
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
                        detail: {
                            key: '',
                            count: '',
                            price: '',
                        },
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
                                totalData[idx][row-3].detail.key = cell.v;
                                break;
                            case 'H':
                                totalData[idx][row-3].detail.key += '_' + cell.v;
                                break;
                            case 'I':
                                totalData[idx][row-3].detail.key += '_' + cell.v;
                                break;
                            case 'J':
                                totalData[idx][row-3].detail.count = parseFloat(cell.v);
                                break;
                            case 'K':
                                totalData[idx][row-3].detail.price = parseFloat(parseFloat(cell.v).toFixed(2));
                                break;
                        }
                    }
                    if (idx == 1) {
                        switch (col) {
                            case 'H':
                                totalData[idx][row-3].detail.key = cell.v;
                                break;
                            case 'I':
                                totalData[idx][row-3].detail.key += '_' + cell.v;
                                break;
                            case 'J':
                                totalData[idx][row-3].detail.key += '_' + cell.v;
                                break;
                            case 'K':
                                totalData[idx][row-3].detail.count = parseFloat(cell.v);
                                break;
                            case 'L':
                                totalData[idx][row-3].detail.price = parseFloat(parseFloat(cell.v).toFixed(2));
                                break;
                        }
                    }
                }
                
            }
            if (idx == 2) {
                if (row >= 6) {
                    let item = {
                        detail: {
                            key: '',
                            count: '',
                            price: '',
                        },
                        row: row,
                        type: type
                    }
                    if(totalData[idx].length < row){
                        totalData[idx].push(item);
                    }
                    switch (col) {
                        case 'A':
                            totalData[idx][row - 6].detail.key = cell.v;
                            break;
                        case 'B':
                            totalData[idx][row-6].detail.key += '_' + cell.v;
                            break;
                        case 'C':
                            totalData[idx][row-6].detail.key += '_' + cell.v;
                            break;
                        case 'M':
                            totalData[idx][row-6].detail.count = parseFloat(cell.v);
                            break;
                        case 'O':
                            totalData[idx][row-6].detail.price = parseFloat(parseFloat(cell.v).toFixed(2));
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
    item = item.filter(e => (e.detail.key && e.detail.key.indexOf('小计') == -1 && e.detail.key.indexOf('__') == -1));
    item.forEach(e => {
        if (e.key == '') {
            console.log(e.row);
        }
    })
    noRpeatKey(item);
    // console.log(item.length);
})
// console.log(totalData.length);
function noRpeatKey(data){
    let arr = [];
    data.forEach(e => {
        if(!arr.find(item => item.detail.key == e.detail.key)){
            arr.push(e);
        }else{
            // console.log('repeat', e);
            // 重复项 则对应数量 金额 相加
            arr[arr.findIndex(item => item.detail.key == e.detail.key)].detail.count += e.detail.count;
            arr[arr.findIndex(item => item.detail.key == e.detail.key)].detail.price += e.detail.price;
        }
    })
}
// 拿到totalData 进项销项期初做处理
// 拿到最大项
let maxLength = Math.max(...totalData.map(e => e.length));
// console.log(maxLength);

let mainData = totalData.find(e => e.length == maxLength);
totalData.splice(totalData.findIndex(e => e.length == maxLength), 1);
// console.log(totalData.length);
// console.log(mainData);
let data = mainData.map((item) => {
    let id = item.detail.key;
    // console.log(id);
    let data1 = totalData[0].find(e => e.detail.key == id) || { detail: {} };
    let data2 = totalData[1].find(e => e.detail.key == id) || { detail: {} };
    // console.log(data2);
    let key1 = data1.type || totalData[0][0].type;
    let key2 = data2.type || totalData[1][0].type;
    // 需要的数据结构
    let dataItem = {
        key: id,
        // 名称
        name: id.split('_')[0],
        // 规格
        spe: id.split('_')[1] || '',
        // 单位
        unit: id.split('_')[2] || '',
        [item.type + '_count']: item.detail.count,
        [item.type + '_price']: item.detail.price,
        [key1 + '_count']: data1.detail.count || 0,
        [key2 + '_count']: data2.detail.count || 0,
        [key1 + '_price']: data1.detail.price || 0,
        [key2 + '_price']: data1.detail.price || 0,
    }
    return dataItem;
});
// 计算单价
data = data.map(e => {
    e.income_price_unit = calNum(e.income_price, e.income_count) || 0;
    e.output_price_unit = calNum(e.output_price, e.output_count) || 0;
    e.in_out_price_unit = calNum(e.in_out_price, e.in_out_count) || 0;
    return e;
});
function calNum(num1, num2) {
    return isFinite(num1 / num2) ? num1 / num2 : 0;
}
console.log(data[0]);
let keys = [];
for (let props in data[0]) {
    if (props != 'key') {
        keys.push(props)
    }
}
console.log(keys);
let valuesMap = {
    'name': '开票项目',
    'spe': '规格型号' ,
    'unit': '计量单位',
    'in_out_count':'期初数量',
    'in_out_price':'期初金额',
    'income_count': '入库数量',
    'output_count': '出库数量',
    'income_price': '入库金额',
    'output_price': '出库金额',
    'income_price_unit': '入库单价',
    'output_price_unit': '出库单价',
    'in_out_price_unit': '期初单价'
}
let values = keys.map(e => valuesMap[e]);
console.log(values);
// {
//     sheet1: [chooseFieldValues, ..._getVal(res.result_class_total, choosenKeys)],
// }
function _getVal(data, choosenKeys) { 
    return data.map(e =>
        choosenKeys.map(j => {
            return e[j];
        })
      );
}
exportExcel('result', {sheet1: [values, ..._getVal(data, keys)]})
function exportExcel(fileName, data) {
    let wb = XLSX.utils.book_new();
    for (let i in data) {
        let tmp = XLSX.utils.aoa_to_sheet(data[i]);
        XLSX.utils.book_append_sheet(wb, tmp, i);
    }
    XLSX.writeFile(wb, fileName + '.xlsx');
}