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
    let worksheet = workbook.Sheets[workbook.SheetNames[0]];
    worksheets.push(worksheet);
    worksheetList.push(worksheet);
});
// json 表读为json的方法
// console.log(XLSX.utils.sheet_to_json(worksheetList[0])[0]);
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
                                totalData[idx][row-3].detail.key += '=' + cell.v;
                                break;
                            case 'I':
                                totalData[idx][row-3].detail.key += '=' + cell.v;
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
                                totalData[idx][row-3].detail.key += '=' + cell.v;
                                break;
                            case 'J':
                                totalData[idx][row-3].detail.key += '=' + cell.v;
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
                            totalData[idx][row-6].detail.key += '=' + cell.v;
                            break;
                        case 'C':
                            totalData[idx][row-6].detail.key += '=' + cell.v;
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
totalData = totalData.map(item => {
    // && e.detail.key.indexOf('__') == -1
    item = item.filter(e => (e.detail.key && e.detail.key.indexOf('小计') == -1 && e.detail.key.indexOf('合计') == -1));
    console.dir(item);
    // item.forEach(e => {
    //     // console.log(e.detail.key);
    //     if (e.detail.key == '' || e.detail.key.indexOf('合计') != -1) {
    //         console.log(e.row);
    //     }
    // })
    noRpeatKey(item);
    return item;
    // console.log(item.length);
})
// return;
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
    // 数据结构决定表结构
    let dataItem = {
        key: id,
        // 名称
        name: id.split('=')[0].replace(/\*([^\*]+)\*/g,''),
        // 规格
        spe: id.split('=')[1] || '',
        // 单位
        unit: id.split('=')[2] || '',
        [item.type + '_count']: item.detail.count,
        [item.type + '_price_unit']: 0,
        [item.type + '_price']: item.detail.price,
        [key1 + '_count']: data1.detail.count || 0,
        [key1 + '_price_unit']: 0,
        [key1 + '_price']: data1.detail.price || 0,
        [key2 + '_count']: data2.detail.count || 0,
        [key2 + '_price_unit']: 0,
        [key2 + '_price']: data2.detail.price || 0,
        final_count: 0,
        final_price_unit: 0,
        final_price: 0
    }
    return dataItem;
});
// 计算单价
data = data.map(e => {
    // 入库单价
    e.income_price_unit = calNum(e.income_price + e.in_out_price, e.income_count + e.in_out_count ) || 0;
    // 出库单价
    e.output_price_unit = e.income_price_unit;
    // 按照公式重新计算出库金额
    // 当
    if (e.income_count != 0 || e.in_out_count != 0) {
        e.output_price = parseFloat((e.output_price_unit * e.output_count).toFixed(2));
        // if(e.output_price != 0){
        //     console.log(e.output_price);
        // }
    }
    //
    // 期初
    e.in_out_price_unit = calNum(e.in_out_price, e.in_out_count) || 0;
    // 8.计算‘销项’的‘单价’，若‘期初’和‘进项’的‘数量’都为0，则‘销项金额’=0.9*‘销项金额’，若‘期初’和‘进项’的‘数量’不全为0， 则‘销项单价’=‘进项单价’，‘销项金额’=‘销项数量’*‘销项单价’
    if (e.income_count == 0 && e.in_out_count == 0) {
        e.final_price *= 0.9;
        e.final_price_unit *= 0.9;
        // 如果期初数量和入库数量都为0时候，出库金额=原始表格出库金额*0.9，然后再由‘出库金额 / 出库数量 = 出库单价’ 来求得出库单价。
        e.output_price *= .9;
        e.output_price = parseFloat(e.output_price.toFixed(2));
        if(e.output_price != 0){
            console.log(e.output_price);
        }
        e.output_price_unit = calNum(e.output_price, e.output_count);
    }
    // 最后计算计算期末
    e.final_price = e.in_out_price + e.income_price - e.output_price;
    e.final_count = e.in_out_count + e.income_count - e.output_count;
    e.final_price_unit = calNum(e.final_price, e.final_count) || 0;
    
    return e;
});
function calNum(num1, num2) {
    return isFinite(num1 / num2) ? num1 / num2 : 0;
}
//
// console.log(data[0]);
let keys = [];
for (let props in data[0]) {
    if (props != 'key') {
        keys.push(props)
    }
}
let AZarr = []
for (let i = 65; i <= 90; i++) {
    AZarr.push(String.fromCharCode(i))
}
console.log(keys);
let colMap = {};
keys.map((e, i) => {
    colMap[e] = AZarr[i]
})
console.log(colMap);
// return;
let valuesMap = {
    'name': '开票项目',
    'spe': '规格型号' ,
    'unit': '计量单位',
    'in_out_count':'期初数量',
    'in_out_price':'期初金额',
    'in_out_price_unit': '期初单价',
    'income_count': '入库数量',
    'income_price': '入库金额',
    'income_price_unit': '入库单价',
    'output_count': '出库数量',
    'output_price': '出库金额',
    'output_price_unit': '出库单价',
    'final_count': '期末数量',
    'final_price': '期末金额',
    'final_price_unit': '期末单价',
}
let values = keys.map(e => valuesMap[e]);
console.log('values', values);
// let values = [{
//     A: '开票项目',
//     B: '规格型号',
//     C: '计量单位',
//     D: '期初数量',
//     E: '期初金额',
//     F: '入库数量',
//     G: '入库金额',
//     H: '入库单价',
//     I: '出库数量',
//     J: '出库金额',
//     K: '出库单价',
//     L: '出库数量',
//     M: '出库金额',
//     N: '出库单价',
// }]
// return;
function _getVal(data, choosenKeys) { 
    console.log(data);
    return data.map(e =>
        choosenKeys.map(j => {
            // if (j.indexOf('in_out') != -1) {
            //     return { B: e[j] };
            // } else if (j.indexOf('income') != -1) {
            //     return { C: e[j]}
            // } else if (j.indexOf('output') != -1) {
            //     return { D: e[j]}
            // } else {
            //     return {A: e[j]}
            // }
            // console.log({ [colMap[j]]: e[j] })
            // return { [colMap[j]]: e[j] }
            return e[j];
        })
      );
}
exportExcel('result', {sheet1: [values, ..._getVal(data, keys)]})
function exportExcel(fileName, data) {
    let wb = XLSX.readFile('schema.xlsx');
    ws = wb.Sheets[wb.SheetNames[0]];
    // console.log(ws);
    // 可以单独修改sheet内某一项的值
    let now = new Date();
    ws['A3'].v = '编制日期：' + now.getFullYear() + '年' + (now.getMonth() + 1) + '月' + now.getDate() + '日';
    // console.log(ws['A3']);
    // let wb = XLSX.utils.book_new();
    for (let i in data) {
        // console.log(data[i].slice(1790));
        // let tmp = XLSX.utils.aoa_to_sheet(data[i], { origin: 'A5' });
        XLSX.utils.sheet_add_aoa(ws, data[i], { origin: 'A5' })
        // XLSX.utils.book_append_sheet(wb, tmp, i);
    }
    XLSX.writeFile(wb, fileName + '.xlsx');
}