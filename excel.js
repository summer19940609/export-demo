const excel = require('exceljs');
const path = require('path');

const wb = new excel.Workbook();
const sheet = wb.addWorksheet('sheet1');

const ws = wb.getWorksheet('sheet1');

// mock数据
const mockData = [
    {
        "city": "-",
        "shopName": "美味一店",
        "shopId": 310
    },
    {
        "city": "北京",
        "shopName": "美味一店",
        "shopId": 311
    },
    {
        "city": "上海",
        "shopName": "美味一店",
        "shopId": 312,
    },
    {
        "city": "上海",
        "shopName": "美味二店",
        "shopId": 313,
    },
    {
        "city": "上海",
        "shopName": "美味三店",
        "shopId": 314,
    },
    {
        "city": "广州",
        "shopName": "美味四店",
        "shopId": 315,
    },
    {
        "city": "广州",
        "shopName": "美味五店",
        "shopId": 316,
    }
]

// 数据填充
ws.columns = [
    {
        header: '门店',
        key: 'city'
    },
    {
        header: '门店',
        key: 'shopName',
    },
    {
        header: '门店id',
        key: 'shopId',
    }
]

mockData.forEach(v => {
    ws.addRow(v);
})

// 计算出data里列合并情况
let merge_index = {}
mockData.forEach((v, i) => {
    if (!merge_index[v.city]) {
        merge_index[v.city] = {
            s: {
                '开始行': i + 1,
                '开始列': 1,
            }
        }
    } else {
        merge_index[v.city]['e'] = {
            '结束行': i + 1,
            '结束列': 1,
        }
    }
})

Object.keys(merge_index).forEach(v => {
    if (merge_index[v]['e']) {
        ws.mergeCells(
            merge_index[v]['s']['开始行'],
            merge_index[v]['s']['开始列'],
            merge_index[v]['e']['结束行'],
            merge_index[v]['e']['结束列'],
        )
    }
})


ws.eachRow(row => {
    row.eachCell(cell => {
        // 四周边框
        cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
        // 单元格垂直居中对齐
        cell.alignment = { vertical: 'middle', horizontal: 'center' }
    })
})

const fileName = '模拟.xlsx';
const savePath = path.join(__dirname, '/' + fileName);

wb.xlsx.writeFile(savePath)