const xlsx = require('xlsx')
const fs = require('fs')

let mockData = [
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

// 表头配置
const headerFields = ['city', 'shopName', 'shopId']
const headerDisplay = {
    "city": '门店',
    "shopId": '店id'
}

let wb = xlsx.utils.book_new()

// 计算出data里列合并情况
let merge_index = {}
mockData.forEach((v, i) => {
    if (!merge_index[v.city]) {
        merge_index[v.city] = {
            s: { c: 0, r: i + 1 },
        }
    } else {
        merge_index[v.city]['e'] = { c: 0, r: i + 1 }
    }
})


mockData = [headerDisplay, ...mockData]
const ws = xlsx.utils.json_to_sheet(mockData, { header: headerFields, skipHeader: true })

// 单元格合并
ws['!merges'] = [
    { s: { c: 0, r: 0 }, e: { c: 1, r: 0 } },
];

Object.keys(merge_index).forEach(v => {
    if (merge_index[v]['e']) {
        ws['!merges'].push(merge_index[v])
    }
})
console.log(ws['!merges'])

xlsx.utils.book_append_sheet(wb, ws, '会员数据')

xlsx.writeFile(wb, 'out.xlsx');