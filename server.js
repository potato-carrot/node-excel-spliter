const xlsx = require('node-xlsx')
const protocol = require('./protocol.json')
const fs = require('fs');
const path = require('path');

const ACTION_LIST = {
    search: "search",
    count: "count",
    sum: "sum",
}

const generateExcelFiles = ({ fileName, formData, extraData }) => {
    const excelData = {
        name: 'Sheet1',
        data: [formData.rows]
    }
    formData.data.forEach((item) => {
        let arr = new Array(formData.rows.length).fill(void 0)
        formData.rows.forEach((col, colIndex) => {
            arr[colIndex] = item[col]
        })
        excelData.data.push(arr)
    })

    let arr = new Array(formData.rows.length).fill(void 0)
    extraData.forEach(({ col, data }) => {
        const index = formData.rows.findIndex(rowName => rowName === col)
        if (index > -1) {
            arr[index] = data
        }

    })
    excelData.data.push(arr)

    const buffer = xlsx.build([excelData]);
    fs.writeFile(path.resolve(__dirname, `./forms/${fileName}.xlsx`), buffer, function (err) {
        if (err) {
            console.log(err);
            return;
        }
        console.log(`~~~~~~生成${fileName}.xlsx成功~~~~~~`);
    });
}

const generateChildExcel = (key, action, originFormData) => {
    let formData = JSON.parse(JSON.stringify(originFormData))
    let formList = formData.data || []
    if (!Object.keys(action).includes(ACTION_LIST.search)) {
        throw Error('protocol.json里某个表缺少【search】字段，请检查！')
    }

    // 根据查询配置过滤数据
    const searchActions = action[ACTION_LIST.search]
    if (!Array.isArray(searchActions)) {
        throw Error('【search】字段必须是【数组类型】，请检查！')
    }
    searchActions.forEach(({ col, value }) => {
        formList = formList.filter(item => item[col] === value)
    })
    formData.data = formList;

    const extraData = []

    // 计算数据条数
    if (Object.keys(action).includes(ACTION_LIST.count)) {
        const countAction = action[ACTION_LIST.count]
        extraData.push({ col: countAction.col, data: formList.length })
    }

    // 计算某些列的值的和
    if (Object.keys(action).includes(ACTION_LIST.sum)) {
        const sumActions = action[ACTION_LIST.sum]
        sumActions.forEach(({ col }) => {
            let total = 0
            formData.data.forEach(item => {
                let value = 0
                if (item[col] && typeof item[col] === 'number') {
                    value = item[col]
                }
                total += value
            })
            extraData.push({ col, data: total })
        })
    }
    return { fileName: key, formData, extraData }
}

const parseXlsx = async ({ fatherForm, childrenForms }) => {
    let sheetList = xlsx.parse(`./${fatherForm}.xlsx`);
    console.log(111111111, sheetList);

    //对数据进行处理
    for (const sheet of sheetList) {
        const formData = formatData(sheet.data)
        for (const key of Object.keys(childrenForms)) {
            const formDataFilted = generateChildExcel(key, childrenForms[key], formData)
            await generateExcelFiles(formDataFilted)
        }
    }
}

const formatData = (sheetData) => {
    const formData = {
        rows: [],
        data: []
    }
    sheetData.forEach((row, index) => {
        if (index === 0) {
            formData.rows = row
        } else {
            const obj = {}
            formData.rows.forEach((key, index) => {
                obj[key] = row[index]
            })
            formData.data.push(obj)
        }
    })
    return formData
}

parseXlsx(protocol)

