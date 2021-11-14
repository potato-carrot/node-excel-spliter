const xlsx = require("node-xlsx");
const option = require("./option.json");
const fs = require("fs");
const path = require("path");
const figlet = require("figlet");
const chalk = require("chalk");
const ora = require("ora");

const { errHanlders } = require("./lib/utils");
const {
  OPTION_KEYS,
  OUTPUT_TYPE_OPTIONS,
  ACTION_LIST,
  TOOL_PREFIEX,
} = require("./lib/constants");

let spinner = void 0;

const generateFiles = (results) => {
  spinner.text = "正在写入文件...";

  const { output } = option;
  let outputType = OUTPUT_TYPE_OPTIONS.SHEETS;

  if (
    Object.keys(output).includes(OPTION_KEYS.OUTPUT_TYPE) &&
    OUTPUT_TYPE_OPTIONS.includes(output.outputType)
  ) {
    outputType = output.outputType;
  }

  if (outputType === OPTION_KEYS.SHEETS) {
    const xlsxBuildParams = [];
    for (const result of results) {
      xlsxBuildParams.push(getSheetData(result));
    }
    writeFile("excel拆解结果", xlsxBuildParams);
  } else {
    for (const { fileName, formData, extraData } of results) {
      writeFile(fileName, [getSheetData({ fileName, formData, extraData })]);
    }
  }
};

const writeFile = (fileName, xlsxBuildParams) => {
  try {
    fs.writeFileSync(
      path.resolve(__dirname, `./${OPTION_KEYS.OUTPUT_DIR}/${fileName}.xlsx`),
      xlsx.build(xlsxBuildParams)
    );
    spinner.stop();
    spinner.succeed("写入文件成功");
    console.log(`${chalk.green(`${TOOL_PREFIEX}生成${fileName}.xlsx成功`)} `);
    console.log(
      "\r\n" +
        figlet.textSync("excel-spliter", {
          horizontalLayout: "default",
          verticalLayout: "default",
          width: 100,
          whitespaceBreak: true,
        })
    );
  } catch (error) {
    return console.log(err);
  }
};

const getSheetData = ({ fileName: name, formData, extraData }) => {
  const sheetData = {
    name,
    data: [formData.rows],
  };
  formData.data.forEach((item) => {
    let arr = new Array(formData.rows.length).fill(void 0);
    formData.rows.forEach((col, colIndex) => {
      arr[colIndex] = item[col];
    });
    sheetData.data.push(arr);
  });

  let arr = new Array(formData.rows.length).fill(void 0);
  extraData.forEach(({ col, data }) => {
    const index = formData.rows.findIndex((rowName) => rowName === col);
    if (index > -1) {
      arr[index] = data;
    }
  });
  sheetData.data.push(arr);
  return sheetData;
};

const generateChildExcel = (key, action, originFormData) => {
  let formData = JSON.parse(JSON.stringify(originFormData));
  let formList = formData.data || [];
  if (!Object.keys(action).includes(ACTION_LIST.search)) {
    errHanlders("请检查options.json里某个表缺少【search】字段，请检查！");
  }

  // 根据查询配置过滤数据
  const searchActions = action[ACTION_LIST.search];
  if (!Array.isArray(searchActions)) {
    errHanlders("【search】字段必须是【数组类型】，请检查！");
  }
  searchActions.forEach(({ col, value }) => {
    formList = formList.filter((item) => item[col] === value);
  });
  formData.data = formList;

  const extraData = [];

  // 计算数据条数
  if (Object.keys(action).includes(ACTION_LIST.count)) {
    const countAction = action[ACTION_LIST.count];
    extraData.push({ col: countAction.col, data: formList.length });
  }

  // 计算某些列的值的和
  if (Object.keys(action).includes(ACTION_LIST.sum)) {
    const sumActions = action[ACTION_LIST.sum];
    sumActions.forEach(({ col }) => {
      let total = 0;
      formData.data.forEach((item) => {
        let value = 0;
        if (item[col] && typeof item[col] === "number") {
          value = item[col];
        }
        total += value;
      });
      extraData.push({ col, data: total });
    });
  }
  return { fileName: key, formData, extraData };
};

const parseXlsx = async (option) => {
  spinner = ora("Excel读取中...");
  spinner.color = "green";
  spinner.start();

  checkIsOptionsValid(option);
  const {
    input: { name },
    output,
  } = option;

  let sheetList = null;
  try {
    sheetList = xlsx.parse(`./${OPTION_KEYS.INPUT_DIR}/${name}.xlsx`);
  } catch (error) {
    return console.log(
      `${chalk.red(
        `${TOOL_PREFIEX}${"解析excel文件失败，请检查文件路径以及文件名后缀"}`
      )} `
    );
  }

  spinner.succeed("Excel读取完毕");
  spinner.text = "Excel拆解中...";

  // 默认把拆表结果放到一个excel的多个sheet中
  const { actions = {} } = output;
  const formData = formatData(getOriginExcelData(sheetList));

  const results = [];
  spinner.succeed("Excel拆解完毕");
  spinner.text = "正在生成子表...";
  for (const actionType of Object.keys(actions)) {
    const formDataFilted = generateChildExcel(
      actionType,
      actions[actionType],
      formData
    );
    results.push(formDataFilted);
  }

  generateFiles(results);
};

const checkIsOptionsValid = (option) => {
  const { input, output } = option;
  if (!input) {
    errHanlders(`[option]缺少[${OPTION_KEYS.INPUT}]配置`);
  }

  if (!output) {
    errHanlders(`[option]缺少[${OPTION_KEYS.OUTPUT}]配置`);
  }

  if (!Object.keys(input).includes(OPTION_KEYS.NAME)) {
    errHanlders(`[option.input]缺少[${OPTION_KEYS.NAME}]配置`);
  }
};

const formatData = (sheetData) => {
  const formData = {
    rows: [],
    data: [],
  };
  sheetData.forEach((row, index) => {
    if (index === 0) {
      formData.rows = row;
    } else {
      const obj = {};
      formData.rows.forEach((key, index) => {
        obj[key] = row[index];
      });
      formData.data.push(obj);
    }
  });
  return formData;
};

const getOriginExcelData = (sheetList) => {
  let sheet = "";
  if (Object.keys(option.input).includes(OPTION_KEYS.SHEET)) {
    sheet = option.input.sheet;
  } else {
    sheet = "Sheet1";
  }
  const index = sheetList.findIndex(({ name }) => sheet === name);
  if (index === -1) {
    errHanlders(
      "没有在Excel中找到对应的sheet，请检查[option.file.sheet]是否合法"
    );
  }
  return sheetList[index].data || [];
};

parseXlsx(option);
