const xlsx = require("node-xlsx");
const fs = require("fs");
const path = require("path");
const figlet = require("figlet");
const chalk = require("chalk");
const ora = require("ora");

const {
  OPTION_KEYS,
  OUTPUT_TYPE_OPTIONS,
  ACTION_LIST,
  TOOL_PREFIEX,
} = require("./lib/constants");

class Spliter {
  constructor() {
    this.option = {};
    this.spinner = void 0;
  }

  run(option) {
    try {
      this.initSpinner();

      this.checkIsOptionsValid(option);

      this.setOption(option);

      const sheetList = this.parseXlsx();

      const originSheetData = this.getOriginSheetData(sheetList);

      const dataFormatted = this.formatData(originSheetData);

      const resultsFilted = this.filterData(dataFormatted);

      this.generateFiles(resultsFilted);
    } catch (error) {
      console.log(error);
    }
  }

  checkIsOptionsValid(option) {
    this.spinner.text = "配置格式校验中...";
    const msgSpinnerFail = "配置格式校验失败";

    if (!option) {
      return this.errHanlders(msgSpinnerFail, "缺少配置信息");
    }

    if (Object.prototype.toString.call(option) !== "[object Object]") {
      return this.errHanlders(msgSpinnerFail, "配置信息格式有误");
    }

    const { input, output } = option;
    if (!input) {
      return this.errHanlders(msgSpinnerFail, `缺少[${OPTION_KEYS.INPUT}]配置`);
    }

    if (!output) {
      return this.errHanlders(
        msgSpinnerFail,
        `缺少[${OPTION_KEYS.OUTPUT}]配置`
      );
    }

    if (!Object.keys(input).includes(OPTION_KEYS.PATH)) {
      return this.errHanlders(
        msgSpinnerFail,
        `缺少[${OPTION_KEYS.INPUT}.${OPTION_KEYS.PATH}]配置`
      );
    }

    if (!Object.keys(output).includes(OPTION_KEYS.PATH)) {
      return this.errHanlders(
        msgSpinnerFail,
        `缺少[${OPTION_KEYS.OUTPUT}.${OPTION_KEYS.PATH}]配置`
      );
    }

    if (!Object.keys(output).includes(OPTION_KEYS.ACTIONS)) {
      return this.errHanlders(
        msgSpinnerFail,
        `缺少[${OPTION_KEYS.OUTPUT}.${OPTION_KEYS.ACTIONS}]配置`
      );
    }

    return this.spinner.succeed("配置格式校验成功");
  }

  setOption(option) {
    this.option = option;
  }

  initSpinner() {
    this.spinner = ora("脚本启动");
    this.spinner.color = "green";
    this.spinner.start();
  }

  spinnerError(msg) {
    this.spinner.stop();
    this.spinner.fail(msg);
  }

  errHanlders(msgSpinner, msgChalk) {
    this.spinnerError(msgSpinner);
    console.log(`${chalk.red(`${TOOL_PREFIEX}${msgChalk}`)} `);
    throw Error(msgChalk);
  }

  parseXlsx() {
    this.spinner.text = "Excel读取中...";

    const {
      input: { path },
    } = this.option;

    let sheetList = [];
    try {
      sheetList = xlsx.parse(path);
    } catch (error) {
      return this.errHanlders(
        "Excel读取失败",
        "解析excel文件失败，请检查文件路径以及文件名后缀"
      );
    }
    this.spinner.succeed("Excel读取完毕");
    return sheetList;
  }

  getOriginSheetData(sheetList) {
    this.spinner.text = "获取原始Excel的Sheet数据...";

    let sheet = "";
    if (Object.keys(this.option.input).includes(OPTION_KEYS.SHEET)) {
      sheet = this.option.input.sheet;
    } else {
      sheet = "Sheet1";
    }
    const index = sheetList.findIndex(({ name }) => sheet === name);
    if (index === -1) {
      return this.errHanlders(
        "获取Sheet数据失败",
        "没有在Excel中找到对应的sheet，请检查[option.file.sheet]是否合法"
      );
    }

    this.spinner.succeed("获取Sheet数据成功");
    return sheetList[index].data || [];
  }

  formatData(sheetData) {
    this.spinner.text = "格式化Sheet数据...";
    const errMsg = "格式化Sheet数据失败";

    const dataFormatted = {
      rows: [],
      data: [],
    };

    try {
      sheetData.forEach((row, index) => {
        if (index === 0) {
          dataFormatted.rows = row;
        } else {
          const obj = {};
          dataFormatted.rows.forEach((key, index) => {
            obj[key] = row[index];
          });
          dataFormatted.data.push(obj);
        }
      });
    } catch (error) {
      return this.errHanlders(errMsg, errMsg);
    }

    this.spinner.succeed("格式化Sheet数据成功");
    return dataFormatted;
  }

  filterData(dataFromatted) {
    this.spinner.text = "根据查询条件拆解表格中...";

    const {
      output: { actions = {} },
    } = this.option;

    const results = [];
    for (const actionType of Object.keys(actions)) {
      results.push(
        this.generateChildExcel(actionType, actions[actionType], dataFromatted)
      );
    }

    this.spinner.succeed("表格数据拆解成功");
    return results;
  }

  generateChildExcel(key, action, originFormData) {
    const msgSpinnerFail = "表格数据拆解失败";
    let formData = JSON.parse(JSON.stringify(originFormData));
    let formList = formData.data || [];
    if (!Object.keys(action).includes(ACTION_LIST.search)) {
      return this.errHanlders(
        msgSpinnerFail,
        `配置缺少[${OPTION_KEYS.OUTPUT}.${OPTION_KEYS.ACTIONS}.[fileName].${OPTION_KEYS.SEARCH}]字段`
      );
    }

    // 根据查询配置过滤数据
    const searchActions = action[ACTION_LIST.search];
    if (!Array.isArray(searchActions)) {
      return this.errHanlders(msgSpinnerFail, "[search]配置必须是Array类型】");
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
  }

  generateFiles(results) {
    this.spinner.text = "正在生成excel文件...";

    const { output } = this.option;
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
        xlsxBuildParams.push(this.getSheetDataFormatted(result));
      }
      this.writeFile("result", xlsxBuildParams);
    } else {
      for (const { fileName, formData, extraData } of results) {
        this.writeFile(fileName, [
          this.getSheetDataFormatted({ fileName, formData, extraData }),
        ]);
      }
    }
  }

  getSheetDataFormatted({ fileName: name, formData, extraData }) {
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
  }

  writeFile(fileName, xlsxBuildParams) {
    try {
      fs.writeFileSync(
        path.resolve(this.option.output.path, `./${fileName}.xlsx`),
        xlsx.build(xlsxBuildParams)
      );

      this.spinner.stop();
      this.spinner.succeed("写入文件成功");
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
      return console.log(error);
    }
  }
}

module.exports = new Spliter();
