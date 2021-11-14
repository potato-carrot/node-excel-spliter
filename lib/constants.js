const OPTION_KEYS = {
  FILE: "file",
  CONFIGS: "configs",
  INPUT: "input",
  OUTPUT: "output",
  NAME: "name",
  PATH: "path",
  ACTIONS: "actions",
  SHEET: "sheet",
  OUTPUT_TYPE: "outputType",
  FILES: "files",
  SHEETS: "sheets",
  SEARCH: "search",
  COUNT: "count",
  SUM: "sum",
  INPUT_DIR: "file-input",
  OUTPUT_DIR: "file-output",
};

const OUTPUT_TYPE_OPTIONS = [OPTION_KEYS.FILES, OPTION_KEYS.SHEETS];

const ACTION_LIST = {
  search: OPTION_KEYS.SEARCH,
  count: OPTION_KEYS.COUNT,
  sum: OPTION_KEYS.SUM,
};

const TOOL_PREFIEX = "[spliter] ";

module.exports = {
  OPTION_KEYS,
  OUTPUT_TYPE_OPTIONS,
  ACTION_LIST,
  TOOL_PREFIEX,
};
