const chalk = require("chalk");
const { TOOL_PREFIEX } = require("./constants");
const errHanlders = (msg) => {
  console.log(`${chalk.red(`${TOOL_PREFIEX}${msg}`)} `);
  throw Error(msg);
};

module.exports = {
  errHanlders,
};
