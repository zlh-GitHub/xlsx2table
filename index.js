#!/usr/bin/env node

const excelJs = require('exceljs');
const fs = require('fs');
const path = require('path');
const ejs = require('ejs');
const { program } = require('commander');
const moment = require('moment');
const chalk = require('chalk');
const cliProgress = require('cli-progress');
const ora = require('ora');
const { isObject } = require('./utils');
const config = require('./config.json');
const { UPDATE_CONFIG_TYPE, EXTNAMES } = require('./constants');

const VALUE_TYPE = excelJs.ValueType;

program.version('1.0.0', '-v, -version');

program.addHelpText('beforeAll', 'x2t xlsx2table');

/**
 * 渲染HTML
 * @param {Object} params ejs参数对象
 * @param {String} params.tableTitle 页面标题
 * @param {String} params.tableStr table字符串
 * @returns {Promise}
 */
const renderHtml = params => new Promise((resolve, reject) => {
  ejs.renderFile(
    path.resolve(__dirname, './public/template.ejs'),
    params,
    (err, data) => {
      if (err) {
        reject(err);
        return;
      }
      resolve(data);
    },
  );
});

/**
 * 格式化value值
 * @param {Number} type 值类型
 * @param {any} model 
 * @returns {String}
 */
const getValue = (type, model) => {
  // https://www.npmjs.com/package/exceljs#value-types
  const { result, value, style, hyperlink, text } = model;
  switch (type) {
    case VALUE_TYPE.String:
    case VALUE_TYPE.SharedString:
    case VALUE_TYPE.Number:
      return value;
    case VALUE_TYPE.Boolean:
      return JSON.stringify(value);
    case VALUE_TYPE.Date:
      // 日期值，需要进行format
      return moment(value).format(config.fmt || style.numFmt);
    case VALUE_TYPE.Error:
      return value.error;
    case VALUE_TYPE.Formula:
      // 公式
      return isObject(result) ? JSON.stringify(result) : result;
    case VALUE_TYPE.Hyperlink:
      return `<a href="${hyperlink}">${text}</a>`;
    case VALUE_TYPE.Merge:
      // 表被合并的单元格，不应该渲染td
      return '';
    case VALUE_TYPE.Null:
      // 表该单元格没有内容，但是还是要渲染td
      return '';
    case VALUE_TYPE.RichText:
      return value.richText?.map(i => i?.text).join('');
    default: return '';
  }
}

/**
 * 输入一个路径，返回该路径下的所有excel文件绝对路径，否则返回空数组
 * @param {String} absolutePath 
 * @returns {Array[String]}
 */
const getExcelFile = absolutePath => {
  try {
    const stat = fs.statSync(absolutePath);
    if (stat.isDirectory()) {
      const filenames = fs.readdirSync(absolutePath);
      return filenames.reduce((res, nextFilename) => {
        const temp = getExcelFile(path.resolve(absolutePath, nextFilename));
        return [...res, ...temp];
      }, []);
    } else if (EXTNAMES.includes(path.extname(absolutePath))) {
      return [absolutePath];
    } else {
      return [];
    }
  } catch {
    return [];
  }
};

const xlsx2table = async filename => {
  const spinner = ora('正在读取需要处理的文件...').start();
  const absolutePath = path.resolve(process.cwd(), filename || './');
  const excelFilenames = getExcelFile(absolutePath); // 所有excel文件的绝对路径数组
  if (!excelFilenames.length) {
    spinner.warn(chalk.yellow('本次没有需要处理的文件，请确认路径是否正确！'));
    return;
  }
  spinner.succeed(chalk.green(`本次共读取到${excelFilenames.length}个需要处理的文件`));
  for (let excelFilename of excelFilenames) {
    try {
      const workbook = new excelJs.Workbook();
      const fileMsg = path.parse(excelFilename);
      const file = fs.readFileSync(excelFilename);
      await workbook.xlsx.load(file);
      const progressBar = new cliProgress.SingleBar({
        format: chalk.green('转换中|{bar}| {percentage}% || 子表：{value}/{total}') + `   文件路径：${excelFilename}`,
        barCompleteChar: '\u2588',
        barIncompleteChar: '\u2591',
        hideCursor: true
      });
      progressBar.start(workbook.worksheets.length, 0);
      for (let sheet of workbook.worksheets) {
        const merges = sheet._merges; // 单元格合并情况数据
        const tableStr = `
          <table border="5">
            <tbody>
              ${Array.from(sheet._rows).map(row => {
          if (!row) {
            return null;
          }
          return `
                <tr>
                  ${Array.from(row._cells).map(cell => {
            if (!cell) {
              return '<td></td>'
            }
            const { model, address, type } = cell;
            const value = getValue(type, model);
            const { top, bottom, left, right } = merges[address]?.model || {};
            const rowSpan = (bottom - top + 1) || '';
            const colSpan = (right - left + 1) || '';
            return type !== VALUE_TYPE.Merge ? `
                        <td
                          ${colSpan ? `colspan=${colSpan}` : ''}
                          ${rowSpan ? `rowspan=${rowSpan}` : ''}
                        >
                          ${value}
                        </td>
                      ` : '';
          }).join('')
            }
                </tr>
              `
        }).join('')
          }
            </tbody>
          </table>
        `;
        const tableTitle = `${fileMsg.name}-${sheet.name}`;
        const data = await renderHtml({
          tableTitle,
          tableStr,
        });
        fs.writeFileSync(path.resolve(fileMsg.dir, `./${tableTitle}.html`), data);
        progressBar.increment();
      }
      progressBar.stop();
    } catch (error) {
      console.error(error.message);
    }
  }
};

/**
 * 更新配置文件
 * @param {Object} action 
 * @param {String} key
 * @param {String} value
 * @param {String} type update | delete
 * @returns 
 */
const updateConfig = ({
  key,
  value,
  type,
}) => {
  const configPath = path.resolve(__dirname, './config.json');
  let config = fs.readFileSync(configPath, {
    encoding: 'utf-8',
  });
  config = JSON.parse(config);
  switch (type) {
    case UPDATE_CONFIG_TYPE.UPDATE:
      config[key] = value;
      break;
    case UPDATE_CONFIG_TYPE.DELETE:
      delete config[key];
      break;
    default:
      console.error(`type应为：${Object.keys(UPDATE_CONFIG_TYPE).join(' | ')}中的一个`);
  }
  config = JSON.stringify(config);
  fs.writeFileSync(configPath, config);
}

program
  .argument('[path]', '文件路径')
  .action(xlsx2table)

program
  .command('config.set.fmt')
  .description('统一设置时间格式')
  .argument('<fmt>', '时间格式')
  .action(fmt => updateConfig({
    key: 'fmt',
    value: fmt,
    type: UPDATE_CONFIG_TYPE.UPDATE
  }))

program.parse(process.argv); // 该行代码必须放在program配置下面