#!/usr/bin/env node
'use strict';

require('regenerator-runtime/runtime.js');
require('core-js/modules/es.object.to-string.js');
require('core-js/modules/es.promise.js');
require('core-js/modules/es.array.concat.js');
require('core-js/modules/es.array.join.js');
require('core-js/modules/es.array.map.js');
require('core-js/modules/es.array.reduce.js');
require('core-js/modules/es.array.includes.js');
require('core-js/modules/es.string.includes.js');
require('core-js/modules/es.string.iterator.js');
require('core-js/modules/es.array.iterator.js');
require('core-js/modules/web.dom-collections.iterator.js');
require('core-js/modules/es.array.from.js');
require('core-js/modules/es.function.name.js');
require('core-js/modules/es.object.keys.js');
require('core-js/modules/es.symbol.js');
require('core-js/modules/es.symbol.description.js');
var require$$0 = require('exceljs');
var require$$1 = require('fs');
var require$$2 = require('path');
var require$$3 = require('ejs');
var require$$4 = require('commander');
var require$$5 = require('moment');
var require$$6 = require('chalk');
var require$$7 = require('cli-progress');
var require$$8 = require('ora');
require('core-js/modules/es.regexp.exec.js');
require('core-js/modules/es.string.replace.js');
require('core-js/modules/es.reflect.apply.js');

function _interopDefaultLegacy (e) { return e && typeof e === 'object' && 'default' in e ? e : { 'default': e }; }

var require$$0__default = /*#__PURE__*/_interopDefaultLegacy(require$$0);
var require$$1__default = /*#__PURE__*/_interopDefaultLegacy(require$$1);
var require$$2__default = /*#__PURE__*/_interopDefaultLegacy(require$$2);
var require$$3__default = /*#__PURE__*/_interopDefaultLegacy(require$$3);
var require$$4__default = /*#__PURE__*/_interopDefaultLegacy(require$$4);
var require$$5__default = /*#__PURE__*/_interopDefaultLegacy(require$$5);
var require$$6__default = /*#__PURE__*/_interopDefaultLegacy(require$$6);
var require$$7__default = /*#__PURE__*/_interopDefaultLegacy(require$$7);
var require$$8__default = /*#__PURE__*/_interopDefaultLegacy(require$$8);

function asyncGeneratorStep(gen, resolve, reject, _next, _throw, key, arg) {
  try {
    var info = gen[key](arg);
    var value = info.value;
  } catch (error) {
    reject(error);
    return;
  }

  if (info.done) {
    resolve(value);
  } else {
    Promise.resolve(value).then(_next, _throw);
  }
}

function _asyncToGenerator(fn) {
  return function () {
    var self = this,
        args = arguments;
    return new Promise(function (resolve, reject) {
      var gen = fn.apply(self, args);

      function _next(value) {
        asyncGeneratorStep(gen, resolve, reject, _next, _throw, "next", value);
      }

      function _throw(err) {
        asyncGeneratorStep(gen, resolve, reject, _next, _throw, "throw", err);
      }

      _next(undefined);
    });
  };
}

function _toConsumableArray(arr) {
  return _arrayWithoutHoles(arr) || _iterableToArray(arr) || _unsupportedIterableToArray(arr) || _nonIterableSpread();
}

function _arrayWithoutHoles(arr) {
  if (Array.isArray(arr)) return _arrayLikeToArray(arr);
}

function _iterableToArray(iter) {
  if (typeof Symbol !== "undefined" && iter[Symbol.iterator] != null || iter["@@iterator"] != null) return Array.from(iter);
}

function _unsupportedIterableToArray(o, minLen) {
  if (!o) return;
  if (typeof o === "string") return _arrayLikeToArray(o, minLen);
  var n = Object.prototype.toString.call(o).slice(8, -1);
  if (n === "Object" && o.constructor) n = o.constructor.name;
  if (n === "Map" || n === "Set") return Array.from(o);
  if (n === "Arguments" || /^(?:Ui|I)nt(?:8|16|32)(?:Clamped)?Array$/.test(n)) return _arrayLikeToArray(o, minLen);
}

function _arrayLikeToArray(arr, len) {
  if (len == null || len > arr.length) len = arr.length;

  for (var i = 0, arr2 = new Array(len); i < len; i++) arr2[i] = arr[i];

  return arr2;
}

function _nonIterableSpread() {
  throw new TypeError("Invalid attempt to spread non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method.");
}

function _createForOfIteratorHelper(o, allowArrayLike) {
  var it = typeof Symbol !== "undefined" && o[Symbol.iterator] || o["@@iterator"];

  if (!it) {
    if (Array.isArray(o) || (it = _unsupportedIterableToArray(o)) || allowArrayLike && o && typeof o.length === "number") {
      if (it) o = it;
      var i = 0;

      var F = function () {};

      return {
        s: F,
        n: function () {
          if (i >= o.length) return {
            done: true
          };
          return {
            done: false,
            value: o[i++]
          };
        },
        e: function (e) {
          throw e;
        },
        f: F
      };
    }

    throw new TypeError("Invalid attempt to iterate non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method.");
  }

  var normalCompletion = true,
      didErr = false,
      err;
  return {
    s: function () {
      it = it.call(o);
    },
    n: function () {
      var step = it.next();
      normalCompletion = step.done;
      return step;
    },
    e: function (e) {
      didErr = true;
      err = e;
    },
    f: function () {
      try {
        if (!normalCompletion && it.return != null) it.return();
      } finally {
        if (didErr) throw err;
      }
    }
  };
}

var xlsx2table$1 = {};

var getType = function getType(obj) {
  return Reflect.apply(Object.prototype.toString, obj, []).replace(/^\[object\s(\w+)\]$/, '$1').toLowerCase();
};
/**
 * 返回给定的值是不是对象
 * @param {any} obj 
 * @returns {Boolean}
 */


var isObject$1 = function isObject(obj) {
  return getType(obj) === 'object';
};
/**
 * 返回给定的值是不是函数
 * @param {any} fn 
 * @returns {Boolean}
 */


var isFunction = function isFunction(fn) {
  return getType(fn) === 'function';
};

var utils = {
  getType: getType,
  isObject: isObject$1,
  isFunction: isFunction
};

var fmt = "yyyy-MM-dd";
var require$$10 = {
	fmt: fmt
};

var UPDATE_CONFIG_TYPE$1 = {
  UPDATE: 'update',
  DELETE: 'delete'
};
var EXTNAMES$1 = ['.xlsx', '.xls'];
var constants = {
  UPDATE_CONFIG_TYPE: UPDATE_CONFIG_TYPE$1,
  EXTNAMES: EXTNAMES$1
};

var excelJs = require$$0__default["default"];
var fs = require$$1__default["default"];
var path = require$$2__default["default"];
var ejs = require$$3__default["default"];
var program = require$$4__default["default"].program;
var moment = require$$5__default["default"];
var chalk = require$$6__default["default"];
var cliProgress = require$$7__default["default"];
var ora = require$$8__default["default"];
var isObject = utils.isObject;
var config = require$$10;
var UPDATE_CONFIG_TYPE = constants.UPDATE_CONFIG_TYPE,
    EXTNAMES = constants.EXTNAMES;
var VALUE_TYPE = excelJs.ValueType;
program.version('1.0.0', '-v, -version');
program.addHelpText('beforeAll', 'x2t xlsx2table');
/**
 * 渲染HTML
 * @param {Object} params ejs参数对象
 * @param {String} params.tableTitle 页面标题
 * @param {String} params.tableStr table字符串
 * @returns {Promise}
 */

var renderHtml = function renderHtml(params) {
  return new Promise(function (resolve, reject) {
    ejs.renderFile(path.resolve(__dirname, './public/template.ejs'), params, function (err, data) {
      if (err) {
        reject(err);
        return;
      }

      resolve(data);
    });
  });
};
/**
 * 格式化value值
 * @param {Number} type 值类型
 * @param {any} model 
 * @returns {String}
 */


var getValue = function getValue(type, model) {
  var _value$richText;

  // https://www.npmjs.com/package/exceljs#value-types
  var result = model.result,
      value = model.value;
      model.style;
      var hyperlink = model.hyperlink,
      text = model.text;

  switch (type) {
    case VALUE_TYPE.String:
    case VALUE_TYPE.SharedString:
    case VALUE_TYPE.Number:
      return value;

    case VALUE_TYPE.Boolean:
      return JSON.stringify(value);

    case VALUE_TYPE.Date:
      // 日期值，需要进行format
      return moment(value).format(config.fmt );

    case VALUE_TYPE.Error:
      return value.error;

    case VALUE_TYPE.Formula:
      // 公式
      return isObject(result) ? JSON.stringify(result) : result;

    case VALUE_TYPE.Hyperlink:
      return "<a href=\"".concat(hyperlink, "\">").concat(text, "</a>");

    case VALUE_TYPE.Merge:
      // 表被合并的单元格，不应该渲染td
      return '';

    case VALUE_TYPE.Null:
      // 表该单元格没有内容，但是还是要渲染td
      return '';

    case VALUE_TYPE.RichText:
      return (_value$richText = value.richText) === null || _value$richText === void 0 ? void 0 : _value$richText.map(function (i) {
        return i === null || i === void 0 ? void 0 : i.text;
      }).join('');

    default:
      return '';
  }
};
/**
 * 输入一个路径，返回该路径下的所有excel文件绝对路径，否则返回空数组
 * @param {String} absolutePath 
 * @returns {Array[String]}
 */


var getExcelFile = function getExcelFile(absolutePath) {
  try {
    var stat = fs.statSync(absolutePath);

    if (stat.isDirectory()) {
      var filenames = fs.readdirSync(absolutePath);
      return filenames.reduce(function (res, nextFilename) {
        var temp = getExcelFile(path.resolve(absolutePath, nextFilename));
        return [].concat(_toConsumableArray(res), _toConsumableArray(temp));
      }, []);
    } else if (EXTNAMES.includes(path.extname(absolutePath))) {
      return [absolutePath];
    } else {
      return [];
    }
  } catch (_unused) {
    return [];
  }
};

var xlsx2table = /*#__PURE__*/function () {
  var _ref = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee(filename) {
    var spinner, absolutePath, excelFilenames, _iterator, _step, excelFilename, workbook, fileMsg, file, progressBar, _iterator2, _step2, _loop;

    return regeneratorRuntime.wrap(function _callee$(_context2) {
      while (1) {
        switch (_context2.prev = _context2.next) {
          case 0:
            spinner = ora('正在读取需要处理的文件...').start();
            absolutePath = path.resolve(process.cwd(), filename || './');
            excelFilenames = getExcelFile(absolutePath); // 所有excel文件的绝对路径数组

            if (excelFilenames.length) {
              _context2.next = 6;
              break;
            }

            spinner.warn(chalk.yellow('本次没有需要处理的文件，请确认路径是否正确！'));
            return _context2.abrupt("return");

          case 6:
            spinner.succeed(chalk.green("\u672C\u6B21\u5171\u8BFB\u53D6\u5230".concat(excelFilenames.length, "\u4E2A\u9700\u8981\u5904\u7406\u7684\u6587\u4EF6")));
            _iterator = _createForOfIteratorHelper(excelFilenames);
            _context2.prev = 8;

            _iterator.s();

          case 10:
            if ((_step = _iterator.n()).done) {
              _context2.next = 44;
              break;
            }

            excelFilename = _step.value;
            _context2.prev = 12;
            workbook = new excelJs.Workbook();
            fileMsg = path.parse(excelFilename);
            file = fs.readFileSync(excelFilename);
            _context2.next = 18;
            return workbook.xlsx.load(file);

          case 18:
            progressBar = new cliProgress.SingleBar({
              format: chalk.green('转换中|{bar}| {percentage}% || 子表：{value}/{total}') + "   \u6587\u4EF6\u8DEF\u5F84\uFF1A".concat(excelFilename),
              barCompleteChar: "\u2588",
              barIncompleteChar: "\u2591",
              hideCursor: true
            });
            progressBar.start(workbook.worksheets.length, 0);
            _iterator2 = _createForOfIteratorHelper(workbook.worksheets);
            _context2.prev = 21;
            _loop = /*#__PURE__*/regeneratorRuntime.mark(function _loop() {
              var sheet, merges, tableStr, tableTitle, data;
              return regeneratorRuntime.wrap(function _loop$(_context) {
                while (1) {
                  switch (_context.prev = _context.next) {
                    case 0:
                      sheet = _step2.value;
                      merges = sheet._merges; // 单元格合并情况数据

                      tableStr = "\n          <table border=\"5\">\n            <tbody>\n              ".concat(Array.from(sheet._rows).map(function (row) {
                        if (!row) {
                          return null;
                        }

                        return "\n                <tr>\n                  ".concat(Array.from(row._cells).map(function (cell) {
                          var _merges$address;

                          if (!cell) {
                            return '<td></td>';
                          }

                          var model = cell.model,
                              address = cell.address,
                              type = cell.type;
                          var value = getValue(type, model);

                          var _ref2 = ((_merges$address = merges[address]) === null || _merges$address === void 0 ? void 0 : _merges$address.model) || {},
                              top = _ref2.top,
                              bottom = _ref2.bottom,
                              left = _ref2.left,
                              right = _ref2.right;

                          var rowSpan = bottom - top + 1 || '';
                          var colSpan = right - left + 1 || '';
                          return type !== VALUE_TYPE.Merge ? "\n                        <td\n                          ".concat(colSpan ? "colspan=".concat(colSpan) : '', "\n                          ").concat(rowSpan ? "rowspan=".concat(rowSpan) : '', "\n                        >\n                          ").concat(value, "\n                        </td>\n                      ") : '';
                        }).join(''), "\n                </tr>\n              ");
                      }).join(''), "\n            </tbody>\n          </table>\n        ");
                      tableTitle = "".concat(fileMsg.name, "-").concat(sheet.name);
                      _context.next = 6;
                      return renderHtml({
                        tableTitle: tableTitle,
                        tableStr: tableStr
                      });

                    case 6:
                      data = _context.sent;
                      fs.writeFileSync(path.resolve(fileMsg.dir, "./".concat(tableTitle, ".html")), data);
                      progressBar.increment();

                    case 9:
                    case "end":
                      return _context.stop();
                  }
                }
              }, _loop);
            });

            _iterator2.s();

          case 24:
            if ((_step2 = _iterator2.n()).done) {
              _context2.next = 28;
              break;
            }

            return _context2.delegateYield(_loop(), "t0", 26);

          case 26:
            _context2.next = 24;
            break;

          case 28:
            _context2.next = 33;
            break;

          case 30:
            _context2.prev = 30;
            _context2.t1 = _context2["catch"](21);

            _iterator2.e(_context2.t1);

          case 33:
            _context2.prev = 33;

            _iterator2.f();

            return _context2.finish(33);

          case 36:
            progressBar.stop();
            _context2.next = 42;
            break;

          case 39:
            _context2.prev = 39;
            _context2.t2 = _context2["catch"](12);
            console.error(_context2.t2.message);

          case 42:
            _context2.next = 10;
            break;

          case 44:
            _context2.next = 49;
            break;

          case 46:
            _context2.prev = 46;
            _context2.t3 = _context2["catch"](8);

            _iterator.e(_context2.t3);

          case 49:
            _context2.prev = 49;

            _iterator.f();

            return _context2.finish(49);

          case 52:
          case "end":
            return _context2.stop();
        }
      }
    }, _callee, null, [[8, 46, 49, 52], [12, 39], [21, 30, 33, 36]]);
  }));

  return function xlsx2table(_x) {
    return _ref.apply(this, arguments);
  };
}();
/**
 * 更新配置文件
 * @param {Object} action 
 * @param {String} key
 * @param {String} value
 * @param {String} type update | delete
 * @returns 
 */


var updateConfig = function updateConfig(_ref3) {
  var key = _ref3.key,
      value = _ref3.value,
      type = _ref3.type;
  var configPath = path.resolve(__dirname, './config.json');
  var config = fs.readFileSync(configPath, {
    encoding: 'utf-8'
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
      console.error("type\u5E94\u4E3A\uFF1A".concat(Object.keys(UPDATE_CONFIG_TYPE).join(' | '), "\u4E2D\u7684\u4E00\u4E2A"));
  }

  config = JSON.stringify(config);
  fs.writeFileSync(configPath, config);
};

program.argument('[path]', '文件路径').action(xlsx2table);
program.command('config.set.fmt').description('统一设置时间格式').argument('<fmt>', '时间格式').action(function (fmt) {
  return updateConfig({
    key: 'fmt',
    value: fmt,
    type: UPDATE_CONFIG_TYPE.UPDATE
  });
});
program.parse(process.argv); // 该行代码必须放在program配置下面

module.exports = xlsx2table$1;
