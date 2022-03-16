/**
 * 返回值类型
 * @param {any} obj 
 * @returns 
 */
const getType = obj =>
  Reflect.apply(Object.prototype.toString, obj, []).replace(/^\[object\s(\w+)\]$/, '$1').toLowerCase();

/**
 * 返回给定的值是不是对象
 * @param {any} obj 
 * @returns {Boolean}
 */
const isObject = obj => getType(obj) === 'object';

/**
 * 返回给定的值是不是函数
 * @param {any} fn 
 * @returns {Boolean}
 */
const isFunction = fn => getType(fn) === 'function';

module.exports = {
  getType,
  isObject,
  isFunction
}