# x2t

`x2t` 是用于将 `.xlsx` 和 `.xls` 文件转换为 `HTML Table` 的一个工具

> x2t 目前仅支持单元格数据比较简单的表格，如字符串，数字，链接等，对于单元格数值是公式、根据其他单元格计算得出以及复杂富文本等情况支持不太好，如有场景无法转换成功，欢迎提 issue

## 安装
```shell
  npm i x2t -g
```

## 使用

1. 转换文件
```shell
  x2t [path]
```
  * 转换当前 `终端目录` 下的所有 `.xlsx` 和 `.xls` 文件
```shell
  x2t
```

  * 转换指定 `文件夹` 或 `文件`
```shell
  x2t path
```
Tips: 如果 `path` 指向一个文件夹，则会递归遍历转换里面所有的 `.xlsx` 和 `.xls` 文件

2. 设置时间格式化的格式，默认值 `yyyy-MM-dd`
```shell
  x2t config.set.fmt <fmt>
```
Tips: 因为 `.xlsx` 和 `.xls` 的日期格式有多种尚未见过的 `fmt` ，当单元格值是日期格式的时候，获取到单元格的 `numFmt` 后使用现有的 `moment` 库可能无法正确的格式化时间，故提供 `config.set.fmt` 的配置，可以统一单元格的日期格式