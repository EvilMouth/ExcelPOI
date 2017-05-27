# ExcelPOI
read and write Excel by POI

## 使用POI读写Excel
由于Apache POI跟Android不兼容，所以不能直接用maven拉依赖，需要重新打包->（此项目library下libs里面的两个poi开头的jar包）所以需要通过导jar包的方式导入自己项目

### double to String
获取单元格内容`cell.getCellValue`中数字格式都会以double格式返回，所以可以用`NumberToTextConverter.toText(cell.getNumericCellValue());`转成String
