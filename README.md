# PyOffice

## 简述

基于 PyWin32 实现对 Excel、Word 等操作。

## 简单使用

### Excel

#### 打开工作簿

```python
from pyoffice.excel import Workbook

wb = Workbook()
wb.open('test.xlsx')
```

#### 获取激活的工作表

```python
ws = wb.getActiveWorkSheet()
print(ws.getName())
```

#### 根据工作表名称获取工作表

```python
ws = wb.getWorkSheetByName('Sheet1')
print(ws.getName())
```

#### 获取工作表中已使用的区域

```python
rg = ws.getUsedRange()
print(rg.getAddress())
```

#### 获取区域中的值

```python
val = rg.getValue()
print(val)
```

#### 获取单元格

```python
cell = ws.getCellByAddress('A2')
print(cell.getAddress())
print(cell.getValue())
```

#### 将值写入单元格

```python
cell.setValue(1)
cell.setValue('2')
cell.setValue([1, 2, 3])
cell.setValue([[1, 2, 3],
               [4, 5, 6]])
```

## 致谢

1. [ruofeng216](https://github.com/ruofeng216)
2. [giftbox](https://github.com/giftbox)

## 引用

1. [Application object (Excel)](https://docs.microsoft.com/en-us/office/vba/api/excel.application(object))
