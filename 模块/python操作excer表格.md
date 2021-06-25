<h1 align='center'>python操作excer表格</h1>

[TOC]

# excer表格处理

## 处理excel表格都有哪些库
![image-20210621194311145](https://cdn.jsdelivr.net/gh/yanshaung/pic/image-20210621194311145.png)

> 各个库之间的区别
>
> - xlrd[^xlrd简介]
>
> [^xlrd简介]:只能读取文件
>
> - xlwt[^xlwt简介]
>
> [^xlwt简介]:只能写入文件
>
> - xlitils
> - xlwings
> - openpyxl
>
> [^openxl简介]:适合程序员使用,因为可以直接使用表格对应的(a,b,c)和(1,2.3)来确定索引位置
>
> - xlswriter
> - win32com
> - DataNitro[^datanitro简介]
>
> [^datanitro简介]:不建议学习,因为该模块收费
>
> - pandas

---





## xlrd读取表格

```python
import xlrd
# 打开表格
xlsl = xlrd.open_workbook('excer/耳带账单记录.xls')
# 打开某个sheet
sheet = xlsl.sheet_by_index(1)
# 读取某一行数据
data = sheet.cell_value(5,1)
# 打印某一行数据
print(data)
```



---



## xlwt写入表格

```python
import xlwt

# 新建工作薄
new_workbook = xlwt.Workbook()
# 新建一个sheet
worksheet = new_workbook.add_sheet('new_test')
# 在某一行写入颜霜
worksheet.write(2, 2, '颜霜')
# 保存该xls文件
new_workbook.save('text.xls')
```



---



## :fish: ​==重点:==xlwings的使用



### :fish::fish: 安装

> ==安装:==`pip install xlwings`



### :fish::fish:基本操作

```python
import xlwings as xs

# 打开excer应用
app = xs.App()   # 打开应用
app = xs.App(visible=False)   # 后台编辑不打开应用
app = xs.App(add_book=False)   # 只打开一个excer文档

# 新建工作薄
wb = app.books.add()

# 打开工作薄
wb = app.books.open()

# 新建工作表
sht = wb.sheets["sheet1"]

# 写入内容
sht.range('a1').value="颜霜"   # 写入单个
sht.range("c4").value[5,6,7,8]   # 写整行(因为默认写整行,如果写整列需要下行代码的操作)
sht.range("c5").options(transpose=True).value=[5,6,7,8]   # 写入整列
sht.range("c6").value[[1,2], [3,4]]   # 插入行列

# 读取内容
print(sht.range('a1').value)   # 读取单个内容
print(sht.range("c4:c8").value)   # 读整行
print(sht.range("c4:f4").value)   # 读整列
print(sht.range("a1:z23").value)   # 读行列


# 保存excer文件
wb.save('ys.xlsx')   # 如果是打开只读该位置可以只为括号

# 关闭excer程序
wb.close()
app.quit()
```







### :fish::fish:动态模板(pycharm可用)

```python
import xlwings as xs

app = xs.App()   # 打开应用

wb = app.books.open()   # 读取应用

sht = wb.sheets["sheet1"]   # 读取单元表

sht.range('a1').value="颜霜"   # 写入单个

print(sht.range('a1').value)   # 读取单个内容

wb.save('ys.xlsx')   # 如果是打开只读该位置可以只为括号

wb.close()
app.quit()
```





