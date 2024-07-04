## 项目功能

合并文件夹中多个excel文件，并在合并后的表格中增加一个列

## 项目目的

练习Python语法、学习Pandas库、os模块、yaml模块和openpyxl库，通过该项目能够熟练掌握DataFrame数据结构和Series数据结构，熟练的通过python操作Excel表格，做数据分析和处理。学会通过os模块操作文件系统，利用yaml来读取yaml的配置文件、熟练掌握openpyxl库操作Excel表格的样式，清楚WorkBook、Font、Alignment、Border、Side等openpyxl内置的对象。

## 项目涉及技术

**项目中用到了pandas、os、openpyxl、yaml库**

1、使用pandas读取分析和处理excel表格

2、使用os模块判断文件的路径是否正确

3、使用yaml库读取yaml配置文件

4、使用openpyxl库改变表格样式，实现Excel文件内容的美化

## 项目配置

在config.yaml文件中配置read_path表示需要读取的文件夹目录，该脚本会自动递归地遍历目录下的所有子文件

1. 在main中调用read_config()方法, 返回要读取的目录和输出的文件路径
2. get_files方法中传入目录，会返回该目录下包括子文件夹中，所有的文件，返回一个绝对路径的列表
3. merge_excels方法接收文件列表和输出路径，是程序的中枢
4. 在merge_excels方法中使用pandas的read_excel方法读取每一个xlsx文件，经处理后合并成最终的combined_excel.xlsx文件
5. handle_data方法接收一个site表示要添加的新的列的值，temp_df是要处理的DataFrame对象
6. 调用prettify_page()方法对输出的combined_excel.xlsx文件调整字体、边框和行高、列宽

## 笔记

### python接收多个参数

**参数使用*arg接收，并且如果有多个参数，该参数必须放在末尾, 在代码块中通过for in遍历**

```python
def combine_dataframe(*args: pd.DataFrame) -> pd.DataFrame:
    # 创建一个空DataFrame等待接收合并
    combine_df = pd.DataFrame()
    # 迭代处理每个传入的 DataFrame 参数
    for arg in args:
        # 合并每个 DataFrame 到 combined_df，concat是pandas的函数，ignore_index表示是否忽略并建立新的索引（序号）
        combine_df = pd.concat([combine_df, arg], ignore_index=True)
    return combine_df

```

### OS模块

**os是Python中内置的模块，与操作系统中的文件系统相关。记录了如下一些基本的操作**
1、getcwd(): 获取当前工作目录, 返回字符串
2、mkdir(): 和linux中的命令相同，只能创建一个目录
3、removedirs(): 递归删除目录 rm -r
4、rmdir(): 删除空目录
5、rename(): 重命名文件或目录
6、listdir(): 返回指定路径下的目录名和文件名, 既有目录也有文件
7、makedirs(): 递归地创建目录，像linux中的mkdir -p
8、walk(): 递归地遍历指定目录下的所有目录，一般配合for循环使用，每次返回3个值：当前路径(str),路径下的文件夹(列表),
路径下的文件(列表)

```python
# os.walk(dir_path)：会递归地遍历指定文件夹下所有的层级目录
# 每一次递归，返回root、 dirs、 files,分别表示当前的目录路径、当前目录路径包含的子目录、当前目录路径包含的文件
for root, dirs, files in os.walk(path):
    # 遍历当前层级的文件
    for file in files:
        # 判断是否是要收集的文件
        if file.endswith(".xlsx"):
            files_path.append(os.path.join(root, file))
```

### yaml模块

**yaml模块用于读取yaml文件配置文件, YAML是一种人类可读的数据序列化格式，可以处理复杂的数据结构，并支持注释，适合于需要更复杂配置的情况。
**
下面是一个yaml文件的案例

```yaml
database:
  host: localhost
  port: 3306
  username: admin
  password: secret

paths:
  input_dir: /path/to/input
  output_dir: /path/to/output

options:
  verbose: true
  max_attempts: 5
```

使用python程序读取

```python
import yaml

# 读取配置文件
with open('config.yaml', 'r') as f:
    config = yaml.safe_load(f)

# 使用配置信息
print(config['database']['host'])
print(config['paths']['output_dir'])
```

### Pandas模块

1、DataFrame（结构）：在Pandas中有一个特殊的数据类型DataFrame，每一个表格类型可以看做是一个DataFrame，可以理解为一个二维数组，然后每个列都有一个列名，每一行都有个索引。

2、Series（一系列，连续）：在一个DataFrame中，每一列都是一个Series对象，DataFrame 可以看作是由若干个 Series
组成的字典，每一个Series的数据类型可以不同。

**Series 是 Pandas 中用于存储一维数据的基本数据结构，而 DataFrame 则是由多个 Series 组成的二维数据结构。**

在项目中使用DataFrame收集Excel数据, 合并数据, 并输出到Excel文件中

**创建DataFrame的方法如下：**

1. 创建一个DataFrame类型的对象

```python
import pandas

pandas.DataFrame()
```

2. 创建一个DataFrame对象，并指定列名标签通过columns

```python
import pandas

pandas.DataFrame(columns=['Column1', 'Column12', 'Column3'])
```

3. 查看DataFrame数据的方式

- .head() 方法显示 DataFrame 的前几行，默认是显示前5行。
- .tail() 方法显示 DataFrame 的最后几行，默认也是显示最后5行。

```python
# 显示前10行
print(df.head(10))

# 显示最后10行
print(df.tail(10))
```

- 可以通过 Pandas 的显示选项 pd.set_option() 来调整显示效果，比如设置显示的最大行数和列数等。

```python
# 设置显示的最大行数和列数
pd.set_option('display.max_rows', 20)  # 设置显示最大行数为20
pd.set_option('display.max_columns', None)  # 显示所有列，None 表示没有限制
```

- 使用 .iloc 和 .loc 可以按位置和标签选择数据，无论索引是否很长，这两个方法都能有效地访问和操作数据。

```python
# 使用 .iloc 选择前10行和所有列
print(df.iloc[:10, :])

# 使用 .loc 选择特定标签或范围
print(df.loc[:, ['Name', 'Age']])
```

- 使用 .shape 属性查看 DataFrame 的行数和列数，以了解整体数据规模。
- 使用 .info() 方法查看 DataFrame 的简洁摘要，包括列名、非空值数量、数据类型等信息。

```python
# 查看 DataFrame 的形状
print(df.shape)

# 查看 DataFrame 的信息摘要
print(df.info())
```

4. 从excel文件中读取并返回DataFrame对象

```python
import pandas

df = pandas.read_excel('your_file.xlsx')
```

5. 通过shape获取当前DataFrame对象的(行, 列), 返回的是元组类型

```python
import pandas

df = pandas.DataFrame()
print(df.shape)
```

6. 通过loc或iloc获取DataFrame对象的某行

- loc：基于行、列标签，适用于字符串或时间戳索引，末端闭合。
- iloc：基于位置，适用于整数索引，末端开放。

```python
import pandas as pd

df = pd.DataFrame()
# 获取第一行
print("iloc[0]\n", df.iloc[0])
# 获取所有的行，第一列
print("iloc[:, 0]\n", df.iloc[:, 0])
# 获取所有的行，第一列iloc不含末尾
print("iloc[0:2, 0:1]\n", df.iloc[0:2, 0:1])
# 获取所有的行，第一列loc含末尾，并且列标签不能是数字
print("loc[0:2,'Column1']\n", df.loc[0:2, 'Column1':'Column4']) 
```

7. columns属性, 通过DF对象的columns属性可以获取当前DF的所有列名, 可以通过这个属性修改列名，但是要保证前后列名的数量是一样的否则报错

```python
# 将列名修改为数据的第一行
data_frame.columns = data_frame.iloc[0].tolist()
```

8. 添加新的列操作, 可以通过直接赋值操作, 也可以通过Series操作
   **需注意的是，如果通过列表直接复制，必须保证列表长度和DataFrame的行相同**

```python
# 添加新的列，并且设置所有行为单一值
df['NewColumn1'] = "ABC"
df['NewColumn2'] = True  # 也可以是常量
df['NewColumn3'] = ['A', 'B', 'C']  # 要保证行数为3

# 通过Series, 使用pandas的Series函数创建新的Series对象
# index的设置确保了新创建的 Series 对象的索引与 DataFrame 的列名保持一致, 保证了数据的对齐
my_Series = pd.Series(dtype=str, index=df.columns)
df['NewColumn4'] = my
```

9. fillna方法用来填充DataFrame中的NaN值, inplace表示原地更改不返回新值

```python
data_frame.fillna("覆盖的值", inplace=True)
# 也可以指定列
data_frame['NewColumn1'].fillna("覆盖的值", inplace=True)
```

10. 删除行，使用drop(row_index)方法

```python
# 删除索引在第0行的数据
df..drop(0)
```

11. 拼接DataFrame数据, pandas提供了concat方法用于拼接多个DataFrame对象

```python
# 第一个参数是列表，里面放置多个DataFrame对象，合并后返回新的DataFrame对象
# ignore_index表示是否忽略并建立新的索引（序号），如果不忽略，则会将两个DataFrame的索引交叉合并,012012的方式
# 合并后的DataFrame的列标签是它们的并集, 故可以拿多列的和只有一列的DF数据合并
combine_df = pd.concat([combine_df, arg], ignore_index=True)
```

12. 写入excel文件, 通过to_excel方法

```python
# index设置是否将索引也存储, index_tabel用来设置索引列的标签
combined_df.to_excel(output_file, index=True, sheet_name='Sheet1', index_label='序号')
```

### re模块

**使用re模块执行正则表达式，在该项目中，通过该模块匹配文件中的指定括号内的文字**

```python
import re

# 使用正则表达式查找括号中的内容
# \( 和 \) 匹配左括号 ( 和右括号 ) 字符
# (.*?) 是一个捕获组，用于匹配括号内的任意字符。.*? 表示非贪婪匹配，即尽可能少地匹配任意字符，直到下一个括号闭合为止。
# re.search(pattern, string) 函数在给定的字符串 string 中搜索第一个匹配 pattern 的位置，并返回一个匹配对象。
# 如果未找到匹配的模式，match 将是 None
match = re.search(r'\((.*?)\)', text)
if match:
    return match.group(1)  # 返回括号中的内容类型str
else:
    return ''  # 如果找不到匹配的内容，返回空字符串
```

### openpyxl模块

**该项目中使用openpyxl模块对工作表的样式进行美化调整, 使用了如下几个部分**

1. load_workbook, 通过该函数加载文件，并返回WorkBook对象

```python
from openpyxl import load_workbook

# 加载现有Excel文件
wb = load_workbook(file)

```

2. WorkBook对象, 是openpyxl模块中Excel文件载体

```python
# 通过WorkBook对象获取默认的工作表
ws = wb.active
```

3. Border对象，表示单元格的边框属性

```python
 # 设置无框线样式
thin_border = Border(left=Side(style='none'),
                     right=Side(style='none'),
                     top=Side(style='none'),
                     bottom=Side(style='none'))
```

4. Font对象，字体对象，表示字体属性

```python
 # 设置字体大小为12
font = Font(size=12, bold=False)
for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
    for cell in row:
        cell.font = font
        cell.border = thin_border
```

5. Alignment对象，单词翻译为n.定义、对准，表示对齐方式

```python
 # 设置单元格对齐方式为居中
alignment = Alignment(horizontal='center', vertical='center')
for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
    for cell in row:
        cell.alignment = alignment
```

6. 一些常用的行列属性(row_dimensions, column_dimensions)。dimensions: n. 尺寸, 范围, 维度

```python
# ws表示当前工作表对象
# 获取行的最大数量
print(ws.max_row)
# 获取某行的对象
ws.column_dimensions[index]  # .height属性是行高
# 获取某列的对象
ws.column_dimensions['col']  # .width属性是列宽
# 获取某列的宽度
```

6. 设置行高和列宽

```python
# 设置行高为21
for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
    # 在每次迭代中，row 是一个元组，其中包含该行的所有单元格对象。row[0] 表示该元组中的第一个单元格对象, 通过 .row 属性获取其行号。
    ws.row_dimensions[row[0].row].height = 21
```

7. 保存修改

```python
# 保存修改后的工作簿
wb.save(file_path)
```

