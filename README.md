# CSV 批量转换工具

这是一个本地 CSV 批量转换工具，可以把 `.csv` 文件转换成：

- `.xlsx` Excel 表格
- `.txt` 类似表格排列的纯文本
- `.md` Markdown 表格

支持选择一个文件夹，也支持一次选择多个 CSV 文件。转换结果会自动生成到源文件所在位置的 `转换结果` 文件夹中。

## 文件说明

```text
csv_batch_convert.py       命令行批量转换主程序
csv_batch_convert_gui.py   图形界面程序
mac一键启动.command        macOS 双击启动脚本
win一键启动.bat            Windows 双击启动脚本
README.md                  使用说明
```

## 运行环境

需要安装 Python 3。

macOS 一般自带 Python 3，可以直接尝试运行。

Windows 如果双击后提示找不到 Python，请安装 Python 3，并在安装时勾选：

```text
Add Python to PATH
```

## macOS 使用方法

1. 把整个工具文件夹放到电脑上。
2. 双击运行：

```text
mac一键启动.command
```

3. 打开窗口后，选择：

```text
选择 CSV 文件（可多选）
```

或者：

```text
选择文件夹
```

4. 选择转换类型：

```text
只转 XLSX
只转 TXT 表格
只转 Markdown 表格
全部转换
```

5. 点击：

```text
开始批量转换
```

6. 转换完成后，结果会出现在源目录下的：

```text
转换结果
```

如果 macOS 提示无法打开 `.command` 文件，可以在终端进入工具目录后执行：

```bash
chmod +x mac一键启动.command
```

然后再双击运行。

## Windows 使用方法

1. 把整个工具文件夹复制到 Windows 电脑上。
2. 双击运行：

```text
win一键启动.bat
```

3. 在打开的窗口中选择 CSV 文件或文件夹。
4. 选择转换类型。
5. 点击开始转换。

转换结果会自动生成到源文件所在目录下的：

```text
转换结果
```

## 输出规则

如果选择的是文件夹，例如：

```text
D:\data\csv_files
```

结果会生成到：

```text
D:\data\csv_files\转换结果
```

如果选择的是 CSV 文件，例如：

```text
D:\data\a.csv
D:\data\b.csv
```

结果会生成到：

```text
D:\data\转换结果
```

如果一次选择了不同文件夹里的多个 CSV 文件，每个 CSV 的结果会放到它自己所在目录下的 `转换结果` 文件夹。

## 转换示例

源文件：

```text
test.csv
```

选择“全部转换”后，会生成：

```text
转换结果/test.xlsx
转换结果/test.txt
转换结果/test.md
```

## 命令行用法

也可以不用图形界面，直接用命令行运行。

转换当前目录下的所有 CSV：

```bash
python3 csv_batch_convert.py --infer-types
```

指定输入文件夹：

```bash
python3 csv_batch_convert.py /path/to/csv_folder -o /path/to/output --infer-types
```

递归扫描子文件夹：

```bash
python3 csv_batch_convert.py /path/to/csv_folder --recursive --infer-types
```

只转换为 XLSX：

```bash
python3 csv_batch_convert.py /path/to/csv_folder --format xlsx --infer-types
```

只转换为 TXT：

```bash
python3 csv_batch_convert.py /path/to/csv_folder --format txt
```

只转换为 Markdown：

```bash
python3 csv_batch_convert.py /path/to/csv_folder --format md
```

全部转换：

```bash
python3 csv_batch_convert.py /path/to/csv_folder --format all --infer-types
```

## 支持的 CSV 编码

程序默认会自动尝试以下编码：

```text
utf-8-sig
utf-8
gb18030
cp936
big5
latin-1
```

所以常见的中文 CSV 文件一般可以直接转换。

如果需要手动指定编码：

```bash
python3 csv_batch_convert.py data.csv --encoding gb18030
```

## 支持的分隔符

默认自动识别：

```text
逗号 ,
分号 ;
Tab
竖线 |
```

如果需要手动指定分隔符：

```bash
python3 csv_batch_convert.py data.csv --delimiter ","
```

Tab 分隔：

```bash
python3 csv_batch_convert.py data.csv --delimiter tab
```

## 常见问题

### 双击后提示找不到 Python

请安装 Python 3。

Windows 安装时需要勾选：

```text
Add Python to PATH
```

### macOS 双击 `.command` 没反应

可以在终端执行：

```bash
chmod +x mac一键启动.command
```

然后重新双击。

### 图形窗口打不开

程序会尝试自动退回到终端提问模式。按照终端里的提示输入文件夹或 CSV 文件路径即可。

### 转换结果在哪里

结果都在源文件或源文件夹旁边自动创建的：

```text
转换结果
```

文件夹中。

## 说明

本工具只使用 Python 标准库，不依赖 pandas、openpyxl 等第三方库。

