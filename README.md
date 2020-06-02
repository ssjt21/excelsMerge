### 使用说明

**将多个表格内容合并到同一个表格**

#### 环境依赖

- 第三方处理库安装 openpyxl
- Python2.6 及以上环境

```python
# 安装 openpyxl
pip install openpyxl==2.4.8
```

#### 目录结构介绍

```python
│  excelMerge.py    #脚本处理程序
│
├─data                  
│      title.conf       #存放导出文件表头信息的配置，字体大小，背景，颜色等
```


#### 如何使用？

- -i:指定合并的Excel文件路径
- -o:指定输出文件的名称
- -n：指定忽略合并Excel的前几行，可以通过该参数来控制表头数据写入文件，默认值为1

```python
python excelsMerge.py -i inputPath -o outputFilenaem
```

#### 如何配置导出文件表格的表头？

**修改路径下data/title.conf文件即可

```python
#demo,使用::分割每个字段的含义，列名称::字体大小::字体样式::字体颜色::背景颜色::列宽::行高
1::序号::11::bold::ffffff::538DD5::33::5
2::资产IP::11::bold::ffffff::538DD5::33::16
3::系统名称::11::bold::ffffff::538DD5::33::16
4::资产类型::11::bold::ffffff::538DD5::33::10
5::适用对象类型::11::bold::ffffff::538DD5::33::16
6::检查标准来源::11::bold::ffffff::538DD5::33::20
7::行业大类::11::bold::ffffff::538DD5::33::43
9::行业小类::11::bold::ffffff::538DD5::33::30
10::要求内容::11::bold::ffffff::538DD5::33::13
11::操作指南::11::bold::ffffff::538DD5::33::10
12::检测方法::11::bold::ffffff::538DD5::33::24
13::现网结果::11::bold::ffffff::538DD5::33::28
14::是否合规::11::bold::ffffff::538DD5::33::13
15::整改建议::11::bold::ffffff::538DD5::33::35
16::发现时间::11::bold::ffffff::538DD5::33::24
17::发现人姓名::11::bold::ffffff::538DD5::33::14
18::是否整改::11::bold::ffffff::538DD5::33::13
19::未整改原因或整改计划说明::11::bold::ffffff::538DD5::33::26

```
#### 其他问题

```shell
usage: exclesMerge.py [-h] [-n [NUM]] [-o [OUTPUT]] -i [INPUT]

Excels Merge tool!

optional arguments:
  -h, --help            show this help message and exit
  -n [NUM], --num [NUM]
                        忽略Excel文件表头的行数，默认值为1
  -o [OUTPUT], --output [OUTPUT]
                        指定导出文件的路径，默认值为 ./reports/output.xlsx
  -i [INPUT], --input [INPUT]
                        指定需要处理的多个Excel文件路径

```



