# CETMarkQuery
全国大学英语四、六级考试成绩批量查询脚本(自动识别验证码)

main.py为不自动识别验证码版本，newmain.py为自动识别验证码版本(**推荐使用newmain.py**)。

在程序末尾结束时会自动计算及格率并附加相关信息在结果表格最后一行。

# 使用说明：
- 安装Python3的64位版本并安装相关依赖。
  - 依赖列表:
    - main.py:
      - urllib
      - PIL
      - tkinter
      - xlrd
      - openpyxl
    - newmain.py:
      - urllib
      - PIL
      - xlrd
      - openpyxl
      - torch(按需选择是否安装cuda版本)
      - torchvision(按需选择是否安装cuda版本)
- 下载main.py或newmain.py(区别见上文)
  - **如果使用newmain.py，需要另行下载(可在下方博客链接内下载)或训练验证码识别模块model.pth并与newmain.py放在一起才可正常使用!**
  - **model.pth为验证码识别模块！使用newmain.py时请勿删除！**
  - 博客地址: https://bytegoing.com/archives/45
- 按照下方的配置说明中正确地写入studentListExcel(studentList.xlsx)文件并与程序文件放在一起.
- 命令行下运行main.py或newmain.py即可。

# 配置说明：
- studentListExcel(studentList.xlsx)为学生信息表.该表结构如下：（第一行就应该开始放学生信息，不要有题头）
  - |准考证号|姓名|
  - 默认地址studentListExcel = 'studentList.xlsx'
- finalResult.xlsx为最终结果存储表，该表结构如下：(会自动生成题头)
  - |准考证号|姓名|查询类型|大学名称|总分|听力|阅读|写作和翻译|口试准考证号|口试等级|
  - 默认地址finalResultExcel = 'finalResult.xlsx'
  
**以下均为程序运行时可以设置的变量**
- queryYear为查询成绩的年份。如要查询2019年上半年考试此变量应为19.不要去掉引号
  - queryYear = '19'
- queryTime为查询第几次考试。如要查询2019年上半年考试此变量应为1,下半年应为2.不要去掉引号
  - queryTime = '1'
- passExamMark为及格线。大于等于此成绩为及格.
  - passExamMark = 425
- 默认查询类型 CET4/CET6
  - cxlx = 'CET4'

# 验证码识别模型说明:
自己参照 https://github.com/ice-tong/pytorch-captcha 大神的教程自行训练而成。

测试集420张验证码, 训练集1695张验证码。(新训练集和模型文件可在本人博客下载)

博客地址: https://bytegoing.com/archives/45

batch_size改为了32, 使用CUDA10.

训练到43个epoch时数据如下
```
train_loss=9.295e-06|train_acc=1.0
test_loss=7.845e-06|test_acc=0.9062
time=7.7154
```


