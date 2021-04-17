# AIC-E-POLICY 电子保单下载
用于批量下载 AIC 电子保单，通过 Python 标准 GUI 库 Tkinter 进行界面化处理。

## 依赖
基于 `Python 3.8.3`
- requests
- xlrd
- etc.

## 运行
$ python [main.py]([www](https://github.com/mrmmmt/AIC-E-POLICY/blob/master/main.py)) 或直接打开可执行文件 [dist/电子保单下载.exe](https://github.com/mrmmmt/AIC-E-POLICY/tree/master/dist)

## 版本
### v 1.1 (Build 20210416)
- 将 xlrd 库版本降低至 `xlrd 1.2.0`，以解决导入 .xlsx 格式文件出错问题
- 通过 [mkicon.py](https://github.com/mrmmmt/AIC-E-POLICY/blob/master/mkicon.py) 生成 [icon.py](https://github.com/mrmmmt/AIC-E-POLICY/blob/master/icon.py)，以解决 pyinstaller 打包图标出错问题

### v 1.0 (Build 20210415)
- 实现多线程下载电子保单
- 实现通过直接输入或由 `.txt .xls .xlsx` 文件导入待下载保单号
- 实现日志的显示及导出功能
- 完成 GUI 界面编写