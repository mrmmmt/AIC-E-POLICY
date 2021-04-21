# AIC-E-POLICY 电子保单下载
利用 AIC 官网下载电子保单接口，实现 AIC 电子保单的自动批量下载，并利用 Python 标准 GUI 库 Tkinter 实现界面化。

## 依赖和打包
仓库中 [env](http://www.baidu.com) 为包含所有依赖项的 Python 虚拟环境，虚拟环境 Python 版本为 3.8.3
- 虚拟环境安装 `$ pip install virtualenv`
- 虚拟环境创建 `$ python -m venv env` 环境名称为 `env`
- 虚拟环境激活 `$ env\Scripts\activate`
- 虚拟环境退出 `$ deactivate`

由于较新版本的 xlrd 库只支持 .xls 文件，不支持 .xlsx，会提示 `Excel xlsx file； not supported` 错误，故需要使用旧版本的 xlrd 库，版本为 1.2.0。

`$ pip install xlrd==1.2.0`

使用 pyinstaller 进行 Python 文件的打包，生成可执行程序

`$ pyinstaller -F -w -i icon.ico main.py`

|相关参数      |参数说明|
|----------------------------|----------------------------|
|-h, --help                  |查看该模块的帮助信息|
|-F, -onefile                |产生单个的可执行文件|
|-D, --onedir                |产生一个目录（包含多个文件）作为可执行程序|
|-a, --ascii                 |不包含 Unicode 字符集支持|
|-d, --debug                 |产生 debug 版本的可执行文件|
|-w, --windowed, --noconsolc |指定程序运行时不显示命令行窗口（仅对 Windows 有效）|
|-c, --nowindowed, --console |指定使用命令行窗口运行程序（仅对 Windows 有效）|
|-o DIR, --out=DIR           |指定 .spec 文件的生成目录。如果没有指定，则默认使用当前目录来生成 .spec 文件|
|-p DIR, --path=DIR          |设置 Python 导入模块的路径（和设置 PYTHONPATH 环境变量的作用相似）。也可使用路径分隔符（Windows 使用分号，Linux 使用冒号）来分隔多个路径|
|-n NAME, --name=NAME        |指定项目（产生的 .spec）名字。如果省略该选项，那么第一个脚本的主文件名将作为 .spec 的名字|

## 运行
$ python [main.py](https://github.com/mrmmmt/AIC-E-POLICY/blob/master/main.py) 或直接打开可执行文件 [dist/电子保单下载.exe](https://github.com/mrmmmt/AIC-E-POLICY/tree/master/dist)

## 版本

### v 1.2 (Build 20210416)
- 设置已存在保单文件大于 150000 字节时跳过下载
- 设置点击 开始下载 按钮时按钮变为不可点击状态 `self.download_button['state'] = 'disabled'`，下载结束恢复正常

### v 1.1 (Build 20210416)
- 将 xlrd 库版本降低至 `xlrd 1.2.0`，以解决导入 .xlsx 格式文件出错问题
- 通过 [mkicon.py](https://github.com/mrmmmt/AIC-E-POLICY/blob/master/mkicon.py) 生成 [icon.py](https://github.com/mrmmmt/AIC-E-POLICY/blob/master/icon.py)，以解决 pyinstaller 打包图标出错问题

### v 1.0 (Build 20210415)
- 实现多线程下载电子保单
- 实现通过直接输入或由 .txt .xls .xlsx 文件导入待下载保单号
- 实现日志的显示及导出功能
- 完成 GUI 界面编写

