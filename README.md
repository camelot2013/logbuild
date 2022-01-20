学习使用docx和xlwt模块生成word文件和Excel文件

根据行内IT任务管理平台导出的工作日志数据分别生成周报、月报、季报 现在还有2个问题没解决： 
  * 1、word文件中添加一个表格，如果指定表格宽度，直接指定table.width = Cm(15.18)和指定每个单元格的width都不能实际生效。
  * 2、word文件中段落使用项目符号时的图标如何选择问题

注意，在使用pip安装包时，程序中用的docx实际是在python-docx中，而不是docx。所以应该是
pip install python-docx

还需要用到xlrd(用于读excel),xlwt(用于写excel)

使用PySide2设计一个简单的界面

使用pyinstaller进行打包，打包为单个文件
  * pyinstaller -w -F logbuild.py
  
  因为使用了动态加载ui文件，所以打包命令需要修改，以前的简单命令能打包成功但是无法执行。新命令如下：
  * pyinstaller -w -F --add-data="mainFace.ui;." --add-data="window.ico;." logbuild.py --hidden-import PySide2.QtXml
  
打包成功后在当前目录会有个dist目录，打包后的可执行文件在该目录内。
* 要成功执行，还需要将person_list.json拷贝到相同目录。

使用PySide6后，在用pyinstaller打包时发现一个问题，就是pyinstaller的版本低了导致打包出的程序无法运行，报错：

qt.qpa.plugin: Could not find the Qt platform plugin "windows" in "D:\SRC\python\logbuild\dist\logbuild\PySide6\plugins\platforms"

This application failed to start because no Qt platform plugin could be initialized. Reinstalling the application may fix this problem.

目前是将pyinstaller升级到最新的4.8后解决了问题。