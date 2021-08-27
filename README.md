学习使用docx和xlwt模块生成word文件和Excel文件

根据行内IT任务管理平台导出的工作日志数据分别生成周报、月报、季报 现在还有2个问题没解决： 
  * 1、word文件中添加一个表格，如果指定表格宽度，直接指定table.width = Cm(15.18)和指定每个单元格的width都不能实际生效。
  * 2、word文件中段落使用项目符号时的图标如何选择问题

注意，在使用pip安装包时，程序中用的docx实际是在python-docx中，而不是docx。所以应该是
pip install python-docx
还需要用到xlrd(用于读excel),xlwt(用于写excel)