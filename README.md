2021.7 推荐一个docx导出库，蛮好用的
仓库地址 https://github.com/egonl/SharpDocx

=========================================================================================

# EasyNPOI
word excel export

借鉴自 https://github.com/holdengong/EasyOffice 这是.net core环境的。

公司很多项目跑在.net fx4上，写了一个通用的word模板导出。

通过npoi提供的替换字符串的功能，替换word模板中的字符串达到导出效果。

主要满足了同一个表格中既有表单字段，又有嵌套的列表。

替换前：

![Image before replace](https://github.com/lc1055/EasyNPOI/blob/master/docs/before.png)

替换后：

![Image after replace](https://github.com/lc1055/EasyNPOI/blob/master/docs/after.png)


=========================================================================================

依赖 

NPOI 2.4

ICSharpCode.SharpZipLib 0.86
