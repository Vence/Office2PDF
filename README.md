# Office2PDF
本项目的功能是将office文件，主要是doc,docx,xls,xlsx,ppt,pptx,txt等文件轻松转成pdf文件，进而可以再前端做文件预览;

首先了解一下目前主流的文件预览方案：

##基于openoffice + swftools + flexpaper

[请访问原文博客](http://blog.csdn.net/z69183787/article/details/17468039)

主要原理

- 通过第三方工具openoffice，将word、excel、ppt、txt等文件转换为pdf文件
- 通过swfTools将pdf文件转换成swf格式的文件
- 通过FlexPaper文档组件在页面上进行展示

缺点：服务器需要安装openoffice,wsftools和flexpaper等插件，比较负重；前端需要安装flash插件

优点：实现的wsf预览方式对文件的失真较少，代码实现起来也不是特别复杂

这也是目前比较主流的方式，例如百度云盘大概就是按照这种方式

##基于Openoffice + jquery-media

主要原理：

- 通过第三方软件openoffice，将word、excel、ppt、txt等文件转换为pdf文件
- 前端使用media展示pdf文件

[jQuery-Media](http://malsup.com/jquery/media/),不仅可以无插件展示pdf文件，还可以html，
及其他媒体文件等，功能比较强大，社区提供了很多的demo，可以查看一下。

本项目代码将office转成pdf后，就可以通过jQuery-media展示pdf，具体做法就是：

`<a class="media" id="others-docpreview" href=""></a>`

js脚本执行如下：

`$("a.media").media({width:800, height:600});`

##附件：

[Openoffice附件（提取码：ceb9）](http://yunpan.cn/ccWteZCuesVMR)
