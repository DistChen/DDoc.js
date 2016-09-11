# DDoc.js
[查看Demo](https://distchen.github.io/DDoc.js/)

使用 DDoc.js，你可以生成一份 word 文档，在此文档中可以添加多种元素并设置多种属性。
## 可添加的元素
> - 段落
> - 超链接
> - 标题
> - 表格
> - 列表
> - 图片

## 可设置的属性
> - font：字体，比如："Microsoft YaHei UI"
> - fontSize：字体大小，比如："44"
> - bold:true：是否加粗，true 或 false，默认不加粗
> - color：颜色，比如："FF0000"
> - highlightColor：高亮颜色，比如："blue"
> - italic：是否斜体，true 或 false，默认非斜体
> - underline: 下划线类型，比如：doc.UnderlineType.Wave(波浪线)，默认无下划线
> - underlineColor：下划线颜色，比如："FF0000"
> - strike：文本删除线，true 或 false
> - shadow：文本底纹颜色，比如："FFFFFF"，默认无
> - textAlign: 文本对齐方式，比如：doc.AlignType.Center，默认两端对齐
> - lineHeight：行间距，比如：3


## 使用方式

### 引用脚本文件
```
<script type="text/javascript" src="jquery-3.1.0.min.js"></script>
<script type="text/javascript" src="jszip.js"></script>
<script type="text/javascript" src="DDoc.js"></script>
```
### 构造 DDoc 实例
```
var doc = new DDoc();
```
### 添加段落

```
doc.addParagraph("添加一个段落");
```
### 添加段落并设置样式
```
doc.addParagraph("添加一个段落，设置字体和大小",{
    font:"Microsoft YaHei UI",
    italic:true,
    underline:doc.UnderlineType.Wave,
    underlineColor:"FF0000",
    strike:true,
    shadow:"FFFFFF",
    textAlign:doc.AlignType.Center,
    lineHeight:3
});
```
### 添加超链接
```
doc.addHyperlink("Github地址","https://github.com/DistChen/DDoc.js");
```
### 添加超链接并设置样式
```
doc.addHyperlink("Github地址","https://github.com/DistChen/DDoc.js",{
    fontSize:"30",
    bold:true
});
```
### 添加标题
```
doc.addHeader("标题1", doc.HeaderType.H1);//H2....H7
```
### 添加标题并设置样式
```
doc.addHeader("标题2", doc.HeaderType.H2,{
    font:"Microsoft YaHei UI",
    underline:doc.UnderlineType.Double,
    color:"67ff56",
    underlineColor:"FF0000"
});
```

### 添加4*5的空表格
```
doc.addEmptyTable(4, 5);
```
### 添加3*3表格(有数据)并设置颜色为红色
```
doc.addTable([
    [1, 2, 3],
    [4, 5, 6],
    [7, 8, 9]
],{color:"FF0000"});
```
### 添加表格(行列分别有合并)

```
//doc.Merge.CC 代表此单元格跨列合并
//doc.Merge.RC 代表此单元格跨行合并
doc.addTable([
    [1, 2, 4,4,5],
    [doc.Merge.CC,3, doc.Merge.RC,5,8],
    [7, 8, 4,doc.Merge.CC,9],
    [1, 2, doc.Merge.RC,doc.Merge.CC,5]
]);
```
### 添加图片

```
var _temp="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAEQAAAAmCAIAAADyTaq0AAAAAXNSR0IArs4c6QAAAAlwSFlzAAAOwwAADsMBx2+oZAAAAWFJREFUWEftVgkOhCAMVB+0L9n/f4flUGQVcIptPQIhakwpnZm2MA6f7/C8YdaQw+foHtOtgRgzZGcMOoLyHzpg7FZts0p1Io+zM0pgWuW3yZOdBX9yyqRSHAkTgsPFUwIzS+9f/3Ba1Sms2+SYt8KVsbXoA6w/ZxBLhynlyeZ/CDj8REYOCQzGtRQAyWqDRMRvgytzbu+QcsJDBUx6GmSrnAmkBYM0kdbdVASJwU0zlkq0zemxX0g8NygUOtqWNKvIE/sM2Jqsma4gsRxVaobCMN12TQAMDFJWZ45IdO1B98DA0OnKryiRwuTfg4nFwOT0KjcejHa9SoFN0qy5BUvFRvbbXDN3VBMEs6/cQBu1zZ20P9AKBENW/JIFymAE7zOWPmUwsoJtwAifasUa4wH5NmXQixEPe5Je3qaMJFe6vrsyunzju3VlcK54LWPfLZxXXRlevvm8dWX4uOT19CplfqqGCgdY+hAkAAAAAElFTkSuQmCC";
doc.addImage(_temp,100,50,{
    textAlign:doc.AlignType.Center
});
```
### 添加列表

```
doc.addList(['第一章', '第二章', '第三章'],{
    color:"FF0000"
});
```

### 添加空行
```
doc.newLine();
```
### 生成word文档

```
doc.generate();
```


## 生成一份文档的示例代码：

```
function generate() {
    var doc = new DDoc();
    doc.addParagraph("添加一个段落");
    doc.addParagraph("添加一个段落，设置字体和大小",{
        font:"Microsoft YaHei UI",
        italic:true,
        underline:doc.UnderlineType.Wave,
        underlineColor:"FF0000",
        strike:true,
        shadow:"FFFFFF",
        textAlign:doc.AlignType.Center,
        lineHeight:3
    });
    doc.addParagraph("再添加一个段落，设置一些样式",{
        font:"Microsoft YaHei UI",
        fontSize:"44",
        bold:true,
        color:"FF0000",
        highlightColor:"blue"
    });

    doc.addHeader("标题1", doc.HeaderType.H1);
    doc.addHeader("标题2", doc.HeaderType.H2,{
        font:"Microsoft YaHei UI",
        underline:doc.UnderlineType.Double,
        color:"67ff56",
        underlineColor:"FF0000"
    });

    doc.addList(['第1章', '第2章', '第3章']);
    doc.addList(['第一章', '第二章', '第三章'],{
        color:"FF0000"
    });

    doc.addEmptyTable(4, 5);
    doc.newLine();
    doc.addTable([
        [1, 2, 3],
        [4, 5, 6],
        [7, 8, 9]
    ],{color:"FF0000"});

    doc.newLine();
    doc.addTable([
        [1, 2, 4,4,5],
        [doc.Merge.CC,3, doc.Merge.RC,5,8],
        [7, 8, 4,doc.Merge.CC,9],
        [1, 2, doc.Merge.RC,doc.Merge.CC,5]
    ]);

    var _temp="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAEQAAAAmCAIAAADyTaq0AAAAAXNSR0IArs4c6QAAAAlwSFlzAAAOwwAADsMBx2+oZAAAAWFJREFUWEftVgkOhCAMVB+0L9n/f4flUGQVcIptPQIhakwpnZm2MA6f7/C8YdaQw+foHtOtgRgzZGcMOoLyHzpg7FZts0p1Io+zM0pgWuW3yZOdBX9yyqRSHAkTgsPFUwIzS+9f/3Ba1Sms2+SYt8KVsbXoA6w/ZxBLhynlyeZ/CDj8REYOCQzGtRQAyWqDRMRvgytzbu+QcsJDBUx6GmSrnAmkBYM0kdbdVASJwU0zlkq0zemxX0g8NygUOtqWNKvIE/sM2Jqsma4gsRxVaobCMN12TQAMDFJWZ45IdO1B98DA0OnKryiRwuTfg4nFwOT0KjcejHa9SoFN0qy5BUvFRvbbXDN3VBMEs6/cQBu1zZ20P9AKBENW/JIFymAE7zOWPmUwsoJtwAifasUa4wH5NmXQixEPe5Je3qaMJFe6vrsyunzju3VlcK54LWPfLZxXXRlevvm8dWX4uOT19CplfqqGCgdY+hAkAAAAAElFTkSuQmCC";
    doc.addImage(_temp,100,50,{
        textAlign:doc.AlignType.Center
    });

    doc.addHyperlink("Github地址","https://github.com/DistChen/DDoc.js",{
        fontSize:"30",
        bold:true
    });

    doc.generate();
}
```

```
<a href="javascript:generate()">生成文档</a>
```
## 生成的word文档及内容如下：

![image](https://raw.githubusercontent.com/DistChen/DDoc/master/demo.png)
