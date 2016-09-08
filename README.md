# DDoc
[查看Demo](https://distchen.github.io/DDoc/)
> DDoc 用 JS 生成一份 word 文档，在文档中可以添加一些常规的元素并给这些元素设置一些常用的属性。

## 使用方式

引用如下脚本文件:
```
<script type="text/javascript" src="jquery-3.1.0.min.js"></script>
<script type="text/javascript" src="jszip.js"></script>
<script type="text/javascript" src="DDoc.js"></script>
```

生成一份文档：

```
function generate() {
    var doc = new DDoc();
    doc.addParagraph("添加一个段落");
    doc.addParagraph("添加一个段落，设置字体和大小",{
        font:"Microsoft YaHei UI",
        fontSize:"44"
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
        color:"FF0000"
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
    ]);
    doc.newLine();
    doc.addTable([
        [1, 2, 3],
        [4, 5, 6],
        [7, 8, 9]
    ],{
        color:"FF0000"
    });
    doc.generate();
}
```

```
<a href="javascript:generate()">生成文档</a>
```
生成的word文档及内容如下：

![image](https://raw.githubusercontent.com/DistChen/DDoc/master/demo.png)