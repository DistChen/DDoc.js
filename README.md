# DDoc
[查看Demo](https://distchen.github.io/DDoc/)
> DDoc 用 JS 生成一份 word 文档，在文档中可以添加一些常规的元素：段落、标题、列表、表格等。

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
    doc.addList(['第一章', '第二章', '第三章']);
    doc.addHeader("标题1", doc.HeaderType.H1);
    doc.addList(['第1章', '第2章', '第3章']);
    doc.addHeader("标题2", doc.HeaderType.H2);
    doc.addHeader("标题3", doc.HeaderType.H3);
    doc.addTable([
        [1, 2, 3],
        [4, 5, 6],
        [7, 8, 9]
    ]);
    doc.addHeader("标题4", doc.HeaderType.H4);
    doc.addEmptyTable(4, 5);
    doc.newLine();
    doc.addParagraph("测试生成文档!");
    doc.generate();
}
```

```
<a href="javascript:generate()">生成文档</a>
```
生成的word文档及内容如下：

![image](https://raw.githubusercontent.com/DistChen/DDoc/master/demo.png)