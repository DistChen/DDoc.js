function DDoc() {
    this.data = [];
    this.relationData = [];
    this.listCount = 0;
    this.counter = 10;
    this.zip = new JSZip("STORE");
}


DDoc.prototype._generateDocument = function () {
    var output = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 w15 wp14"><w:body>';

    for (var i in this.data) {
        output += this.data[i];
    }

    output +='<w:sectPr w:rsidR="00566DA1"><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="851" w:footer="992" w:gutter="0"/><w:cols w:space="425"/><w:docGrid w:type="lines" w:linePitch="312"/></w:sectPr></w:body></w:document>';

    return output;
};


/**
 * 根据请求参数得到样式元素
 * @param paras
 * example：
 *      {
 *          font:"Microsoft YaHei UI",
 *          fontSize:"44",
 *          bold:true,
 *          color:"FF0000",
 *          highlightColor:"blue",
 *          italic:true,
 *          underline:DDoc.prototype.UnderlineType.Single
 *          underlineColor:"FF0000",
 *          strike:true,
 *          shadow:"FFFFFF",
 *          textAlign:DDoc.prototype.AlignType.Center
 *          lineHeight:3
 *      }
 * @private
 */
DDoc.prototype._getRPrStyle=function(paras){
    var style ='<w:rPr>';
    if(paras.font){
        style +='<w:rFonts w:ascii="'+paras.font+'" w:eastAsia="'+paras.font+'" w:hAnsi="'+paras.font+'"/>';
    }
    if(paras.bold){
        style +='<w:b/>';
    }
    if(paras.fontSize){
        style +='<w:sz w:val="'+paras.fontSize+'"/><w:szCs w:val="'+paras.fontSize+'"/>';
    }
    if(paras.color){
        style += '<w:color w:val="'+paras.color+'"/>';
    }
    if(paras.highlightColor){
        style += '<w:highlight w:val="'+paras.highlightColor+'"/>';
    }
    if(paras.italic){
        style += '<w:i/>';
    }
    if(paras.underline){
        style += '<w:u w:val="'+paras.underline+'"';
        if(paras.underlineColor){
            style +=' w:color="'+paras.underlineColor+'"';
        }
        style +='/>';
    }
    if(paras.strike){
        style +='<w:strike/>'
    }
    if(paras.shadow){
        style += '<w:shd w:val="pct15" w:color="auto" w:fill="'+paras.shadow+'"/>';
    }
    style +='</w:rPr>';
    return style;
};


DDoc.prototype._getPPrStyle= function(paras){
    var style = '<w:pPr>';
    if(paras.textAlign){
        style += '<w:jc w:val="'+paras.textAlign+'"/>'
    }
    if(paras.lineHeight){
        style += '<w:spacing w:line="'+paras.lineHeight*240+'" w:lineRule="auto"/>';
    }
    style +='</w:pPr>';
    return style;
};

/**
 * 添加段落
 * @param text 被添加的文本
 * @param styles 样式
 */
DDoc.prototype.addParagraph = function (text,styles) {
    var p ='<w:p>' + this._getPPrStyle(styles||{})+
                '<w:r>' + this._getRPrStyle(styles||{})+
                    '<w:t>'+text+'</w:t>' +
                '</w:r>' +
            '</w:p>';
    this.data.push(p);
};

/**
 * 添加新行
 */
DDoc.prototype.newLine=function(){
    this.data.push('<w:p />');
};

/**
 * 添加标题
 * @param text
 * @param type DDoc.prototype.HeaderType
 * @param styles 样式
 */
DDoc.prototype.addHeader=function(text,type,styles){
    var h='<w:p>'+ this._getPPrStyle(styles||{})+
            '<w:pPr>' +
                '<w:pStyle w:val="'+type+'"/>' +
            '</w:pPr>' +
            '<w:r>' + this._getRPrStyle(styles||{})+
                '<w:t>'+text+'</w:t>' +
            '</w:r>' +
          '</w:p>';
    this.data.push(h);
};

/**
 *  添加一个空表格
 * @param row
 * @param col
 * @param styles 样式
 */
DDoc.prototype.addEmptyTable=function(row,col,styles){
    var rowArr=[];
    for(var i=0;i<row;i++){
        var colArr=[];
        for(var j=0;j<col;j++){
            colArr.push('');
        }
        rowArr.push(colArr);
    }
    this.addTable(rowArr,styles);
};

/**
 * 添加表格
 * @param arrs
 * @param styles 样式
 */
DDoc.prototype.addTable=function(arrs,styles){
    var _style = this._getRPrStyle(styles||{});
    var row = arrs.length;
    var col = arrs[0].length;
    var width = parseInt(8296/col);
    var table = '<w:tbl>' +
                    '<w:tblPr>' +
                        '<w:tblStyle w:val="a3"/>' +
                        '<w:tblW w:w="0" w:type="auto"/>' +
                    '</w:tblPr>' +
                '<w:tblGrid>';
    for(var i=0;i<col;i++){
        table +='<w:gridCol w:w="'+width+'"/>';
    }
    table +='</w:tblGrid>';
    for(var i=0;i<row;i++){
        table += this._addRow(width,_style,arrs[i]);
    }
    table+='</w:tbl>';
    this.data.push(table);
};

/**
 * 添加表格行
 * @param width 行中每个单元格的宽度
 * @param style
 * @param cols
 * @returns {string}
 * @private
 */
DDoc.prototype._addRow=function(width,style,cols){
    var row ='<w:tr>';
    var gridSpanning = false;
    var gridSpanCount = 1;
    for(var i=0;i<cols.length;i++){
        if(cols[i] === this.Merge.CC){
            gridSpanCount++;
            if(!gridSpanning){
                gridSpanning = true;
            }
            continue;
        }
        row +=  '<w:tc>' +
                    '<w:tcPr>' +
                        '<w:tcW w:w="'+gridSpanCount*width+'"/>' +
                        '<w:gridSpan w:val="'+gridSpanCount+'"/>' +
                        (cols[i]===this.Merge.RC?'<w:vMerge/>':'<w:vMerge w:val="restart"/>')+
                    '</w:tcPr>' +
                    '<w:p>' +
                        '<w:r>' + style +
                            '<w:t>'+cols[i]+'</w:t>' +
                        '</w:r>' +
                    '</w:p>' +
                '</w:tc>';
        gridSpanning = false;
        gridSpanCount = 1;
    }
    row +='</w:tr>';
    return row;
};

/**
 * 添加列表
 * @param items
 * @param styles 样式
 */
DDoc.prototype.addList=function(items,styles){
    var _style = this._getRPrStyle(styles||{});
    this.listCount++;
    var list='';
    for(var i in items){
        list +='<w:p>' +
                    '<w:pPr>' +
                        '<w:pStyle w:val="a5"/> ' +
                        '<w:numPr>' +
                            '<w:ilvl w:val="0"/> ' +
                            '<w:numId w:val="'+this.listCount+'"/> ' +
                        '</w:numPr> ' +
                        '<w:ind w:firstLineChars="0"/> ' +
                    '</w:pPr> ' +
                    '<w:r>' + _style+
                        '<w:t>'+items[i]+'</w:t> ' +
                    '</w:r> ' +
                '</w:p>';
    }
    this.data.push(list);
};

/**
 * 添加图片
 * @param data base64
 * @param width
 * @param height
 * @param style
 */
DDoc.prototype.addImage=function(data,width,height,styles){
    this.counter++;
    var imageName = "media/image"+this.counter+"."+data.substring(11,data.indexOf(";"));
    this.relationData.push('<Relationship Id="rId'+this.counter+'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="'+imageName+'"/>');
    this.zip.add("word/"+imageName,data.substring(data.indexOf(",")+1),{base64: true});

    var p = '<w:p>' + this._getPPrStyle(styles||{})+
                '<w:r>' +
                    '<w:drawing>' +
                        '<wp:inline>' +
                            '<wp:extent cx="'+width*9525+'" cy="'+height*9525+'"/> ' +
                            '<wp:docPr id="1" name=""/> ' +
                            '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"> ' +
                                '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"> ' +
                                    '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"> ' +
                                        '<pic:nvPicPr> ' +
                                            '<pic:cNvPr id="1" name=""/> ' +
                                            '<pic:cNvPicPr/> ' +
                                        '</pic:nvPicPr> ' +
                                        '<pic:blipFill> ' +
                                            '<a:blip r:embed="rId'+this.counter+'"/> ' +
                                            '<a:stretch> ' +
                                                '<a:fillRect/> ' +
                                            '</a:stretch> ' +
                                        '</pic:blipFill> ' +
                                        '<pic:spPr> ' +
                                            '<a:xfrm> ' +
                                                '<a:off x="0" y="0"/> ' +
                                                '<a:ext cx="'+width*9525+'" cy="'+height*9525+'"/> ' +
                                            '</a:xfrm> ' +
                                            '<a:prstGeom prst="rect"> ' +
                                                '<a:avLst/> ' +
                                            '</a:prstGeom> ' +
                                        '</pic:spPr> ' +
                                    '</pic:pic> ' +
                                '</a:graphicData> ' +
                            '</a:graphic> ' +
                        '</wp:inline> ' +
                    '</w:drawing> ' +
                '</w:r> ' +
            '</w:p>';
    this.data.push(p);
};

/**
 * 添加超链接
 * @param displayName 显示名称
 * @param url
 * @param style
 */
DDoc.prototype.addHyperlink=function(displayName,url,style){
    this.counter++;
    this.relationData.push('<Relationship Id="rId'+this.counter+'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="'+url+'" TargetMode="External"/>');
    var p ='<w:p> ' +
            '<w:hyperlink r:id="rId'+this.counter+'"> ' +
                '<w:r> ' + this._getRPrStyle(style)+
                    '<w:rPr> ' +
                        '<w:rStyle w:val="a3"/> ' +
                    '</w:rPr> ' +
                   '<w:t>'+displayName+'</w:t>' +
                   '</w:r> ' +
                '</w:hyperlink> ' +
            '</w:p>';
    this.data.push(p);
};

/**
 * 生成文档并下载
 */
DDoc.prototype.generate = function () {
    for (var i in this.Templates) {
        if(this.Templates[i].name == "word/_rels/document.xml.rels"){
            var temp = this.Templates[i].value.substring(0,this.Templates[i].value.length - 16);
            this.relationData.forEach(function(item){
                temp += item;
            });
            temp += '</Relationships>';
            this.zip.add("word/_rels/document.xml.rels", temp);
        }else{
            this.zip.add(this.Templates[i].name, this.Templates[i].value);
        }
    }
    this.zip.add("word/document.xml", this._generateDocument());
    document.location.href = 'data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,' + this.zip.generate();
};


DDoc.prototype.AlignType={
    Left:"left",
    Center:"center",
    Right:"right",
    Distribute:"distribute" //分散对齐
};


DDoc.prototype.Merge={
    RC:{},//RowCell: 跨行合并的单元格，这些单元格都在同一列
    CC:{}//ColumnCell:跨列合并的单元格，这些单元格在同一行
};

DDoc.prototype.HeaderType={
    H1:"1",
    H2:"2",
    H3:"3",
    H4:"4",
    H5:"5",
    H6:"6",
    H7:"7"
};

DDoc.prototype.UnderlineType={
    Single:"single",
    Double:"double",
    Thick:"thick",
    Dotted:"dotted",
    Dash:"dash",
    DotDash:"dotDash",
    DotDotDash:"dotDotDash",
    Wave:"wave"
};

DDoc.prototype.Templates = [
    {
        "name": "[Content_Types].xml",
        "value": "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"png\" ContentType=\"image/png\"/><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/><Override PartName=\"/word/numbering.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml\"/><Override PartName=\"/word/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml\"/><Override PartName=\"/word/settings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml\"/><Override PartName=\"/word/webSettings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml\"/><Override PartName=\"/word/fontTable.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml\"/><Override PartName=\"/word/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/><Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/><Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/></Types>"
    }, {
        "name": "_rels/.rels",
        "value": "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/app.xml\"/><Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/></Relationships>"
    }, {
        "name": "docProps/app.xml",
        "value": "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\"><Template>Normal.dotm</Template><TotalTime>2</TotalTime><Pages>1</Pages><Words>10</Words><Characters>61</Characters><Application>Microsoft Office Word</Application><DocSecurity>0</DocSecurity><Lines>1</Lines><Paragraphs>1</Paragraphs><ScaleCrop>false</ScaleCrop><Company></Company><LinksUpToDate>false</LinksUpToDate><CharactersWithSpaces>70</CharactersWithSpaces><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>15.0000</AppVersion></Properties>"
    }, {
        "name": "docProps/core.xml",
        "value": "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"><dc:title></dc:title><dc:subject></dc:subject><dc:creator>chenyp</dc:creator><cp:keywords></cp:keywords><dc:description></dc:description><cp:lastModifiedBy>chenyp</cp:lastModifiedBy><cp:revision>9</cp:revision><dcterms:created xsi:type=\"dcterms:W3CDTF\">2016-09-09T05:41:00Z</dcterms:created><dcterms:modified xsi:type=\"dcterms:W3CDTF\">2016-09-09T05:43:00Z</dcterms:modified></cp:coreProperties>"
    }, {
        "name": "word/_rels/document.xml.rels",
        "value": "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n<Relationships \r\n    xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\r\n    <Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings\" Target=\"settings.xml\"/>\r\n    <Relationship Id=\"rId7\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"theme/theme1.xml\"/>\r\n    <Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>\r\n    <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering\" Target=\"numbering.xml\"/>\r\n    <Relationship Id=\"rId6\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable\" Target=\"fontTable.xml\"/>\r\n    <Relationship Id=\"rId4\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings\" Target=\"webSettings.xml\"/>\r\n</Relationships>"
    }, {
        "name": "word/fontTable.xml",
        "value": "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n<w:fonts xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:w15=\"http://schemas.microsoft.com/office/word/2012/wordml\" mc:Ignorable=\"w14 w15\"><w:font w:name=\"Times New Roman\"><w:panose1 w:val=\"02020603050405020304\"/><w:charset w:val=\"00\"/><w:family w:val=\"roman\"/><w:pitch w:val=\"variable\"/><w:sig w:usb0=\"E0002EFF\" w:usb1=\"C0007843\" w:usb2=\"00000009\" w:usb3=\"00000000\" w:csb0=\"000001FF\" w:csb1=\"00000000\"/></w:font><w:font w:name=\"Wingdings\"><w:panose1 w:val=\"05000000000000000000\"/><w:charset w:val=\"02\"/><w:family w:val=\"auto\"/><w:pitch w:val=\"variable\"/><w:sig w:usb0=\"00000000\" w:usb1=\"10000000\" w:usb2=\"00000000\" w:usb3=\"00000000\" w:csb0=\"80000000\" w:csb1=\"00000000\"/></w:font><w:font w:name=\"Calibri\"><w:panose1 w:val=\"020F0502020204030204\"/><w:charset w:val=\"00\"/><w:family w:val=\"swiss\"/><w:pitch w:val=\"variable\"/><w:sig w:usb0=\"E00002FF\" w:usb1=\"4000ACFF\" w:usb2=\"00000001\" w:usb3=\"00000000\" w:csb0=\"0000019F\" w:csb1=\"00000000\"/></w:font><w:font w:name=\"宋体\"><w:altName w:val=\"SimSun\"/><w:panose1 w:val=\"02010600030101010101\"/><w:charset w:val=\"86\"/><w:family w:val=\"auto\"/><w:pitch w:val=\"variable\"/><w:sig w:usb0=\"00000003\" w:usb1=\"288F0000\" w:usb2=\"00000016\" w:usb3=\"00000000\" w:csb0=\"00040001\" w:csb1=\"00000000\"/></w:font><w:font w:name=\"Calibri Light\"><w:panose1 w:val=\"020F0302020204030204\"/><w:charset w:val=\"00\"/><w:family w:val=\"swiss\"/><w:pitch w:val=\"variable\"/><w:sig w:usb0=\"A00002EF\" w:usb1=\"4000207B\" w:usb2=\"00000000\" w:usb3=\"00000000\" w:csb0=\"0000019F\" w:csb1=\"00000000\"/></w:font></w:fonts>"
    }, {
        "name": "word/settings.xml",
        "value": "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n<w:settings xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:w10=\"urn:schemas-microsoft-com:office:word\" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:w15=\"http://schemas.microsoft.com/office/word/2012/wordml\" xmlns:sl=\"http://schemas.openxmlformats.org/schemaLibrary/2006/main\" mc:Ignorable=\"w14 w15\"><w:zoom w:percent=\"100\"/><w:bordersDoNotSurroundHeader/><w:bordersDoNotSurroundFooter/><w:proofState w:spelling=\"clean\" w:grammar=\"clean\"/><w:defaultTabStop w:val=\"420\"/><w:drawingGridVerticalSpacing w:val=\"156\"/><w:displayHorizontalDrawingGridEvery w:val=\"0\"/><w:displayVerticalDrawingGridEvery w:val=\"2\"/><w:characterSpacingControl w:val=\"compressPunctuation\"/><w:compat><w:spaceForUL/><w:balanceSingleByteDoubleByteWidth/><w:doNotLeaveBackslashAlone/><w:ulTrailSpace/><w:doNotExpandShiftReturn/><w:adjustLineHeightInTable/><w:useFELayout/><w:compatSetting w:name=\"compatibilityMode\" w:uri=\"http://schemas.microsoft.com/office/word\" w:val=\"15\"/><w:compatSetting w:name=\"overrideTableStyleFontSizeAndJustification\" w:uri=\"http://schemas.microsoft.com/office/word\" w:val=\"1\"/><w:compatSetting w:name=\"enableOpenTypeFeatures\" w:uri=\"http://schemas.microsoft.com/office/word\" w:val=\"1\"/><w:compatSetting w:name=\"doNotFlipMirrorIndents\" w:uri=\"http://schemas.microsoft.com/office/word\" w:val=\"1\"/><w:compatSetting w:name=\"differentiateMultirowTableHeaders\" w:uri=\"http://schemas.microsoft.com/office/word\" w:val=\"1\"/></w:compat><w:rsids><w:rsidRoot w:val=\"00AB5A46\"/><w:rsid w:val=\"00000A28\"/><w:rsid w:val=\"00001135\"/><w:rsid w:val=\"00004260\"/><w:rsid w:val=\"00012351\"/><w:rsid w:val=\"00017D39\"/><w:rsid w:val=\"00021DAC\"/><w:rsid w:val=\"00030AC3\"/><w:rsid w:val=\"00033F77\"/><w:rsid w:val=\"00034BD8\"/><w:rsid w:val=\"00051FB7\"/><w:rsid w:val=\"00052734\"/><w:rsid w:val=\"00060AF1\"/><w:rsid w:val=\"000733FF\"/><w:rsid w:val=\"00087216\"/><w:rsid w:val=\"00090D61\"/><w:rsid w:val=\"000A6950\"/><w:rsid w:val=\"000B68F8\"/><w:rsid w:val=\"000C0A9B\"/><w:rsid w:val=\"000D07B0\"/><w:rsid w:val=\"000D460C\"/><w:rsid w:val=\"000D4B8D\"/><w:rsid w:val=\"000D53A1\"/><w:rsid w:val=\"00113719\"/><w:rsid w:val=\"001157D5\"/><w:rsid w:val=\"001318BB\"/><w:rsid w:val=\"0014096F\"/><w:rsid w:val=\"001475D5\"/><w:rsid w:val=\"001759E4\"/><w:rsid w:val=\"0017713D\"/><w:rsid w:val=\"001A329E\"/><w:rsid w:val=\"001A5B86\"/><w:rsid w:val=\"001A6C01\"/><w:rsid w:val=\"001B4DB3\"/><w:rsid w:val=\"001C2F12\"/><w:rsid w:val=\"001C5C42\"/><w:rsid w:val=\"001E0CD0\"/><w:rsid w:val=\"001F2129\"/><w:rsid w:val=\"001F4C8F\"/><w:rsid w:val=\"00217F04\"/><w:rsid w:val=\"00220A9B\"/><w:rsid w:val=\"002332C7\"/><w:rsid w:val=\"00240064\"/><w:rsid w:val=\"002524C2\"/><w:rsid w:val=\"002530CA\"/><w:rsid w:val=\"00255618\"/><w:rsid w:val=\"002668D6\"/><w:rsid w:val=\"00266BFF\"/><w:rsid w:val=\"002763F7\"/><w:rsid w:val=\"002A6D3A\"/><w:rsid w:val=\"002B3FC1\"/><w:rsid w:val=\"002B410D\"/><w:rsid w:val=\"002B5421\"/><w:rsid w:val=\"002B657F\"/><w:rsid w:val=\"002B7015\"/><w:rsid w:val=\"002E5C41\"/><w:rsid w:val=\"002F03E5\"/><w:rsid w:val=\"002F4D1B\"/><w:rsid w:val=\"00323EE9\"/><w:rsid w:val=\"00345748\"/><w:rsid w:val=\"00354962\"/><w:rsid w:val=\"00354B08\"/><w:rsid w:val=\"00357DC3\"/><w:rsid w:val=\"00366111\"/><w:rsid w:val=\"00380E61\"/><w:rsid w:val=\"0038326B\"/><w:rsid w:val=\"00391FE7\"/><w:rsid w:val=\"00392DFE\"/><w:rsid w:val=\"00397889\"/><w:rsid w:val=\"00397C95\"/><w:rsid w:val=\"003B5835\"/><w:rsid w:val=\"003C0DAF\"/><w:rsid w:val=\"003F4923\"/><w:rsid w:val=\"003F6D33\"/><w:rsid w:val=\"00401800\"/><w:rsid w:val=\"00404567\"/><w:rsid w:val=\"004100DC\"/><w:rsid w:val=\"0041086E\"/><w:rsid w:val=\"004133BC\"/><w:rsid w:val=\"00413653\"/><w:rsid w:val=\"00427107\"/><w:rsid w:val=\"0043065E\"/><w:rsid w:val=\"004340DB\"/><w:rsid w:val=\"0044367D\"/><w:rsid w:val=\"00457755\"/><w:rsid w:val=\"00460C0B\"/><w:rsid w:val=\"0046753A\"/><w:rsid w:val=\"00471E51\"/><w:rsid w:val=\"00484AA8\"/><w:rsid w:val=\"004A1603\"/><w:rsid w:val=\"004C17CE\"/><w:rsid w:val=\"004C374C\"/><w:rsid w:val=\"004D1698\"/><w:rsid w:val=\"004D5369\"/><w:rsid w:val=\"004E0774\"/><w:rsid w:val=\"004E66ED\"/><w:rsid w:val=\"004E77FA\"/><w:rsid w:val=\"00514D7B\"/><w:rsid w:val=\"00517EBF\"/><w:rsid w:val=\"00531EB4\"/><w:rsid w:val=\"00547C7B\"/><w:rsid w:val=\"005746F9\"/><w:rsid w:val=\"00586EF4\"/><w:rsid w:val=\"005A0BEB\"/><w:rsid w:val=\"005A16BA\"/><w:rsid w:val=\"005A17FE\"/><w:rsid w:val=\"005A6247\"/><w:rsid w:val=\"005B38CB\"/><w:rsid w:val=\"005C15E7\"/><w:rsid w:val=\"005E2C6C\"/><w:rsid w:val=\"005F5061\"/><w:rsid w:val=\"00602A69\"/><w:rsid w:val=\"00602E65\"/><w:rsid w:val=\"00602F28\"/><w:rsid w:val=\"006031B6\"/><w:rsid w:val=\"00606B06\"/><w:rsid w:val=\"00612A5C\"/><w:rsid w:val=\"00612E23\"/><w:rsid w:val=\"00617E3B\"/><w:rsid w:val=\"006363A1\"/><w:rsid w:val=\"00636FE2\"/><w:rsid w:val=\"00645FA3\"/><w:rsid w:val=\"006857EB\"/><w:rsid w:val=\"00686B4C\"/><w:rsid w:val=\"0069074B\"/><w:rsid w:val=\"00690D21\"/><w:rsid w:val=\"00696B65\"/><w:rsid w:val=\"006A2977\"/><w:rsid w:val=\"006A6DA8\"/><w:rsid w:val=\"006B69FA\"/><w:rsid w:val=\"006C37FC\"/><w:rsid w:val=\"006D2EEA\"/><w:rsid w:val=\"006E5832\"/><w:rsid w:val=\"006F659C\"/><w:rsid w:val=\"00721979\"/><w:rsid w:val=\"00722F5A\"/><w:rsid w:val=\"007245C7\"/><w:rsid w:val=\"00736D1F\"/><w:rsid w:val=\"0074617F\"/><w:rsid w:val=\"00763790\"/><w:rsid w:val=\"0077460F\"/><w:rsid w:val=\"00780A6D\"/><w:rsid w:val=\"00796F89\"/><w:rsid w:val=\"007A04F3\"/><w:rsid w:val=\"007B23A7\"/><w:rsid w:val=\"007B3E08\"/><w:rsid w:val=\"007C76FF\"/><w:rsid w:val=\"007E22C0\"/><w:rsid w:val=\"007E3057\"/><w:rsid w:val=\"007E54FF\"/><w:rsid w:val=\"007F0DB0\"/><w:rsid w:val=\"00800BB7\"/><w:rsid w:val=\"00815C04\"/><w:rsid w:val=\"00824B10\"/><w:rsid w:val=\"00826637\"/><w:rsid w:val=\"00841596\"/><w:rsid w:val=\"0084505A\"/><w:rsid w:val=\"008529CB\"/><w:rsid w:val=\"0087334F\"/><w:rsid w:val=\"00875DA6\"/><w:rsid w:val=\"008778C1\"/><w:rsid w:val=\"00883D2B\"/><w:rsid w:val=\"00892F43\"/><w:rsid w:val=\"008A2F22\"/><w:rsid w:val=\"008B4891\"/><w:rsid w:val=\"008B7C49\"/><w:rsid w:val=\"008C309B\"/><w:rsid w:val=\"008C67DF\"/><w:rsid w:val=\"008D7552\"/><w:rsid w:val=\"008E7C1C\"/><w:rsid w:val=\"008F2F78\"/><w:rsid w:val=\"008F4E91\"/><w:rsid w:val=\"008F5B8B\"/><w:rsid w:val=\"009071E7\"/><w:rsid w:val=\"0091606D\"/><w:rsid w:val=\"00921C8C\"/><w:rsid w:val=\"00925F2E\"/><w:rsid w:val=\"009260B0\"/><w:rsid w:val=\"009419AF\"/><w:rsid w:val=\"00947012\"/><w:rsid w:val=\"009663FC\"/><w:rsid w:val=\"00966C1D\"/><w:rsid w:val=\"00974561\"/><w:rsid w:val=\"00982532\"/><w:rsid w:val=\"009845E6\"/><w:rsid w:val=\"009A45EF\"/><w:rsid w:val=\"009B13A0\"/><w:rsid w:val=\"009B6E57\"/><w:rsid w:val=\"009C5F68\"/><w:rsid w:val=\"009C7595\"/><w:rsid w:val=\"009D616B\"/><w:rsid w:val=\"009E2B25\"/><w:rsid w:val=\"009F0FFC\"/><w:rsid w:val=\"009F4A35\"/><w:rsid w:val=\"00A03888\"/><w:rsid w:val=\"00A125DE\"/><w:rsid w:val=\"00A221A6\"/><w:rsid w:val=\"00A229DC\"/><w:rsid w:val=\"00A264D3\"/><w:rsid w:val=\"00A359A9\"/><w:rsid w:val=\"00A370CB\"/><w:rsid w:val=\"00A50AA5\"/><w:rsid w:val=\"00A52E93\"/><w:rsid w:val=\"00A5303E\"/><w:rsid w:val=\"00A63312\"/><w:rsid w:val=\"00A65B99\"/><w:rsid w:val=\"00A67363\"/><w:rsid w:val=\"00A70E5D\"/><w:rsid w:val=\"00A710E5\"/><w:rsid w:val=\"00A735CC\"/><w:rsid w:val=\"00A74A30\"/><w:rsid w:val=\"00A758B2\"/><w:rsid w:val=\"00A80208\"/><w:rsid w:val=\"00A84CCD\"/><w:rsid w:val=\"00AA056C\"/><w:rsid w:val=\"00AA6EAD\"/><w:rsid w:val=\"00AB3962\"/><w:rsid w:val=\"00AB5A46\"/><w:rsid w:val=\"00AB66BA\"/><w:rsid w:val=\"00AD0A91\"/><w:rsid w:val=\"00AD3A98\"/><w:rsid w:val=\"00AE2FFD\"/><w:rsid w:val=\"00AE34D1\"/><w:rsid w:val=\"00B06EAB\"/><w:rsid w:val=\"00B16967\"/><w:rsid w:val=\"00B2578F\"/><w:rsid w:val=\"00B50842\"/><w:rsid w:val=\"00B53856\"/><w:rsid w:val=\"00B7017A\"/><w:rsid w:val=\"00B701CF\"/><w:rsid w:val=\"00B81C63\"/><w:rsid w:val=\"00B85C41\"/><w:rsid w:val=\"00B87014\"/><w:rsid w:val=\"00B91431\"/><w:rsid w:val=\"00BA75BC\"/><w:rsid w:val=\"00BB469F\"/><w:rsid w:val=\"00BC49A4\"/><w:rsid w:val=\"00BD2EFE\"/><w:rsid w:val=\"00BE281C\"/><w:rsid w:val=\"00BF7D52\"/><w:rsid w:val=\"00C03C76\"/><w:rsid w:val=\"00C04691\"/><w:rsid w:val=\"00C11995\"/><w:rsid w:val=\"00C1263B\"/><w:rsid w:val=\"00C23596\"/><w:rsid w:val=\"00C23B8A\"/><w:rsid w:val=\"00C25207\"/><w:rsid w:val=\"00C313BB\"/><w:rsid w:val=\"00C3403A\"/><w:rsid w:val=\"00C346C0\"/><w:rsid w:val=\"00C46C3A\"/><w:rsid w:val=\"00C5337A\"/><w:rsid w:val=\"00C64C2F\"/><w:rsid w:val=\"00C81793\"/><w:rsid w:val=\"00C95438\"/><w:rsid w:val=\"00CB51F0\"/><w:rsid w:val=\"00CB5EC0\"/><w:rsid w:val=\"00CD2137\"/><w:rsid w:val=\"00D03D69\"/><w:rsid w:val=\"00D21947\"/><w:rsid w:val=\"00D24153\"/><w:rsid w:val=\"00D273C1\"/><w:rsid w:val=\"00D36EFC\"/><w:rsid w:val=\"00D44C34\"/><w:rsid w:val=\"00D475DD\"/><w:rsid w:val=\"00D51C90\"/><w:rsid w:val=\"00D51F60\"/><w:rsid w:val=\"00D5775A\"/><w:rsid w:val=\"00D749C1\"/><w:rsid w:val=\"00DA750C\"/><w:rsid w:val=\"00DB0E2B\"/><w:rsid w:val=\"00DB3D24\"/><w:rsid w:val=\"00DC4821\"/><w:rsid w:val=\"00DD65A5\"/><w:rsid w:val=\"00DE33B5\"/><w:rsid w:val=\"00DF50C9\"/><w:rsid w:val=\"00E14812\"/><w:rsid w:val=\"00E150FA\"/><w:rsid w:val=\"00E1705D\"/><w:rsid w:val=\"00E21804\"/><w:rsid w:val=\"00E452CA\"/><w:rsid w:val=\"00E57933\"/><w:rsid w:val=\"00E63F41\"/><w:rsid w:val=\"00E7130D\"/><w:rsid w:val=\"00E81997\"/><w:rsid w:val=\"00E90749\"/><w:rsid w:val=\"00EA0F04\"/><w:rsid w:val=\"00EC0E16\"/><w:rsid w:val=\"00ED1F4F\"/><w:rsid w:val=\"00ED338B\"/><w:rsid w:val=\"00EE6638\"/><w:rsid w:val=\"00F101FE\"/><w:rsid w:val=\"00F15304\"/><w:rsid w:val=\"00F1624F\"/><w:rsid w:val=\"00F163A4\"/><w:rsid w:val=\"00F201F1\"/><w:rsid w:val=\"00F46D6A\"/><w:rsid w:val=\"00F53A01\"/><w:rsid w:val=\"00F61A67\"/><w:rsid w:val=\"00F812C8\"/><w:rsid w:val=\"00F81424\"/><w:rsid w:val=\"00F96C92\"/><w:rsid w:val=\"00F97263\"/><w:rsid w:val=\"00FA3082\"/><w:rsid w:val=\"00FA392E\"/><w:rsid w:val=\"00FB6508\"/><w:rsid w:val=\"00FC14E8\"/><w:rsid w:val=\"00FC154E\"/><w:rsid w:val=\"00FC4DF1\"/><w:rsid w:val=\"00FD274A\"/><w:rsid w:val=\"00FE419A\"/><w:rsid w:val=\"00FF181A\"/><w:rsid w:val=\"00FF42A7\"/></w:rsids><m:mathPr><m:mathFont m:val=\"Cambria Math\"/><m:brkBin m:val=\"before\"/><m:brkBinSub m:val=\"--\"/><m:smallFrac m:val=\"0\"/><m:dispDef/><m:lMargin m:val=\"0\"/><m:rMargin m:val=\"0\"/><m:defJc m:val=\"centerGroup\"/><m:wrapIndent m:val=\"1440\"/><m:intLim m:val=\"subSup\"/><m:naryLim m:val=\"undOvr\"/></m:mathPr><w:themeFontLang w:val=\"en-US\" w:eastAsia=\"zh-CN\"/><w:clrSchemeMapping w:bg1=\"light1\" w:t1=\"dark1\" w:bg2=\"light2\" w:t2=\"dark2\" w:accent1=\"accent1\" w:accent2=\"accent2\" w:accent3=\"accent3\" w:accent4=\"accent4\" w:accent5=\"accent5\" w:accent6=\"accent6\" w:hyperlink=\"hyperlink\" w:followedHyperlink=\"followedHyperlink\"/><w:shapeDefaults><o:shapedefaults v:ext=\"edit\" spidmax=\"1026\"/><o:shapelayout v:ext=\"edit\"><o:idmap v:ext=\"edit\" data=\"1\"/></o:shapelayout></w:shapeDefaults><w:decimalSymbol w:val=\".\"/><w:listSeparator w:val=\",\"/><w15:chartTrackingRefBased/><w15:docId w15:val=\"{8CACAFF1-A9DB-4363-93E0-3BEEA35DBDF6}\"/></w:settings>"
    }, {
        "name": "word/styles.xml",
        "value": "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n<w:styles xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:w15=\"http://schemas.microsoft.com/office/word/2012/wordml\" mc:Ignorable=\"w14 w15\"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:asciiTheme=\"minorHAnsi\" w:eastAsiaTheme=\"minorEastAsia\" w:hAnsiTheme=\"minorHAnsi\" w:cstheme=\"minorBidi\"/><w:kern w:val=\"2\"/><w:sz w:val=\"21\"/><w:szCs w:val=\"22\"/><w:lang w:val=\"en-US\" w:eastAsia=\"zh-CN\" w:bidi=\"ar-SA\"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults><w:latentStyles w:defLockedState=\"0\" w:defUIPriority=\"99\" w:defSemiHidden=\"0\" w:defUnhideWhenUsed=\"0\" w:defQFormat=\"0\" w:count=\"371\"><w:lsdException w:name=\"Normal\" w:uiPriority=\"0\" w:qFormat=\"1\"/><w:lsdException w:name=\"heading 1\" w:uiPriority=\"9\" w:qFormat=\"1\"/><w:lsdException w:name=\"heading 2\" w:semiHidden=\"1\" w:uiPriority=\"9\" w:unhideWhenUsed=\"1\" w:qFormat=\"1\"/><w:lsdException w:name=\"heading 3\" w:semiHidden=\"1\" w:uiPriority=\"9\" w:unhideWhenUsed=\"1\" w:qFormat=\"1\"/><w:lsdException w:name=\"heading 4\" w:semiHidden=\"1\" w:uiPriority=\"9\" w:unhideWhenUsed=\"1\" w:qFormat=\"1\"/><w:lsdException w:name=\"heading 5\" w:semiHidden=\"1\" w:uiPriority=\"9\" w:unhideWhenUsed=\"1\" w:qFormat=\"1\"/><w:lsdException w:name=\"heading 6\" w:semiHidden=\"1\" w:uiPriority=\"9\" w:unhideWhenUsed=\"1\" w:qFormat=\"1\"/><w:lsdException w:name=\"heading 7\" w:semiHidden=\"1\" w:uiPriority=\"9\" w:unhideWhenUsed=\"1\" w:qFormat=\"1\"/><w:lsdException w:name=\"heading 8\" w:semiHidden=\"1\" w:uiPriority=\"9\" w:unhideWhenUsed=\"1\" w:qFormat=\"1\"/><w:lsdException w:name=\"heading 9\" w:semiHidden=\"1\" w:uiPriority=\"9\" w:unhideWhenUsed=\"1\" w:qFormat=\"1\"/><w:lsdException w:name=\"index 1\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"index 2\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"index 3\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"index 4\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"index 5\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"index 6\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"index 7\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"index 8\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"index 9\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"toc 1\" w:semiHidden=\"1\" w:uiPriority=\"39\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"toc 2\" w:semiHidden=\"1\" w:uiPriority=\"39\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"toc 3\" w:semiHidden=\"1\" w:uiPriority=\"39\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"toc 4\" w:semiHidden=\"1\" w:uiPriority=\"39\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"toc 5\" w:semiHidden=\"1\" w:uiPriority=\"39\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"toc 6\" w:semiHidden=\"1\" w:uiPriority=\"39\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"toc 7\" w:semiHidden=\"1\" w:uiPriority=\"39\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"toc 8\" w:semiHidden=\"1\" w:uiPriority=\"39\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"toc 9\" w:semiHidden=\"1\" w:uiPriority=\"39\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Normal Indent\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"footnote text\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"annotation text\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"header\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"footer\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"index heading\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"caption\" w:semiHidden=\"1\" w:uiPriority=\"35\" w:unhideWhenUsed=\"1\" w:qFormat=\"1\"/><w:lsdException w:name=\"table of figures\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"envelope address\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"envelope return\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"footnote reference\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"annotation reference\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"line number\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"page number\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"endnote reference\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"endnote text\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"table of authorities\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"macro\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"toa heading\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"List\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"List Bullet\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"List Number\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"List 2\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"List 3\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"List 4\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"List 5\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"List Bullet 2\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"List Bullet 3\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"List Bullet 4\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"List Bullet 5\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"List Number 2\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"List Number 3\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"List Number 4\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"List Number 5\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Title\" w:uiPriority=\"10\" w:qFormat=\"1\"/><w:lsdException w:name=\"Closing\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Signature\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Default Paragraph Font\" w:semiHidden=\"1\" w:uiPriority=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Body Text\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Body Text Indent\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"List Continue\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"List Continue 2\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"List Continue 3\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"List Continue 4\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"List Continue 5\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Message Header\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Subtitle\" w:uiPriority=\"11\" w:qFormat=\"1\"/><w:lsdException w:name=\"Salutation\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Date\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Body Text First Indent\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Body Text First Indent 2\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Note Heading\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Body Text 2\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Body Text 3\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Body Text Indent 2\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Body Text Indent 3\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Block Text\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Hyperlink\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"FollowedHyperlink\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Strong\" w:uiPriority=\"22\" w:qFormat=\"1\"/><w:lsdException w:name=\"Emphasis\" w:uiPriority=\"20\" w:qFormat=\"1\"/><w:lsdException w:name=\"Document Map\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Plain Text\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"E-mail Signature\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"HTML Top of Form\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"HTML Bottom of Form\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Normal (Web)\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"HTML Acronym\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"HTML Address\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"HTML Cite\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"HTML Code\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"HTML Definition\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"HTML Keyboard\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"HTML Preformatted\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"HTML Sample\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"HTML Typewriter\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"HTML Variable\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Normal Table\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"annotation subject\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"No List\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Outline List 1\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Outline List 2\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Outline List 3\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table Simple 1\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table Simple 2\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table Simple 3\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table Classic 1\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table Classic 2\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table Classic 3\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table Classic 4\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table Colorful 1\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table Colorful 2\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table Colorful 3\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table Columns 1\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table Columns 2\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table Columns 3\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table Columns 4\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table Columns 5\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table Grid 1\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table Grid 2\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table Grid 3\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table Grid 4\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table Grid 5\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table Grid 6\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table Grid 7\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table Grid 8\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table List 1\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table List 2\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table List 3\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table List 4\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table List 5\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table List 6\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table List 7\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table List 8\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table 3D effects 1\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table 3D effects 2\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table 3D effects 3\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table Contemporary\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table Elegant\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table Professional\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table Subtle 1\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table Subtle 2\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table Web 1\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table Web 2\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table Web 3\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Balloon Text\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table Grid\" w:uiPriority=\"39\"/><w:lsdException w:name=\"Table Theme\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Placeholder Text\" w:semiHidden=\"1\"/><w:lsdException w:name=\"No Spacing\" w:uiPriority=\"1\" w:qFormat=\"1\"/><w:lsdException w:name=\"Light Shading\" w:uiPriority=\"60\"/><w:lsdException w:name=\"Light List\" w:uiPriority=\"61\"/><w:lsdException w:name=\"Light Grid\" w:uiPriority=\"62\"/><w:lsdException w:name=\"Medium Shading 1\" w:uiPriority=\"63\"/><w:lsdException w:name=\"Medium Shading 2\" w:uiPriority=\"64\"/><w:lsdException w:name=\"Medium List 1\" w:uiPriority=\"65\"/><w:lsdException w:name=\"Medium List 2\" w:uiPriority=\"66\"/><w:lsdException w:name=\"Medium Grid 1\" w:uiPriority=\"67\"/><w:lsdException w:name=\"Medium Grid 2\" w:uiPriority=\"68\"/><w:lsdException w:name=\"Medium Grid 3\" w:uiPriority=\"69\"/><w:lsdException w:name=\"Dark List\" w:uiPriority=\"70\"/><w:lsdException w:name=\"Colorful Shading\" w:uiPriority=\"71\"/><w:lsdException w:name=\"Colorful List\" w:uiPriority=\"72\"/><w:lsdException w:name=\"Colorful Grid\" w:uiPriority=\"73\"/><w:lsdException w:name=\"Light Shading Accent 1\" w:uiPriority=\"60\"/><w:lsdException w:name=\"Light List Accent 1\" w:uiPriority=\"61\"/><w:lsdException w:name=\"Light Grid Accent 1\" w:uiPriority=\"62\"/><w:lsdException w:name=\"Medium Shading 1 Accent 1\" w:uiPriority=\"63\"/><w:lsdException w:name=\"Medium Shading 2 Accent 1\" w:uiPriority=\"64\"/><w:lsdException w:name=\"Medium List 1 Accent 1\" w:uiPriority=\"65\"/><w:lsdException w:name=\"Revision\" w:semiHidden=\"1\"/><w:lsdException w:name=\"List Paragraph\" w:uiPriority=\"34\" w:qFormat=\"1\"/><w:lsdException w:name=\"Quote\" w:uiPriority=\"29\" w:qFormat=\"1\"/><w:lsdException w:name=\"Intense Quote\" w:uiPriority=\"30\" w:qFormat=\"1\"/><w:lsdException w:name=\"Medium List 2 Accent 1\" w:uiPriority=\"66\"/><w:lsdException w:name=\"Medium Grid 1 Accent 1\" w:uiPriority=\"67\"/><w:lsdException w:name=\"Medium Grid 2 Accent 1\" w:uiPriority=\"68\"/><w:lsdException w:name=\"Medium Grid 3 Accent 1\" w:uiPriority=\"69\"/><w:lsdException w:name=\"Dark List Accent 1\" w:uiPriority=\"70\"/><w:lsdException w:name=\"Colorful Shading Accent 1\" w:uiPriority=\"71\"/><w:lsdException w:name=\"Colorful List Accent 1\" w:uiPriority=\"72\"/><w:lsdException w:name=\"Colorful Grid Accent 1\" w:uiPriority=\"73\"/><w:lsdException w:name=\"Light Shading Accent 2\" w:uiPriority=\"60\"/><w:lsdException w:name=\"Light List Accent 2\" w:uiPriority=\"61\"/><w:lsdException w:name=\"Light Grid Accent 2\" w:uiPriority=\"62\"/><w:lsdException w:name=\"Medium Shading 1 Accent 2\" w:uiPriority=\"63\"/><w:lsdException w:name=\"Medium Shading 2 Accent 2\" w:uiPriority=\"64\"/><w:lsdException w:name=\"Medium List 1 Accent 2\" w:uiPriority=\"65\"/><w:lsdException w:name=\"Medium List 2 Accent 2\" w:uiPriority=\"66\"/><w:lsdException w:name=\"Medium Grid 1 Accent 2\" w:uiPriority=\"67\"/><w:lsdException w:name=\"Medium Grid 2 Accent 2\" w:uiPriority=\"68\"/><w:lsdException w:name=\"Medium Grid 3 Accent 2\" w:uiPriority=\"69\"/><w:lsdException w:name=\"Dark List Accent 2\" w:uiPriority=\"70\"/><w:lsdException w:name=\"Colorful Shading Accent 2\" w:uiPriority=\"71\"/><w:lsdException w:name=\"Colorful List Accent 2\" w:uiPriority=\"72\"/><w:lsdException w:name=\"Colorful Grid Accent 2\" w:uiPriority=\"73\"/><w:lsdException w:name=\"Light Shading Accent 3\" w:uiPriority=\"60\"/><w:lsdException w:name=\"Light List Accent 3\" w:uiPriority=\"61\"/><w:lsdException w:name=\"Light Grid Accent 3\" w:uiPriority=\"62\"/><w:lsdException w:name=\"Medium Shading 1 Accent 3\" w:uiPriority=\"63\"/><w:lsdException w:name=\"Medium Shading 2 Accent 3\" w:uiPriority=\"64\"/><w:lsdException w:name=\"Medium List 1 Accent 3\" w:uiPriority=\"65\"/><w:lsdException w:name=\"Medium List 2 Accent 3\" w:uiPriority=\"66\"/><w:lsdException w:name=\"Medium Grid 1 Accent 3\" w:uiPriority=\"67\"/><w:lsdException w:name=\"Medium Grid 2 Accent 3\" w:uiPriority=\"68\"/><w:lsdException w:name=\"Medium Grid 3 Accent 3\" w:uiPriority=\"69\"/><w:lsdException w:name=\"Dark List Accent 3\" w:uiPriority=\"70\"/><w:lsdException w:name=\"Colorful Shading Accent 3\" w:uiPriority=\"71\"/><w:lsdException w:name=\"Colorful List Accent 3\" w:uiPriority=\"72\"/><w:lsdException w:name=\"Colorful Grid Accent 3\" w:uiPriority=\"73\"/><w:lsdException w:name=\"Light Shading Accent 4\" w:uiPriority=\"60\"/><w:lsdException w:name=\"Light List Accent 4\" w:uiPriority=\"61\"/><w:lsdException w:name=\"Light Grid Accent 4\" w:uiPriority=\"62\"/><w:lsdException w:name=\"Medium Shading 1 Accent 4\" w:uiPriority=\"63\"/><w:lsdException w:name=\"Medium Shading 2 Accent 4\" w:uiPriority=\"64\"/><w:lsdException w:name=\"Medium List 1 Accent 4\" w:uiPriority=\"65\"/><w:lsdException w:name=\"Medium List 2 Accent 4\" w:uiPriority=\"66\"/><w:lsdException w:name=\"Medium Grid 1 Accent 4\" w:uiPriority=\"67\"/><w:lsdException w:name=\"Medium Grid 2 Accent 4\" w:uiPriority=\"68\"/><w:lsdException w:name=\"Medium Grid 3 Accent 4\" w:uiPriority=\"69\"/><w:lsdException w:name=\"Dark List Accent 4\" w:uiPriority=\"70\"/><w:lsdException w:name=\"Colorful Shading Accent 4\" w:uiPriority=\"71\"/><w:lsdException w:name=\"Colorful List Accent 4\" w:uiPriority=\"72\"/><w:lsdException w:name=\"Colorful Grid Accent 4\" w:uiPriority=\"73\"/><w:lsdException w:name=\"Light Shading Accent 5\" w:uiPriority=\"60\"/><w:lsdException w:name=\"Light List Accent 5\" w:uiPriority=\"61\"/><w:lsdException w:name=\"Light Grid Accent 5\" w:uiPriority=\"62\"/><w:lsdException w:name=\"Medium Shading 1 Accent 5\" w:uiPriority=\"63\"/><w:lsdException w:name=\"Medium Shading 2 Accent 5\" w:uiPriority=\"64\"/><w:lsdException w:name=\"Medium List 1 Accent 5\" w:uiPriority=\"65\"/><w:lsdException w:name=\"Medium List 2 Accent 5\" w:uiPriority=\"66\"/><w:lsdException w:name=\"Medium Grid 1 Accent 5\" w:uiPriority=\"67\"/><w:lsdException w:name=\"Medium Grid 2 Accent 5\" w:uiPriority=\"68\"/><w:lsdException w:name=\"Medium Grid 3 Accent 5\" w:uiPriority=\"69\"/><w:lsdException w:name=\"Dark List Accent 5\" w:uiPriority=\"70\"/><w:lsdException w:name=\"Colorful Shading Accent 5\" w:uiPriority=\"71\"/><w:lsdException w:name=\"Colorful List Accent 5\" w:uiPriority=\"72\"/><w:lsdException w:name=\"Colorful Grid Accent 5\" w:uiPriority=\"73\"/><w:lsdException w:name=\"Light Shading Accent 6\" w:uiPriority=\"60\"/><w:lsdException w:name=\"Light List Accent 6\" w:uiPriority=\"61\"/><w:lsdException w:name=\"Light Grid Accent 6\" w:uiPriority=\"62\"/><w:lsdException w:name=\"Medium Shading 1 Accent 6\" w:uiPriority=\"63\"/><w:lsdException w:name=\"Medium Shading 2 Accent 6\" w:uiPriority=\"64\"/><w:lsdException w:name=\"Medium List 1 Accent 6\" w:uiPriority=\"65\"/><w:lsdException w:name=\"Medium List 2 Accent 6\" w:uiPriority=\"66\"/><w:lsdException w:name=\"Medium Grid 1 Accent 6\" w:uiPriority=\"67\"/><w:lsdException w:name=\"Medium Grid 2 Accent 6\" w:uiPriority=\"68\"/><w:lsdException w:name=\"Medium Grid 3 Accent 6\" w:uiPriority=\"69\"/><w:lsdException w:name=\"Dark List Accent 6\" w:uiPriority=\"70\"/><w:lsdException w:name=\"Colorful Shading Accent 6\" w:uiPriority=\"71\"/><w:lsdException w:name=\"Colorful List Accent 6\" w:uiPriority=\"72\"/><w:lsdException w:name=\"Colorful Grid Accent 6\" w:uiPriority=\"73\"/><w:lsdException w:name=\"Subtle Emphasis\" w:uiPriority=\"19\" w:qFormat=\"1\"/><w:lsdException w:name=\"Intense Emphasis\" w:uiPriority=\"21\" w:qFormat=\"1\"/><w:lsdException w:name=\"Subtle Reference\" w:uiPriority=\"31\" w:qFormat=\"1\"/><w:lsdException w:name=\"Intense Reference\" w:uiPriority=\"32\" w:qFormat=\"1\"/><w:lsdException w:name=\"Book Title\" w:uiPriority=\"33\" w:qFormat=\"1\"/><w:lsdException w:name=\"Bibliography\" w:semiHidden=\"1\" w:uiPriority=\"37\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"TOC Heading\" w:semiHidden=\"1\" w:uiPriority=\"39\" w:unhideWhenUsed=\"1\" w:qFormat=\"1\"/><w:lsdException w:name=\"Plain Table 1\" w:uiPriority=\"41\"/><w:lsdException w:name=\"Plain Table 2\" w:uiPriority=\"42\"/><w:lsdException w:name=\"Plain Table 3\" w:uiPriority=\"43\"/><w:lsdException w:name=\"Plain Table 4\" w:uiPriority=\"44\"/><w:lsdException w:name=\"Plain Table 5\" w:uiPriority=\"45\"/><w:lsdException w:name=\"Grid Table Light\" w:uiPriority=\"40\"/><w:lsdException w:name=\"Grid Table 1 Light\" w:uiPriority=\"46\"/><w:lsdException w:name=\"Grid Table 2\" w:uiPriority=\"47\"/><w:lsdException w:name=\"Grid Table 3\" w:uiPriority=\"48\"/><w:lsdException w:name=\"Grid Table 4\" w:uiPriority=\"49\"/><w:lsdException w:name=\"Grid Table 5 Dark\" w:uiPriority=\"50\"/><w:lsdException w:name=\"Grid Table 6 Colorful\" w:uiPriority=\"51\"/><w:lsdException w:name=\"Grid Table 7 Colorful\" w:uiPriority=\"52\"/><w:lsdException w:name=\"Grid Table 1 Light Accent 1\" w:uiPriority=\"46\"/><w:lsdException w:name=\"Grid Table 2 Accent 1\" w:uiPriority=\"47\"/><w:lsdException w:name=\"Grid Table 3 Accent 1\" w:uiPriority=\"48\"/><w:lsdException w:name=\"Grid Table 4 Accent 1\" w:uiPriority=\"49\"/><w:lsdException w:name=\"Grid Table 5 Dark Accent 1\" w:uiPriority=\"50\"/><w:lsdException w:name=\"Grid Table 6 Colorful Accent 1\" w:uiPriority=\"51\"/><w:lsdException w:name=\"Grid Table 7 Colorful Accent 1\" w:uiPriority=\"52\"/><w:lsdException w:name=\"Grid Table 1 Light Accent 2\" w:uiPriority=\"46\"/><w:lsdException w:name=\"Grid Table 2 Accent 2\" w:uiPriority=\"47\"/><w:lsdException w:name=\"Grid Table 3 Accent 2\" w:uiPriority=\"48\"/><w:lsdException w:name=\"Grid Table 4 Accent 2\" w:uiPriority=\"49\"/><w:lsdException w:name=\"Grid Table 5 Dark Accent 2\" w:uiPriority=\"50\"/><w:lsdException w:name=\"Grid Table 6 Colorful Accent 2\" w:uiPriority=\"51\"/><w:lsdException w:name=\"Grid Table 7 Colorful Accent 2\" w:uiPriority=\"52\"/><w:lsdException w:name=\"Grid Table 1 Light Accent 3\" w:uiPriority=\"46\"/><w:lsdException w:name=\"Grid Table 2 Accent 3\" w:uiPriority=\"47\"/><w:lsdException w:name=\"Grid Table 3 Accent 3\" w:uiPriority=\"48\"/><w:lsdException w:name=\"Grid Table 4 Accent 3\" w:uiPriority=\"49\"/><w:lsdException w:name=\"Grid Table 5 Dark Accent 3\" w:uiPriority=\"50\"/><w:lsdException w:name=\"Grid Table 6 Colorful Accent 3\" w:uiPriority=\"51\"/><w:lsdException w:name=\"Grid Table 7 Colorful Accent 3\" w:uiPriority=\"52\"/><w:lsdException w:name=\"Grid Table 1 Light Accent 4\" w:uiPriority=\"46\"/><w:lsdException w:name=\"Grid Table 2 Accent 4\" w:uiPriority=\"47\"/><w:lsdException w:name=\"Grid Table 3 Accent 4\" w:uiPriority=\"48\"/><w:lsdException w:name=\"Grid Table 4 Accent 4\" w:uiPriority=\"49\"/><w:lsdException w:name=\"Grid Table 5 Dark Accent 4\" w:uiPriority=\"50\"/><w:lsdException w:name=\"Grid Table 6 Colorful Accent 4\" w:uiPriority=\"51\"/><w:lsdException w:name=\"Grid Table 7 Colorful Accent 4\" w:uiPriority=\"52\"/><w:lsdException w:name=\"Grid Table 1 Light Accent 5\" w:uiPriority=\"46\"/><w:lsdException w:name=\"Grid Table 2 Accent 5\" w:uiPriority=\"47\"/><w:lsdException w:name=\"Grid Table 3 Accent 5\" w:uiPriority=\"48\"/><w:lsdException w:name=\"Grid Table 4 Accent 5\" w:uiPriority=\"49\"/><w:lsdException w:name=\"Grid Table 5 Dark Accent 5\" w:uiPriority=\"50\"/><w:lsdException w:name=\"Grid Table 6 Colorful Accent 5\" w:uiPriority=\"51\"/><w:lsdException w:name=\"Grid Table 7 Colorful Accent 5\" w:uiPriority=\"52\"/><w:lsdException w:name=\"Grid Table 1 Light Accent 6\" w:uiPriority=\"46\"/><w:lsdException w:name=\"Grid Table 2 Accent 6\" w:uiPriority=\"47\"/><w:lsdException w:name=\"Grid Table 3 Accent 6\" w:uiPriority=\"48\"/><w:lsdException w:name=\"Grid Table 4 Accent 6\" w:uiPriority=\"49\"/><w:lsdException w:name=\"Grid Table 5 Dark Accent 6\" w:uiPriority=\"50\"/><w:lsdException w:name=\"Grid Table 6 Colorful Accent 6\" w:uiPriority=\"51\"/><w:lsdException w:name=\"Grid Table 7 Colorful Accent 6\" w:uiPriority=\"52\"/><w:lsdException w:name=\"List Table 1 Light\" w:uiPriority=\"46\"/><w:lsdException w:name=\"List Table 2\" w:uiPriority=\"47\"/><w:lsdException w:name=\"List Table 3\" w:uiPriority=\"48\"/><w:lsdException w:name=\"List Table 4\" w:uiPriority=\"49\"/><w:lsdException w:name=\"List Table 5 Dark\" w:uiPriority=\"50\"/><w:lsdException w:name=\"List Table 6 Colorful\" w:uiPriority=\"51\"/><w:lsdException w:name=\"List Table 7 Colorful\" w:uiPriority=\"52\"/><w:lsdException w:name=\"List Table 1 Light Accent 1\" w:uiPriority=\"46\"/><w:lsdException w:name=\"List Table 2 Accent 1\" w:uiPriority=\"47\"/><w:lsdException w:name=\"List Table 3 Accent 1\" w:uiPriority=\"48\"/><w:lsdException w:name=\"List Table 4 Accent 1\" w:uiPriority=\"49\"/><w:lsdException w:name=\"List Table 5 Dark Accent 1\" w:uiPriority=\"50\"/><w:lsdException w:name=\"List Table 6 Colorful Accent 1\" w:uiPriority=\"51\"/><w:lsdException w:name=\"List Table 7 Colorful Accent 1\" w:uiPriority=\"52\"/><w:lsdException w:name=\"List Table 1 Light Accent 2\" w:uiPriority=\"46\"/><w:lsdException w:name=\"List Table 2 Accent 2\" w:uiPriority=\"47\"/><w:lsdException w:name=\"List Table 3 Accent 2\" w:uiPriority=\"48\"/><w:lsdException w:name=\"List Table 4 Accent 2\" w:uiPriority=\"49\"/><w:lsdException w:name=\"List Table 5 Dark Accent 2\" w:uiPriority=\"50\"/><w:lsdException w:name=\"List Table 6 Colorful Accent 2\" w:uiPriority=\"51\"/><w:lsdException w:name=\"List Table 7 Colorful Accent 2\" w:uiPriority=\"52\"/><w:lsdException w:name=\"List Table 1 Light Accent 3\" w:uiPriority=\"46\"/><w:lsdException w:name=\"List Table 2 Accent 3\" w:uiPriority=\"47\"/><w:lsdException w:name=\"List Table 3 Accent 3\" w:uiPriority=\"48\"/><w:lsdException w:name=\"List Table 4 Accent 3\" w:uiPriority=\"49\"/><w:lsdException w:name=\"List Table 5 Dark Accent 3\" w:uiPriority=\"50\"/><w:lsdException w:name=\"List Table 6 Colorful Accent 3\" w:uiPriority=\"51\"/><w:lsdException w:name=\"List Table 7 Colorful Accent 3\" w:uiPriority=\"52\"/><w:lsdException w:name=\"List Table 1 Light Accent 4\" w:uiPriority=\"46\"/><w:lsdException w:name=\"List Table 2 Accent 4\" w:uiPriority=\"47\"/><w:lsdException w:name=\"List Table 3 Accent 4\" w:uiPriority=\"48\"/><w:lsdException w:name=\"List Table 4 Accent 4\" w:uiPriority=\"49\"/><w:lsdException w:name=\"List Table 5 Dark Accent 4\" w:uiPriority=\"50\"/><w:lsdException w:name=\"List Table 6 Colorful Accent 4\" w:uiPriority=\"51\"/><w:lsdException w:name=\"List Table 7 Colorful Accent 4\" w:uiPriority=\"52\"/><w:lsdException w:name=\"List Table 1 Light Accent 5\" w:uiPriority=\"46\"/><w:lsdException w:name=\"List Table 2 Accent 5\" w:uiPriority=\"47\"/><w:lsdException w:name=\"List Table 3 Accent 5\" w:uiPriority=\"48\"/><w:lsdException w:name=\"List Table 4 Accent 5\" w:uiPriority=\"49\"/><w:lsdException w:name=\"List Table 5 Dark Accent 5\" w:uiPriority=\"50\"/><w:lsdException w:name=\"List Table 6 Colorful Accent 5\" w:uiPriority=\"51\"/><w:lsdException w:name=\"List Table 7 Colorful Accent 5\" w:uiPriority=\"52\"/><w:lsdException w:name=\"List Table 1 Light Accent 6\" w:uiPriority=\"46\"/><w:lsdException w:name=\"List Table 2 Accent 6\" w:uiPriority=\"47\"/><w:lsdException w:name=\"List Table 3 Accent 6\" w:uiPriority=\"48\"/><w:lsdException w:name=\"List Table 4 Accent 6\" w:uiPriority=\"49\"/><w:lsdException w:name=\"List Table 5 Dark Accent 6\" w:uiPriority=\"50\"/><w:lsdException w:name=\"List Table 6 Colorful Accent 6\" w:uiPriority=\"51\"/><w:lsdException w:name=\"List Table 7 Colorful Accent 6\" w:uiPriority=\"52\"/></w:latentStyles><w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"a\"><w:name w:val=\"Normal\"/><w:qFormat/><w:pPr><w:widowControl w:val=\"0\"/><w:jc w:val=\"both\"/></w:pPr></w:style><w:style w:type=\"paragraph\" w:styleId=\"1\"><w:name w:val=\"heading 1\"/><w:basedOn w:val=\"a\"/><w:next w:val=\"a\"/><w:link w:val=\"1Char\"/><w:uiPriority w:val=\"9\"/><w:qFormat/><w:rsid w:val=\"007E54FF\"/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before=\"340\" w:after=\"330\" w:line=\"578\" w:lineRule=\"auto\"/><w:outlineLvl w:val=\"0\"/></w:pPr><w:rPr><w:b/><w:bCs/><w:kern w:val=\"44\"/><w:sz w:val=\"44\"/><w:szCs w:val=\"44\"/></w:rPr></w:style><w:style w:type=\"paragraph\" w:styleId=\"2\"><w:name w:val=\"heading 2\"/><w:basedOn w:val=\"a\"/><w:next w:val=\"a\"/><w:link w:val=\"2Char\"/><w:uiPriority w:val=\"9\"/><w:unhideWhenUsed/><w:qFormat/><w:rsid w:val=\"007E54FF\"/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before=\"260\" w:after=\"260\" w:line=\"416\" w:lineRule=\"auto\"/><w:outlineLvl w:val=\"1\"/></w:pPr><w:rPr><w:rFonts w:asciiTheme=\"majorHAnsi\" w:eastAsiaTheme=\"majorEastAsia\" w:hAnsiTheme=\"majorHAnsi\" w:cstheme=\"majorBidi\"/><w:b/><w:bCs/><w:sz w:val=\"32\"/><w:szCs w:val=\"32\"/></w:rPr></w:style><w:style w:type=\"paragraph\" w:styleId=\"3\"><w:name w:val=\"heading 3\"/><w:basedOn w:val=\"a\"/><w:next w:val=\"a\"/><w:link w:val=\"3Char\"/><w:uiPriority w:val=\"9\"/><w:unhideWhenUsed/><w:qFormat/><w:rsid w:val=\"007E54FF\"/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before=\"260\" w:after=\"260\" w:line=\"416\" w:lineRule=\"auto\"/><w:outlineLvl w:val=\"2\"/></w:pPr><w:rPr><w:b/><w:bCs/><w:sz w:val=\"32\"/><w:szCs w:val=\"32\"/></w:rPr></w:style><w:style w:type=\"paragraph\" w:styleId=\"4\"><w:name w:val=\"heading 4\"/><w:basedOn w:val=\"a\"/><w:next w:val=\"a\"/><w:link w:val=\"4Char\"/><w:uiPriority w:val=\"9\"/><w:unhideWhenUsed/><w:qFormat/><w:rsid w:val=\"007E54FF\"/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before=\"280\" w:after=\"290\" w:line=\"376\" w:lineRule=\"auto\"/><w:outlineLvl w:val=\"3\"/></w:pPr><w:rPr><w:rFonts w:asciiTheme=\"majorHAnsi\" w:eastAsiaTheme=\"majorEastAsia\" w:hAnsiTheme=\"majorHAnsi\" w:cstheme=\"majorBidi\"/><w:b/><w:bCs/><w:sz w:val=\"28\"/><w:szCs w:val=\"28\"/></w:rPr></w:style><w:style w:type=\"paragraph\" w:styleId=\"5\"><w:name w:val=\"heading 5\"/><w:basedOn w:val=\"a\"/><w:next w:val=\"a\"/><w:link w:val=\"5Char\"/><w:uiPriority w:val=\"9\"/><w:unhideWhenUsed/><w:qFormat/><w:rsid w:val=\"007E54FF\"/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before=\"280\" w:after=\"290\" w:line=\"376\" w:lineRule=\"auto\"/><w:outlineLvl w:val=\"4\"/></w:pPr><w:rPr><w:b/><w:bCs/><w:sz w:val=\"28\"/><w:szCs w:val=\"28\"/></w:rPr></w:style><w:style w:type=\"paragraph\" w:styleId=\"6\"><w:name w:val=\"heading 6\"/><w:basedOn w:val=\"a\"/><w:next w:val=\"a\"/><w:link w:val=\"6Char\"/><w:uiPriority w:val=\"9\"/><w:unhideWhenUsed/><w:qFormat/><w:rsid w:val=\"007E54FF\"/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before=\"240\" w:after=\"64\" w:line=\"320\" w:lineRule=\"auto\"/><w:outlineLvl w:val=\"5\"/></w:pPr><w:rPr><w:rFonts w:asciiTheme=\"majorHAnsi\" w:eastAsiaTheme=\"majorEastAsia\" w:hAnsiTheme=\"majorHAnsi\" w:cstheme=\"majorBidi\"/><w:b/><w:bCs/><w:sz w:val=\"24\"/><w:szCs w:val=\"24\"/></w:rPr></w:style><w:style w:type=\"character\" w:default=\"1\" w:styleId=\"a0\"><w:name w:val=\"Default Paragraph Font\"/><w:uiPriority w:val=\"1\"/><w:semiHidden/><w:unhideWhenUsed/></w:style><w:style w:type=\"table\" w:default=\"1\" w:styleId=\"a1\"><w:name w:val=\"Normal Table\"/><w:uiPriority w:val=\"99\"/><w:semiHidden/><w:unhideWhenUsed/><w:tblPr><w:tblInd w:w=\"0\" w:type=\"dxa\"/><w:tblCellMar><w:top w:w=\"0\" w:type=\"dxa\"/><w:left w:w=\"108\" w:type=\"dxa\"/><w:bottom w:w=\"0\" w:type=\"dxa\"/><w:right w:w=\"108\" w:type=\"dxa\"/></w:tblCellMar></w:tblPr></w:style><w:style w:type=\"numbering\" w:default=\"1\" w:styleId=\"a2\"><w:name w:val=\"No List\"/><w:uiPriority w:val=\"99\"/><w:semiHidden/><w:unhideWhenUsed/></w:style><w:style w:type=\"character\" w:customStyle=\"1\" w:styleId=\"1Char\"><w:name w:val=\"标题 1 Char\"/><w:basedOn w:val=\"a0\"/><w:link w:val=\"1\"/><w:uiPriority w:val=\"9\"/><w:rsid w:val=\"007E54FF\"/><w:rPr><w:b/><w:bCs/><w:kern w:val=\"44\"/><w:sz w:val=\"44\"/><w:szCs w:val=\"44\"/></w:rPr></w:style><w:style w:type=\"character\" w:customStyle=\"1\" w:styleId=\"2Char\"><w:name w:val=\"标题 2 Char\"/><w:basedOn w:val=\"a0\"/><w:link w:val=\"2\"/><w:uiPriority w:val=\"9\"/><w:rsid w:val=\"007E54FF\"/><w:rPr><w:rFonts w:asciiTheme=\"majorHAnsi\" w:eastAsiaTheme=\"majorEastAsia\" w:hAnsiTheme=\"majorHAnsi\" w:cstheme=\"majorBidi\"/><w:b/><w:bCs/><w:sz w:val=\"32\"/><w:szCs w:val=\"32\"/></w:rPr></w:style><w:style w:type=\"character\" w:customStyle=\"1\" w:styleId=\"3Char\"><w:name w:val=\"标题 3 Char\"/><w:basedOn w:val=\"a0\"/><w:link w:val=\"3\"/><w:uiPriority w:val=\"9\"/><w:rsid w:val=\"007E54FF\"/><w:rPr><w:b/><w:bCs/><w:sz w:val=\"32\"/><w:szCs w:val=\"32\"/></w:rPr></w:style><w:style w:type=\"character\" w:customStyle=\"1\" w:styleId=\"4Char\"><w:name w:val=\"标题 4 Char\"/><w:basedOn w:val=\"a0\"/><w:link w:val=\"4\"/><w:uiPriority w:val=\"9\"/><w:rsid w:val=\"007E54FF\"/><w:rPr><w:rFonts w:asciiTheme=\"majorHAnsi\" w:eastAsiaTheme=\"majorEastAsia\" w:hAnsiTheme=\"majorHAnsi\" w:cstheme=\"majorBidi\"/><w:b/><w:bCs/><w:sz w:val=\"28\"/><w:szCs w:val=\"28\"/></w:rPr></w:style><w:style w:type=\"character\" w:customStyle=\"1\" w:styleId=\"5Char\"><w:name w:val=\"标题 5 Char\"/><w:basedOn w:val=\"a0\"/><w:link w:val=\"5\"/><w:uiPriority w:val=\"9\"/><w:rsid w:val=\"007E54FF\"/><w:rPr><w:b/><w:bCs/><w:sz w:val=\"28\"/><w:szCs w:val=\"28\"/></w:rPr></w:style><w:style w:type=\"character\" w:customStyle=\"1\" w:styleId=\"6Char\"><w:name w:val=\"标题 6 Char\"/><w:basedOn w:val=\"a0\"/><w:link w:val=\"6\"/><w:uiPriority w:val=\"9\"/><w:rsid w:val=\"007E54FF\"/><w:rPr><w:rFonts w:asciiTheme=\"majorHAnsi\" w:eastAsiaTheme=\"majorEastAsia\" w:hAnsiTheme=\"majorHAnsi\" w:cstheme=\"majorBidi\"/><w:b/><w:bCs/><w:sz w:val=\"24\"/><w:szCs w:val=\"24\"/></w:rPr></w:style><w:style w:type=\"table\" w:styleId=\"a3\"><w:name w:val=\"Table Grid\"/><w:basedOn w:val=\"a1\"/><w:uiPriority w:val=\"39\"/><w:rsid w:val=\"00DF50C9\"/><w:tblPr><w:tblBorders><w:top w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/><w:left w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/><w:bottom w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/><w:right w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/><w:insideH w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/><w:insideV w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/></w:tblBorders></w:tblPr></w:style><w:style w:type=\"paragraph\" w:styleId=\"a4\"><w:name w:val=\"List Paragraph\"/><w:basedOn w:val=\"a\"/><w:uiPriority w:val=\"34\"/><w:qFormat/><w:rsid w:val=\"000D4B8D\"/><w:pPr><w:ind w:firstLineChars=\"200\" w:firstLine=\"420\"/></w:pPr></w:style></w:styles>"
    }, {
        "name": "word/webSettings.xml",
        "value": "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n<w:webSettings xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:w15=\"http://schemas.microsoft.com/office/word/2012/wordml\" mc:Ignorable=\"w14 w15\"><w:optimizeForBrowser/><w:allowPNG/></w:webSettings>"
    }, {
        "name": "word/theme/theme1.xml",
        "value": "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"Office 主题\"><a:themeElements><a:clrScheme name=\"Office\"><a:dk1><a:sysClr val=\"windowText\" lastClr=\"000000\"/></a:dk1><a:lt1><a:sysClr val=\"window\" lastClr=\"FFFFFF\"/></a:lt1><a:dk2><a:srgbClr val=\"44546A\"/></a:dk2><a:lt2><a:srgbClr val=\"E7E6E6\"/></a:lt2><a:accent1><a:srgbClr val=\"5B9BD5\"/></a:accent1><a:accent2><a:srgbClr val=\"ED7D31\"/></a:accent2><a:accent3><a:srgbClr val=\"A5A5A5\"/></a:accent3><a:accent4><a:srgbClr val=\"FFC000\"/></a:accent4><a:accent5><a:srgbClr val=\"4472C4\"/></a:accent5><a:accent6><a:srgbClr val=\"70AD47\"/></a:accent6><a:hlink><a:srgbClr val=\"0563C1\"/></a:hlink><a:folHlink><a:srgbClr val=\"954F72\"/></a:folHlink></a:clrScheme><a:fontScheme name=\"Office\"><a:majorFont><a:latin typeface=\"Calibri Light\" panose=\"020F0302020204030204\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/><a:font script=\"Jpan\" typeface=\"ＭＳ ゴシック\"/><a:font script=\"Hang\" typeface=\"맑은 고딕\"/><a:font script=\"Hans\" typeface=\"宋体\"/><a:font script=\"Hant\" typeface=\"新細明體\"/><a:font script=\"Arab\" typeface=\"Times New Roman\"/><a:font script=\"Hebr\" typeface=\"Times New Roman\"/><a:font script=\"Thai\" typeface=\"Angsana New\"/><a:font script=\"Ethi\" typeface=\"Nyala\"/><a:font script=\"Beng\" typeface=\"Vrinda\"/><a:font script=\"Gujr\" typeface=\"Shruti\"/><a:font script=\"Khmr\" typeface=\"MoolBoran\"/><a:font script=\"Knda\" typeface=\"Tunga\"/><a:font script=\"Guru\" typeface=\"Raavi\"/><a:font script=\"Cans\" typeface=\"Euphemia\"/><a:font script=\"Cher\" typeface=\"Plantagenet Cherokee\"/><a:font script=\"Yiii\" typeface=\"Microsoft Yi Baiti\"/><a:font script=\"Tibt\" typeface=\"Microsoft Himalaya\"/><a:font script=\"Thaa\" typeface=\"MV Boli\"/><a:font script=\"Deva\" typeface=\"Mangal\"/><a:font script=\"Telu\" typeface=\"Gautami\"/><a:font script=\"Taml\" typeface=\"Latha\"/><a:font script=\"Syrc\" typeface=\"Estrangelo Edessa\"/><a:font script=\"Orya\" typeface=\"Kalinga\"/><a:font script=\"Mlym\" typeface=\"Kartika\"/><a:font script=\"Laoo\" typeface=\"DokChampa\"/><a:font script=\"Sinh\" typeface=\"Iskoola Pota\"/><a:font script=\"Mong\" typeface=\"Mongolian Baiti\"/><a:font script=\"Viet\" typeface=\"Times New Roman\"/><a:font script=\"Uigh\" typeface=\"Microsoft Uighur\"/><a:font script=\"Geor\" typeface=\"Sylfaen\"/></a:majorFont><a:minorFont><a:latin typeface=\"Calibri\" panose=\"020F0502020204030204\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/><a:font script=\"Jpan\" typeface=\"ＭＳ 明朝\"/><a:font script=\"Hang\" typeface=\"맑은 고딕\"/><a:font script=\"Hans\" typeface=\"宋体\"/><a:font script=\"Hant\" typeface=\"新細明體\"/><a:font script=\"Arab\" typeface=\"Arial\"/><a:font script=\"Hebr\" typeface=\"Arial\"/><a:font script=\"Thai\" typeface=\"Cordia New\"/><a:font script=\"Ethi\" typeface=\"Nyala\"/><a:font script=\"Beng\" typeface=\"Vrinda\"/><a:font script=\"Gujr\" typeface=\"Shruti\"/><a:font script=\"Khmr\" typeface=\"DaunPenh\"/><a:font script=\"Knda\" typeface=\"Tunga\"/><a:font script=\"Guru\" typeface=\"Raavi\"/><a:font script=\"Cans\" typeface=\"Euphemia\"/><a:font script=\"Cher\" typeface=\"Plantagenet Cherokee\"/><a:font script=\"Yiii\" typeface=\"Microsoft Yi Baiti\"/><a:font script=\"Tibt\" typeface=\"Microsoft Himalaya\"/><a:font script=\"Thaa\" typeface=\"MV Boli\"/><a:font script=\"Deva\" typeface=\"Mangal\"/><a:font script=\"Telu\" typeface=\"Gautami\"/><a:font script=\"Taml\" typeface=\"Latha\"/><a:font script=\"Syrc\" typeface=\"Estrangelo Edessa\"/><a:font script=\"Orya\" typeface=\"Kalinga\"/><a:font script=\"Mlym\" typeface=\"Kartika\"/><a:font script=\"Laoo\" typeface=\"DokChampa\"/><a:font script=\"Sinh\" typeface=\"Iskoola Pota\"/><a:font script=\"Mong\" typeface=\"Mongolian Baiti\"/><a:font script=\"Viet\" typeface=\"Arial\"/><a:font script=\"Uigh\" typeface=\"Microsoft Uighur\"/><a:font script=\"Geor\" typeface=\"Sylfaen\"/></a:minorFont></a:fontScheme><a:fmtScheme name=\"Office\"><a:fillStyleLst><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:gradFill rotWithShape=\"1\"><a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"110000\"/><a:satMod val=\"105000\"/><a:tint val=\"67000\"/></a:schemeClr></a:gs><a:gs pos=\"50000\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"105000\"/><a:satMod val=\"103000\"/><a:tint val=\"73000\"/></a:schemeClr></a:gs><a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"105000\"/><a:satMod val=\"109000\"/><a:tint val=\"81000\"/></a:schemeClr></a:gs></a:gsLst><a:lin ang=\"5400000\" scaled=\"0\"/></a:gradFill><a:gradFill rotWithShape=\"1\"><a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:satMod val=\"103000\"/><a:lumMod val=\"102000\"/><a:tint val=\"94000\"/></a:schemeClr></a:gs><a:gs pos=\"50000\"><a:schemeClr val=\"phClr\"><a:satMod val=\"110000\"/><a:lumMod val=\"100000\"/><a:shade val=\"100000\"/></a:schemeClr></a:gs><a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"99000\"/><a:satMod val=\"120000\"/><a:shade val=\"78000\"/></a:schemeClr></a:gs></a:gsLst><a:lin ang=\"5400000\" scaled=\"0\"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w=\"6350\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/><a:miter lim=\"800000\"/></a:ln><a:ln w=\"12700\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/><a:miter lim=\"800000\"/></a:ln><a:ln w=\"19050\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/><a:miter lim=\"800000\"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad=\"57150\" dist=\"19050\" dir=\"5400000\" algn=\"ctr\" rotWithShape=\"0\"><a:srgbClr val=\"000000\"><a:alpha val=\"63000\"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:solidFill><a:schemeClr val=\"phClr\"><a:tint val=\"95000\"/><a:satMod val=\"170000\"/></a:schemeClr></a:solidFill><a:gradFill rotWithShape=\"1\"><a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:tint val=\"93000\"/><a:satMod val=\"150000\"/><a:shade val=\"98000\"/><a:lumMod val=\"102000\"/></a:schemeClr></a:gs><a:gs pos=\"50000\"><a:schemeClr val=\"phClr\"><a:tint val=\"98000\"/><a:satMod val=\"130000\"/><a:shade val=\"90000\"/><a:lumMod val=\"103000\"/></a:schemeClr></a:gs><a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:shade val=\"63000\"/><a:satMod val=\"120000\"/></a:schemeClr></a:gs></a:gsLst><a:lin ang=\"5400000\" scaled=\"0\"/></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/><a:extLst><a:ext uri=\"{05A4C25C-085E-4340-85A3-A5531E510DB2}\"><thm15:themeFamily xmlns:thm15=\"http://schemas.microsoft.com/office/thememl/2012/main\" name=\"Office Theme\" id=\"{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}\" vid=\"{4A3C46E8-61CC-4603-A589-7422A47A8E4A}\"/></a:ext></a:extLst></a:theme>"
    }
];