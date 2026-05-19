// 将文档按照GB/T 9704-2012 规范格式化
// 牟雷
// 2026年02月05日 21:00
// JS版本：2026年03月10日

function Format2GBT9704_2012() {
    var doc = Application.ActiveDocument;

    // 将所有手动换行符(软回车)替换为段落标记(硬回车)
    // 关闭屏幕更新以加快速度并防止闪烁
    Application.ScreenUpdating = false;
    var find = doc.Content.Find;
    find.ClearFormatting();
    find.Replacement.ClearFormatting();
    // ^l 代表手动换行符 (Shift+Enter)
    find.Text = "^l";
    // ^p 代表段落标记 (Enter)
    find.Replacement.Text = "^p";
    find.Execute(null, null, null, null, null, null, null, null, null, null, wdReplaceAll);
    // 恢复屏幕更新
    Application.ScreenUpdating = true;

    // 一次性转换全文自动编号为文本（保留递增编号）
    var docRange = doc.Content;  // 获取整个文档的范围
    // 一次性转换所有自动编号为纯文本（保留原递增编号：1.、2.、3.）
    docRange.ListFormat.ConvertNumbersToText();

    // 页面设置
    doc.PageSetup.TopMargin = Application.CentimetersToPoints(3.7);
    doc.PageSetup.BottomMargin = Application.CentimetersToPoints(3.5);
    doc.PageSetup.LeftMargin = Application.CentimetersToPoints(2.7);
    doc.PageSetup.RightMargin = Application.CentimetersToPoints(2.7);

    // 定义字号
    var FontSize_Normal = 16;   // 三号 = 16 磅
    var FontSize_Title1 = 22;   // 二号 ≈ 22 磅

    // 设置 Normal 样式（正文）
    var normalStyle = doc.Styles.Item(wdStyleNormal);
    normalStyle.Font.Name = "仿宋_GB2312";
    normalStyle.Font.Size = FontSize_Normal;
    normalStyle.ParagraphFormat.LineSpacingRule = wdLineSpaceExactly;
    normalStyle.ParagraphFormat.LineSpacing = 28;
    normalStyle.ParagraphFormat.FirstLineIndent = CharactersToPoints(2);
    normalStyle.ParagraphFormat.SpaceBefore = 0;
    normalStyle.ParagraphFormat.SpaceAfter = 0;

    // 设置 一级标题
    var h1Style = doc.Styles.Item(wdStyleHeading1);
    h1Style.AutomaticallyUpdate = false;
    h1Style.Font.Name = "黑体";
    h1Style.Font.Size = FontSize_Normal;
    h1Style.Font.Bold = false;   // 取消加粗
    h1Style.ParagraphFormat.LineSpacingRule = wdLineSpaceExactly;
    h1Style.ParagraphFormat.LineSpacing = 28;
    h1Style.ParagraphFormat.SpaceBefore = 0;
    h1Style.ParagraphFormat.SpaceAfter = 0;
    h1Style.ParagraphFormat.SpaceBeforeAuto = false;
    h1Style.ParagraphFormat.SpaceAfterAuto = false;
    h1Style.ParagraphFormat.FirstLineIndent = CharactersToPoints(2);

    // 设置 二级标题
    var h2Style = doc.Styles.Item(wdStyleHeading2);
    h2Style.AutomaticallyUpdate = false;
    h2Style.Font.Name = "楷体_GB2312";
    h2Style.Font.Size = FontSize_Normal;
    h2Style.Font.Bold = false;   // 取消加粗
    h2Style.ParagraphFormat.LineSpacingRule = wdLineSpaceExactly;
    h2Style.ParagraphFormat.LineSpacing = 28;
    h2Style.ParagraphFormat.SpaceBefore = 0;
    h2Style.ParagraphFormat.SpaceAfter = 0;
    h2Style.ParagraphFormat.SpaceBeforeAuto = false;
    h2Style.ParagraphFormat.SpaceAfterAuto = false;
    h2Style.ParagraphFormat.FirstLineIndent = CharactersToPoints(2);

    // 设置 三级标题
    var h3Style = doc.Styles.Item(wdStyleHeading3);
    h3Style.AutomaticallyUpdate = false;
    h3Style.Font.Name = "仿宋_GB2312";
    h3Style.Font.Size = FontSize_Normal;
    h3Style.Font.Bold = true;    // 加粗
    h3Style.ParagraphFormat.LineSpacingRule = wdLineSpaceExactly;
    h3Style.ParagraphFormat.LineSpacing = 28;
    h3Style.ParagraphFormat.SpaceBefore = 0;
    h3Style.ParagraphFormat.SpaceAfter = 0;
    h3Style.ParagraphFormat.SpaceBeforeAuto = false;
    h3Style.ParagraphFormat.SpaceAfterAuto = false;
    h3Style.ParagraphFormat.FirstLineIndent = CharactersToPoints(2);

    // 设置 四级标题
    var h4Style = doc.Styles.Item(wdStyleHeading4);
    h4Style.AutomaticallyUpdate = false;
    h4Style.Font.Name = "仿宋_GB2312";
    h4Style.Font.Size = FontSize_Normal;
    h4Style.Font.Bold = true;    // 加粗
    h4Style.ParagraphFormat.LineSpacingRule = wdLineSpaceExactly;
    h4Style.ParagraphFormat.LineSpacing = 28;
    h4Style.ParagraphFormat.SpaceBefore = 0;
    h4Style.ParagraphFormat.SpaceAfter = 0;
    h4Style.ParagraphFormat.SpaceBeforeAuto = false;
    h4Style.ParagraphFormat.SpaceAfterAuto = false;
    h4Style.ParagraphFormat.FirstLineIndent = CharactersToPoints(2);

    // 应用样式到全文段落（清除直接格式）
    for (var p = 1; p <= doc.Paragraphs.Count; p++) {
        var para = doc.Paragraphs.Item(p);
        if (para.Range.Text.trim().length > 1) {
            para.Range.Select();
            Application.Selection.ClearFormatting();

            var styleName = para.Style;
            if (styleName === "标题 1" || styleName === "Heading 1") {
                para.Style = wdStyleHeading1;
            } else if (styleName === "标题 2" || styleName === "Heading 2") {
                para.Style = wdStyleHeading2;
            } else if (styleName === "标题 3" || styleName === "Heading 3") {
                para.Style = wdStyleHeading3;
            } else if (styleName === "标题 4" || styleName === "Heading 4") {
                para.Style = wdStyleHeading4;
            } else {
                para.Style = wdStyleNormal;
            }
        }
    }

    // 自动匹配标题1和标题2（基于前缀）
    var level1Prefixes = ["一、","二、","三、","四、","五、","六、","七、","八、","九、","十、",
        "十一、","十二、","十三、","十四、","十五、","十六、","十七、","十八、","十九、","二十、"];
    var level2Prefixes = ["（一）","（二）","（三）","（四）","（五）","（六）","（七）","（八）","（九）","（十）",
        "（十一）","（十二）","（十三）","（十四）","（十五）","（十六）","（十七）","（十八）","（十九）","（二十）"];
    var level3Prefixes = ["1.","2.","3.","4.","5.","6.","7.","8.","9.","10.",
        "11.","12.","13.","14.","15.","16.","17.","18.","19.","20.",
        "21.","22.","23.","24.","25.","26.","27.","28.","29.","30."];
    var level4Prefixes = ["(1)","(2)","(3)","(4)","(5)","(6)","(7)","(8)","(9)","(10)",
        "(11)","(12)","(13)","(14)","(15)","(16)","(17)","(18)","(19)","(20)",
        "(21)","(22)","(23)","(24)","(25)","(26)","(27)","(28)","(29)","(30)"];
    var level4PrefixesChinese = ["（1）","（2）","（3）","（4）","（5）","（6）","（7）","（8）","（9）","（10）",
        "（11）","（12）","（13）","（14）","（15）","（16）","（17）","（18）","（19）","（20）",
        "（21）","（22）","（23）","（24）","（25）","（26）","（27）","（28）","（29）","（30）"];

    for (var p2 = 1; p2 <= doc.Paragraphs.Count; p2++) {
        var para2 = doc.Paragraphs.Item(p2);
        var text = para2.Range.Text;
        if (text.length <= 1) continue;  // 跳过空段落

        var trimmedText = text.slice(0, -1); // 去掉末尾的段落标记 (Chr(13))

        // 跳过纯空白段落
        if (trimmedText.trim() === "") continue;

        var matched = false;

        // 标题超过一行（段落行数大于1）则不判定为标题
        var lineCount = para2.Range.ComputeStatistics(wdStatisticLines);
        if (lineCount > 1) continue;

        // 检查是否为一级标题
        for (var i = 0; i < level1Prefixes.length; i++) {
            if (trimmedText.length >= level1Prefixes[i].length) {
                if (trimmedText.slice(0, level1Prefixes[i].length) === level1Prefixes[i]) {
                    para2.Style = wdStyleHeading1;
                    matched = true;
                    break;
                }
            }
        }

        if (!matched) {
            // 检查是否为二级标题
            for (var i = 0; i < level2Prefixes.length; i++) {
                if (trimmedText.length >= level2Prefixes[i].length) {
                    if (trimmedText.slice(0, level2Prefixes[i].length) === level2Prefixes[i]) {
                        para2.Style = wdStyleHeading2;
                        matched = true;
                        break;
                    }
                }
            }
        }

        if (!matched) {
            // 检查是否为三级标题（数字加点格式：1.、2.、3.等）
            for (var i = 0; i < level3Prefixes.length; i++) {
                if (trimmedText.length >= level3Prefixes[i].length) {
                    if (trimmedText.slice(0, level3Prefixes[i].length) === level3Prefixes[i]) {
                        para2.Style = wdStyleHeading3;
                        matched = true;
                        break;
                    }
                }
            }
        }

        if (!matched) {
            // 检查是否为四级标题（括号数字格式：(1)、(2)、(3)等）
            for (var i = 0; i < level4Prefixes.length; i++) {
                if (trimmedText.length >= level4Prefixes[i].length) {
                    if (trimmedText.slice(0, level4Prefixes[i].length) === level4Prefixes[i]) {
                        para2.Style = wdStyleHeading4;
                        var replaceRange = para2.Range;
                        replaceRange.End = replaceRange.End - 1;
                        replaceRange.Text = level4PrefixesChinese[i] + trimmedText.slice(level4Prefixes[i].length);
                        matched = true;
                        break;
                    }
                }
            }
        }

        // 如果都不是，则设为正文
        if (!matched) {
            para2.Style = wdStyleNormal;
        }
    }

    // 全文加粗特定前缀（一是、二是等）
    var boldPrefixes = ["一是","二是","三是","四是","五是","六是","七是","八是","九是","十是"];

    for (var b = 0; b < boldPrefixes.length; b++) {
        var boldPrefix = boldPrefixes[b];
        var find2 = doc.Content.Find;
        find2.ClearFormatting();                 // 清除查找格式
        find2.Replacement.ClearFormatting();     // 清除替换格式
        find2.Text = boldPrefix;                 // 查找内容（如"一是"、"二是"等）
        find2.Replacement.Text = boldPrefix;     // 替换内容（保持原文不变）
        find2.Replacement.Font.Bold = true;      // 替换后设置为加粗
        find2.Forward = true;                    // 向前搜索
        find2.Wrap = wdFindContinue;             // 搜索到结尾后继续从头开始
        find2.Format = true;                     // 启用格式替换
        find2.MatchCase = true;                  // 区分大小写
        find2.MatchWholeWord = false;            // 不要求全字匹配
        find2.MatchByte = true;                  // 区分全角/半角字符
        find2.Execute(null, null, null, null, null, null, null, null, null, null, wdReplaceAll);
    }

    // 格式化第一段为公文标题（方正小标宋简体 二号 居中）
    var firstPara = FindFirstNonEmptyParagraph(doc);

    if (firstPara !== null) {
        firstPara.Range.Font.Name = "方正小标宋简体";
        firstPara.Range.Font.Size = FontSize_Title1;
        firstPara.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter;
        firstPara.Range.ParagraphFormat.SpaceAfter = 0;
        firstPara.Range.ParagraphFormat.SpaceBefore = 0;
        firstPara.Range.ParagraphFormat.LineSpacingRule = wdLineSpaceExactly;
        firstPara.Range.ParagraphFormat.LineSpacing = 28;
        firstPara.Range.Font.Bold = false;

        // 在标题后插入一个空段落（即"空一行"）
        var nextRange = firstPara.Range.Duplicate;
        nextRange.Collapse(wdCollapseEnd);
        nextRange.InsertParagraphAfter();  // 插入一个空段落

        // 可选：确保这个空段落是 Normal 样式（符合正文规范）
        var emptyPara = nextRange.Paragraphs.Item(1);
        emptyPara.Style = wdStyleNormal;
        // 清除可能的直接格式
        emptyPara.Range.Font.Reset();
    }

    // 全部字体设为Times New Roman（因为此字体不包含中文，所以只会改变数字、英文的字体格式，中文字体保留原样）
    doc.Content.Font.Name = "Times New Roman";

    // 清除全文所有段落的自动编号
    for (var p3 = 1; p3 <= doc.Paragraphs.Count; p3++) {
        var finalPara = doc.Paragraphs.Item(p3);
        // 强制移除该段落的任何自动编号/列表格式
        finalPara.Range.ListFormat.RemoveNumbers(wdNumberAllNumbers);
    }

    // 设置所有图片为上下环绕
    // 处理嵌入型图片
    for (var is = 1; is <= doc.InlineShapes.Count; is++) {
        var inlineShape = doc.InlineShapes.Item(is);
        if (inlineShape.Type === wdInlineShapePicture) {
            inlineShape.ConvertToShape();
        }
    }

    for (var s = 1; s <= doc.Shapes.Count; s++) {
        var shape = doc.Shapes.Item(s);
        if (shape.Type === msoPicture) {
            // 确保不是嵌入型
            if (shape.WrapFormat.Type === wdWrapInline) {
                shape.WrapFormat.Type = wdWrapSquare;
            }
            shape.WrapFormat.Type = wdWrapTopBottom;
        }
    }

    // 处理表格：取消表格中所有段落的首行缩进
    for (var t = 1; t <= doc.Tables.Count; t++) {
        var tbl = doc.Tables.Item(t);
        for (var c = 1; c <= tbl.Range.Cells.Count; c++) {
            var cell = tbl.Range.Cells.Item(c);
            var rng = cell.Range;
            // 确保不包含单元格结束标记
            rng.End = rng.End - 1;
            for (var pp = 1; pp <= rng.Paragraphs.Count; pp++) {
                rng.Paragraphs.Item(pp).Range.ParagraphFormat.FirstLineIndent = 0;
            }
        }
    }

    // 替换全文空格：保留英文之间的空格，删除其他空格
    // 第一步：将英文之间的空格替换为临时标记
    var find3 = doc.Content.Find;
    find3.ClearFormatting();
    find3.Replacement.ClearFormatting();
    find3.Text = "([a-zA-Z]) ([a-zA-Z])";
    find3.Replacement.Text = "\\1§TEMP§\\2";
    find3.MatchWildcards = true;
    find3.Execute(null, null, null, null, null, null, null, null, null, null, wdReplaceAll);

    // 第二步：删除所有剩余空格
    var find4 = doc.Content.Find;
    find4.ClearFormatting();
    find4.Replacement.ClearFormatting();
    find4.Text = " ";
    find4.Replacement.Text = "";
    find4.MatchWildcards = false;
    find4.Execute(null, null, null, null, null, null, null, null, null, null, wdReplaceAll);

    // 第三步：将临时标记恢复为空格
    var find5 = doc.Content.Find;
    find5.ClearFormatting();
    find5.Replacement.ClearFormatting();
    find5.Text = "§TEMP§";
    find5.Replacement.Text = " ";
    find5.Execute(null, null, null, null, null, null, null, null, null, null, wdReplaceAll);

    // 设置页码：格式为"— 1 —"，奇数页右下角，偶数页左下角
    // 启用奇偶页不同页眉页脚
    doc.PageSetup.OddAndEvenPagesHeaderFooter = true;

    for (var secIdx = 1; secIdx <= doc.Sections.Count; secIdx++) {
        var sec = doc.Sections.Item(secIdx);

        // 奇数页页脚（右侧）
        var footerOdd = sec.Footers.Item(wdHeaderFooterPrimary);
        footerOdd.Range.Text = "";
        footerOdd.Range.Select();
        Application.Selection.TypeText("— ");
        Application.Selection.Fields.Add(Application.Selection.Range, wdFieldPage);
        Application.Selection.TypeText(" —");
        footerOdd.Range.ParagraphFormat.Alignment = wdAlignParagraphRight;
        footerOdd.Range.Font.Name = "宋体";
        footerOdd.Range.Font.Size = 14;

        // 偶数页页脚（左侧）
        var footerEven = sec.Footers.Item(wdHeaderFooterEvenPages);
        footerEven.Range.Text = "";
        footerEven.Range.Select();
        Application.Selection.TypeText("— ");
        Application.Selection.Fields.Add(Application.Selection.Range, wdFieldPage);
        Application.Selection.TypeText(" —");
        footerEven.Range.ParagraphFormat.Alignment = wdAlignParagraphLeft;
        footerEven.Range.Font.Name = "宋体";
        footerEven.Range.Font.Size = 14;
    }

    // 清除选择
    doc.UndoClear();
    doc.Range(0, 0).Select();

    MsgBox("文档已按照GB/T 9704-2012格式化完成。", 0x40);  // vbInformation = 0x40
}

// 辅助函数：查找第一个非空段落
function FindFirstNonEmptyParagraph(doc) {
    for (var p = 1; p <= doc.Paragraphs.Count; p++) {
        var para = doc.Paragraphs.Item(p);
        if (para.Range.Text.trim().length > 1) {
            return para;
        }
    }
    return null;
}

// 辅助函数：2字符缩进（32磅）
function CharactersToPoints(charCount) {
    return charCount * 16; // 1字符 = 16 磅
}