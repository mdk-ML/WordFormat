' 将文档按照GB/T 9704-2012 规范格式化
' 牟雷
' 2026年02月05日 21:00
Sub Format2GBT9704_2012()
    Dim doc As Document
    Set doc = ActiveDocument

    ' 将所有手动换行符(软回车)替换为段落标记(硬回车)
    ' 关闭屏幕更新以加快速度并防止闪烁
    Application.ScreenUpdating = False
    With doc.Content.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        ' ^l 代表手动换行符 (Shift+Enter)
        .text = "^l"
        ' ^p 代表段落标记 (Enter)
        .Replacement.text = "^p"
        ' 执行全局替换
        .Execute Replace:=wdReplaceAll
    End With
    ' 恢复屏幕更新
    Application.ScreenUpdating = True

    ' 一次性转换全文自动编号为文本（保留递增编号）
    Dim docRange As Range
    Set docRange = doc.Content  ' 获取整个文档的范围
    ' 一次性转换所有自动编号为纯文本（保留原递增编号：1.、2.、3.）
    docRange.ListFormat.ConvertNumbersToText

    ' 清除全文所有段落的自动编号
    Dim clearPara As Paragraph
    For Each clearPara In doc.Paragraphs
        clearPara.Range.ListFormat.RemoveNumbers NumberType:=wdNumberAllNumbers
    Next clearPara

    ' 替换全文空格：保留英文之间的空格，删除其他空格
    With doc.Content.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .text = "([a-zA-Z]) ([a-zA-Z])"
        .Replacement.text = "\1§TEMP§\2"
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    With doc.Content.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .text = " "
        .Replacement.text = ""
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With
    With doc.Content.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .text = "§TEMP§"
        .Replacement.text = " "
        .Execute Replace:=wdReplaceAll
    End With

    ' 全文替换英文标点为中文标点
    ' 替换括号
    With doc.Content.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "("
        .Replacement.Text = "（"
        .Execute Replace:=wdReplaceAll
    End With
    With doc.Content.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = ")"
        .Replacement.Text = "）"
        .Execute Replace:=wdReplaceAll
    End With
    ' 替换逗号
    With doc.Content.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = ","
        .Replacement.Text = "，"
        .Execute Replace:=wdReplaceAll
    End With

    ' 替换冒号
    With doc.Content.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = ":"
        .Replacement.Text = "："
        .Execute Replace:=wdReplaceAll
    End With

    ' 替换问号
    With doc.Content.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "?"
        .Replacement.Text = "？"
        .Execute Replace:=wdReplaceAll
    End With

    ' 替换分号
    With doc.Content.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = ";"
        .Replacement.Text = "；"
        .Execute Replace:=wdReplaceAll
    End With

    ' 替换双引号（区分左右）
    Dim rngDouble As Range
    Set rngDouble = doc.Content
    Dim isFirstDouble As Boolean
    isFirstDouble = True

    ' 遍历替换双引号
    With rngDouble.Find
        .Text = """"
        .Forward = True
        .Wrap = wdFindStop
        Do While .Execute
            If isFirstDouble Then
                rngDouble.Text = "“"  ' 左双引号 "
            Else
                rngDouble.Text = "”"  ' 右双引号 "
            End If
            isFirstDouble = Not isFirstDouble
            rngDouble.Collapse wdCollapseEnd
        Loop
    End With

    ' 替换单引号（区分左右）
    Dim rngSingle As Range
    Set rngSingle = doc.Content
    Dim isFirstSingle As Boolean
    isFirstSingle = True

    ' 遍历替换单引号
    With rngSingle.Find
        .Text = "'"
        .Forward = True
        .Wrap = wdFindStop
        Do While .Execute
            If isFirstSingle Then
                rngSingle.Text = "‘"  ' 左单引号 '
            Else
                rngSingle.Text = "’"  ' 右单引号 '
            End If
            isFirstSingle = Not isFirstSingle
            rngSingle.Collapse wdCollapseEnd
        Loop
    End With

    ' 页面设置
    With doc.PageSetup
        .TopMargin = CentimetersToPoints(3.7)
        .BottomMargin = CentimetersToPoints(3.5)
        .LeftMargin = CentimetersToPoints(2.8)
        .RightMargin = CentimetersToPoints(2.6)
    End With

    ' 定义字号
    Const FontSize_Normal As Integer = 16   ' 三号 = 16 磅
    Const FontSize_Title1 As Integer = 22   ' 二号 ≈ 22 磅

    ' 设置 Normal 样式（正文）
    With doc.Styles(wdStyleNormal)
        With .Font
            .Name = "仿宋_GB2312"
            .Size = FontSize_Normal
        End With
        With .ParagraphFormat
            .LineSpacingRule = wdLineSpaceExactly
            .LineSpacing = 28
            .FirstLineIndent = CharactersToPoints(2)
            .SpaceBefore = 0
            .SpaceAfter = 0
        End With
    End With

    ' 设置 一级标题
    With doc.Styles(wdStyleHeading1)
        .AutomaticallyUpdate = False
        With .Font
            .Name = "黑体"
            .Size = FontSize_Normal
            .Bold = False   ' 取消加粗
        End With
        With .ParagraphFormat
            .LineSpacingRule = wdLineSpaceExactly
            .LineSpacing = 28
            .SpaceBefore = 0
            .SpaceAfter = 0
            .SpaceBeforeAuto = False   ' ← 禁用“自动”
            .SpaceAfterAuto = False    ' ← 禁用“自动”
            .FirstLineIndent = CharactersToPoints(2)
        End With
    End With

    ' 设置 二级标题
    With doc.Styles(wdStyleHeading2)
        .AutomaticallyUpdate = False
        With .Font
            .Name = "楷体_GB2312"
            .Size = FontSize_Normal
            .Bold = False   ' 取消加粗
        End With
        With .ParagraphFormat
            .LineSpacingRule = wdLineSpaceExactly
            .LineSpacing = 28
            .SpaceBefore = 0
            .SpaceAfter = 0
            .SpaceBeforeAuto = False   ' ← 禁用"自动"
            .SpaceAfterAuto = False    ' ← 禁用"自动"
            .FirstLineIndent = CharactersToPoints(2)
        End With
    End With

    ' 设置 三级标题
    With doc.Styles(wdStyleHeading3)
        .AutomaticallyUpdate = False
        With .Font
            .Name = "仿宋_GB2312"
            .Size = FontSize_Normal
            .Bold = True    ' 加粗
        End With
        With .ParagraphFormat
            .LineSpacingRule = wdLineSpaceExactly
            .LineSpacing = 28
            .SpaceBefore = 0
            .SpaceAfter = 0
            .SpaceBeforeAuto = False
            .SpaceAfterAuto = False
            .FirstLineIndent = CharactersToPoints(2)
        End With
    End With

    ' 设置 四级标题
    With doc.Styles(wdStyleHeading4)
        .AutomaticallyUpdate = False
        With .Font
            .Name = "仿宋_GB2312"
            .Size = FontSize_Normal
            .Bold = True    ' 加粗
        End With
        With .ParagraphFormat
            .LineSpacingRule = wdLineSpaceExactly
            .LineSpacing = 28
            .SpaceBefore = 0
            .SpaceAfter = 0
            .SpaceBeforeAuto = False
            .SpaceAfterAuto = False
            .FirstLineIndent = CharactersToPoints(2)
        End With
    End With

    ' 应用样式到全文段落（清除直接格式）
    Dim para As Paragraph
    For Each para In doc.Paragraphs
        If Len(Trim(para.Range.text)) > 1 Then
            ' 检查是否包含嵌入型图片
            If HasInlineImage(para) Then
                ' 图片段落特殊处理：使用单倍行距，不缩进
                para.Range.ParagraphFormat.LineSpacingRule = wdLineSpaceSingle
                para.Range.ParagraphFormat.FirstLineIndent = 0
            Else
                para.Range.Select
                Selection.ClearFormatting

                If para.Style = "标题 1" Or para.Style = "Heading 1" Then
                    para.Style = wdStyleHeading1
                ElseIf para.Style = "标题 2" Or para.Style = "Heading 2" Then
                    para.Style = wdStyleHeading2
                ElseIf para.Style = "标题 3" Or para.Style = "Heading 3" Then
                    para.Style = wdStyleHeading3
                ElseIf para.Style = "标题 4" Or para.Style = "Heading 4" Then
                    para.Style = wdStyleHeading4
                Else
                    para.Style = wdStyleNormal
                End If
            End If
        End If
    Next para

    ' 自动匹配标题1和标题2（基于前缀）
    Dim text As String
    Dim trimmedText As String

    ' 定义前缀
    Dim level1Prefixes As Variant
    Dim level2Prefixes As Variant
    Dim level3Prefixes As Variant
    Dim level4Prefixes As Variant
    level1Prefixes = Array("一、", "二、", "三、", "四、", "五、", "六、", "七、", "八、", "九、", "十、", "十一、", "十二、", "十三、", "十四、", "十五、", "十六、", "十七、", "十八、", "十九、", "二十、")
    level2Prefixes = Array("（一）", "（二）", "（三）", "（四）", "（五）", "（六）", "（七）", "（八）", "（九）", "（十）", "（十一）", "（十二）", "（十三）", "（十四）", "（十五）", "（十六）", "（十七）", "（十八）", "（十九）", "（二十）")
    level3Prefixes = Array("1.", "2.", "3.", "4.", "5.", "6.", "7.", "8.", "9.", "10.", "11.", "12.", "13.", "14.", "15.", "16.", "17.", "18.", "19.", "20.", "21.", "22.", "23.", "24.", "25.", "26.", "27.", "28.", "29.", "30.")
    level4Prefixes = Array("（1）", "（2）", "（3）", "（4）", "（5）", "（6）", "（7）", "（8）", "（9）", "（10）", "（11）", "（12）", "（13）", "（14）", "（15）", "（16）", "（17）", "（18）", "（19）", "（20）", "（21）", "（22）", "（23）", "（24）", "（25）", "（26）", "（27）", "（28）", "（29）", "（30）")
    

    Dim p As Long
    For p = 1 To doc.Paragraphs.Count
        Set para = doc.Paragraphs(p)
        text = para.Range.text
        If Len(text) <= 1 Then GoTo NextPara  ' 跳过空段落

        trimmedText = Left(text, Len(text) - 1) ' 去掉末尾的段落标记 (Chr(13))

        ' 跳过纯空白段落
        If Trim(trimmedText) = "" Then GoTo NextPara

        Dim i As Integer
        Dim matched As Boolean
        matched = False

        ' 标题超过一行（段落行数大于1）则不判定为标题
        Dim lineCount As Long
        lineCount = para.Range.ComputeStatistics(wdStatisticLines)
        If lineCount > 1 Then GoTo NextPara

        ' 检查是否为一级标题
        For i = 0 To UBound(level1Prefixes)
            If Len(trimmedText) >= Len(level1Prefixes(i)) Then
                If Left(trimmedText, Len(level1Prefixes(i))) = level1Prefixes(i) Then
                    para.Style = wdStyleHeading1
                    matched = True
                    Exit For
                End If
            End If
        Next i

        If Not matched Then
            ' 检查是否为二级标题
            For i = 0 To UBound(level2Prefixes)
                If Len(trimmedText) >= Len(level2Prefixes(i)) Then
                    If Left(trimmedText, Len(level2Prefixes(i))) = level2Prefixes(i) Then
                        para.Style = wdStyleHeading2
                        matched = True
                        Exit For
                    End If
                End If
            Next i
        End If

        If Not matched Then
            ' 检查是否为三级标题（数字加点格式：1.、2.、3.等）
            For i = 0 To UBound(level3Prefixes)
                If Len(trimmedText) >= Len(level3Prefixes(i)) Then
                    If Left(trimmedText, Len(level3Prefixes(i))) = level3Prefixes(i) Then
                        ' 检查下一段是否也是同级别前缀
                        Dim nextParaIndex3 As Long
                        nextParaIndex3 = p + 1
                        Dim isConsecutive3 As Boolean
                        isConsecutive3 = False
                        
                        ' 检查下一段（如果有）
                        If nextParaIndex3 <= doc.Paragraphs.Count Then
                            Dim nextPara3 As Paragraph
                            Set nextPara3 = doc.Paragraphs(nextParaIndex3)
                            Dim nextText3 As String
                            nextText3 = nextPara3.Range.Text
                            Dim nextTrimmed3 As String
                            nextTrimmed3 = Left(nextText3, Len(nextText3) - 1)
                            
                            ' 检查下一段是否以同级别前缀开头
                            Dim j3 As Integer
                            For j3 = 0 To UBound(level3Prefixes)
                                If Len(nextTrimmed3) >= Len(level3Prefixes(j3)) Then
                                    If Left(nextTrimmed3, Len(level3Prefixes(j3))) = level3Prefixes(j3) Then
                                        isConsecutive3 = True
                                        Exit For
                                    End If
                                End If
                            Next j3
                        End If
                        
                        ' 检查上一段（如果下一段不是同级别前缀）
                        If Not isConsecutive3 And p > 1 Then
                            Dim prevPara3 As Paragraph
                            Set prevPara3 = doc.Paragraphs(p - 1)
                            Dim prevText3 As String
                            prevText3 = prevPara3.Range.Text
                            Dim prevTrimmed3 As String
                            prevTrimmed3 = Left(prevText3, Len(prevText3) - 1)
                            
                            For j3 = 0 To UBound(level3Prefixes)
                                If Len(prevTrimmed3) >= Len(level3Prefixes(j3)) Then
                                    If Left(prevTrimmed3, Len(level3Prefixes(j3))) = level3Prefixes(j3) Then
                                        isConsecutive3 = True
                                        Exit For
                                    End If
                                End If
                            Next j3
                        End If
                        
                        ' 只有不连续时才判定为标题
                        If Not isConsecutive3 Then
                            para.Style = wdStyleHeading3
                            matched = True
                        End If
                        Exit For
                    End If
                End If
            Next i
        End If

        If Not matched Then
            ' 检查是否为四级标题（括号数字格式：(1)、(2)、(3)等）
            For i = 0 To UBound(level4Prefixes)
                If Len(trimmedText) >= Len(level4Prefixes(i)) Then
                    If Left(trimmedText, Len(level4Prefixes(i))) = level4Prefixes(i) Then
                        ' 检查下一段是否也是同级别前缀
                        Dim nextParaIndex4 As Long
                        nextParaIndex4 = p + 1
                        Dim isConsecutive4 As Boolean
                        isConsecutive4 = False
                        
                        ' 检查下一段（如果有）
                        If nextParaIndex4 <= doc.Paragraphs.Count Then
                            Dim nextPara4 As Paragraph
                            Set nextPara4 = doc.Paragraphs(nextParaIndex4)
                            Dim nextText4 As String
                            nextText4 = nextPara4.Range.Text
                            Dim nextTrimmed4 As String
                            nextTrimmed4 = Left(nextText4, Len(nextText4) - 1)
                            
                            ' 检查下一段是否以同级别前缀开头
                            Dim j4 As Integer
                            For j4 = 0 To UBound(level4Prefixes)
                                If Len(nextTrimmed4) >= Len(level4Prefixes(j4)) Then
                                    If Left(nextTrimmed4, Len(level4Prefixes(j4))) = level4Prefixes(j4) Then
                                        isConsecutive4 = True
                                        Exit For
                                    End If
                                End If
                            Next j4
                        End If
                        
                        ' 检查上一段（如果下一段不是同级别前缀）
                        If Not isConsecutive4 And p > 1 Then
                            Dim prevPara4 As Paragraph
                            Set prevPara4 = doc.Paragraphs(p - 1)
                            Dim prevText4 As String
                            prevText4 = prevPara4.Range.Text
                            Dim prevTrimmed4 As String
                            prevTrimmed4 = Left(prevText4, Len(prevText4) - 1)
                            
                            For j4 = 0 To UBound(level4Prefixes)
                                If Len(prevTrimmed4) >= Len(level4Prefixes(j4)) Then
                                    If Left(prevTrimmed4, Len(level4Prefixes(j4))) = level4Prefixes(j4) Then
                                        isConsecutive4 = True
                                        Exit For
                                    End If
                                End If
                            Next j4
                        End If
                        
                        ' 只有不连续时才判定为标题
                        If Not isConsecutive4 Then
                            para.Style = wdStyleHeading4
                            matched = True
                        End If
                        Exit For
                    End If
                End If
            Next i
        End If

        ' 如果都不是，则设为正文
        If Not matched Then
            para.Style = wdStyleNormal
        End If

NextPara:
    Next p

    ' 全文加粗特定前缀（一是、二是等）
    Dim boldPrefixes As Variant
    boldPrefixes = Array("一是", "二是", "三是", "四是", "五是", "六是", "七是", "八是", "九是", "十是")

    Dim boldPrefix As Variant
    For Each boldPrefix In boldPrefixes
        With doc.Content.Find
            .ClearFormatting                    ' 清除查找格式
            .Replacement.ClearFormatting        ' 清除替换格式
            .text = boldPrefix                  ' 查找内容（如"一是"、"二是"等）
            .Replacement.text = boldPrefix      ' 替换内容（保持原文不变）
            .Replacement.Font.Bold = True       ' 替换后设置为加粗
            .Forward = True                     ' 向前搜索
            .Wrap = wdFindContinue              ' 搜索到结尾后继续从头开始
            .Format = True                      ' 启用格式替换
            .MatchCase = True                   ' 区分大小写
            .MatchWholeWord = False             ' 不要求全字匹配
            .MatchByte = True                   ' 区分全角/半角字符
            .Execute Replace:=wdReplaceAll      ' 执行全部替换
        End With
    Next boldPrefix

    ' 格式化第一段为公文标题（方正小标宋简体 二号 居中）
    Dim titleRange As Range
    Set titleRange = FindFirstNonEmptyRange(doc)

    If Not titleRange Is Nothing Then
        With titleRange
            .Font.Name = "方正小标宋简体"
            .Font.Size = FontSize_Title1
            .ParagraphFormat.Alignment = wdAlignParagraphCenter
            .ParagraphFormat.SpaceAfter = 0
            .ParagraphFormat.SpaceBefore = 0
            .ParagraphFormat.LineSpacingRule = wdLineSpaceExactly
            .ParagraphFormat.LineSpacing = 28
            .ParagraphFormat.FirstLineIndent = 0
            .Font.Bold = False
        End With

        ' 在标题后插入一个空段落（即"空一行"）
        ' 检查标题范围是否以空行结尾
        Dim needInsert As Boolean
        needInsert = True
        
        ' 获取 titleRange 之后的下一个段落
        Dim afterRange As Range
        Set afterRange = doc.Range(titleRange.End, doc.Range.End)
        If afterRange.Paragraphs.Count > 0 Then
            Dim nextPara As Paragraph
            Set nextPara = afterRange.Paragraphs(1)
            ' 检查下一段是否为空段落（只有段落标记或空白）
            If Len(Trim(nextPara.Range.Text)) <= 1 Then
                needInsert = False
            End If
        End If
        
        ' 只有标题后没有空段落时才插入
        If needInsert Then
            Dim nextRange As Range
            Set nextRange = titleRange.Duplicate
            nextRange.Collapse Direction:=wdCollapseEnd
            nextRange.InsertParagraphAfter  ' 插入一个空段落

            ' 可选：确保这个空段落是 Normal 样式（符合正文规范）
            Dim emptyPara As Paragraph
            Set emptyPara = nextRange.Paragraphs(1)
            emptyPara.Style = wdStyleNormal
            ' 清除可能的直接格式
            emptyPara.Range.Font.Reset
        End If
    End If


    ' 全部字体设为Times New Roman（因为此字体不包含中文，所以只会改变数字、英文的字体格式，中文字体保留原样）
    With doc.Range
        .Font.Name = "Times New Roman"
    End With

    ' 处理表格：取消表格中所有段落的首行缩进
    Dim tbl As Table
    Dim cell As cell
    Dim rng As Range
    For Each tbl In doc.Tables
        For Each cell In tbl.Range.Cells
            Set rng = cell.Range
            ' 确保不包含单元格结束标记
            rng.End = rng.End - 1
            For Each para In rng.Paragraphs
                para.Range.ParagraphFormat.FirstLineIndent = 0
            Next para
        Next cell
    Next tbl

    ' 设置页码：格式为"— 1 —"，奇数页右下角，偶数页左下角
    Dim sec As Section
    Dim footerEven As HeaderFooter
    Dim footerOdd As HeaderFooter

    ' 启用奇偶页不同页眉页脚
    doc.PageSetup.OddAndEvenPagesHeaderFooter = True

    For Each sec In doc.Sections
        ' 奇数页页脚（右侧）
        Set footerOdd = sec.Footers(wdHeaderFooterPrimary)
        footerOdd.Range.Text = ""
        footerOdd.Range.Select
        Selection.TypeText "— "
        Selection.Fields.Add Selection.Range, wdFieldPage
        Selection.TypeText " —"
        footerOdd.Range.ParagraphFormat.Alignment = wdAlignParagraphRight
        With footerOdd.Range.Font
            .Name = "宋体"
            .Size = 14
        End With

        ' 偶数页页脚（左侧）
        Set footerEven = sec.Footers(wdHeaderFooterEvenPages)
        footerEven.Range.Text = ""
        footerEven.Range.Select
        Selection.TypeText "— "
        Selection.Fields.Add Selection.Range, wdFieldPage
        Selection.TypeText " —"
        footerEven.Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
        With footerEven.Range.Font
            .Name = "宋体"
            .Size = 14
        End With
    Next sec

    ' 清除选择
    doc.UndoClear
    doc.Range(0, 0).Select

    MsgBox "文档已按照GB/T 9704-2012格式化完成。", vbInformation
End Sub

' 辅助函数：查找第一个非空段落范围（从第一个非空段落到遇到空白段落之前的所有连续段落，最多5个段落）
Function FindFirstNonEmptyRange(doc As Document) As Range
    Dim startPara As Paragraph
    Dim lastNonEmptyPara As Paragraph
    Dim para As Paragraph
    Dim paraCount As Integer
    
    paraCount = 0
    
    For Each para In doc.Paragraphs
        If Len(Trim(para.Range.text)) > 1 Then
            If startPara Is Nothing Then
                Set startPara = para
            End If
            paraCount = paraCount + 1
            If paraCount <= 5 Then
                Set lastNonEmptyPara = para
            End If
        Else
            ' 遇到空白段落，且已经找到了起始段落
            If Not startPara Is Nothing Then
                Exit For
            End If
        End If
    Next para
    
    If startPara Is Nothing Then
        Set FindFirstNonEmptyRange = Nothing
    ElseIf paraCount > 5 Then
        ' 超过5个段落，只返回第一个段落
        Set FindFirstNonEmptyRange = startPara.Range
    Else
        Set FindFirstNonEmptyRange = doc.Range(startPara.Range.Start, lastNonEmptyPara.Range.End)
    End If
End Function

' 辅助函数：2字符缩进（32磅）
Function CharactersToPoints(charCount As Integer) As Single
    CharactersToPoints = charCount * 16 ' 1字符 = 16 磅
End Function

' 辅助函数：检查段落是否包含嵌入型图片
Function HasInlineImage(para As Paragraph) As Boolean
    Dim ilShape As InlineShape
    For Each ilShape In para.Range.InlineShapes
        If ilShape.Type = wdInlineShapePicture Then
            HasInlineImage = True
            Exit Function
        End If
    Next ilShape
    HasInlineImage = False
End Function
