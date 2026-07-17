/**
 * 格式化器
 * 将Format2GBT9704_2012.js的逻辑适配到Web环境
 */

class Formatter {
    constructor() {
        this.handler = new DocxHandler();
    }
    
    /**
     * 格式化文档
     * @param {File} file - 文件对象
     * @returns {Blob} - 格式化后的文档Blob
     */
    async format(file) {
        // 1. 加载文档
        await this.handler.load(file);
        
        // 2. 软回车→硬回车转换（对应原脚本 ^l → ^p）
        this.handler.convertSoftReturnsToParagraphs();
        
        // 3. 获取所有段落
        let paragraphs = this.handler.getParagraphs();
        
        // 4. 更新样式定义（Normal/Heading1-4）
        await this.handler.updateStyleDefinitions();
        
        // 5. 删除自动编号
        this._removeNumbering(paragraphs);
        
        // 6. 处理文本（标点替换、空格处理）
        this._processText(paragraphs);
        
        // 7. 识别标题并设置样式
        this._applyStyles(paragraphs);
        
        // 8. 设置首段标题格式（可能插入空段落，需要重新获取段落列表）
        this._formatFirstParagraph(paragraphs);
        paragraphs = this.handler.getParagraphs(); // 刷新引用
        
        // 9. 加粗特定前缀
        this._applyBoldPrefixes(paragraphs);
        
        // 10. 设置全文档 Times New Roman 字体（仅影响数字和英文）
        this.handler.setAllFontsTimesNewRoman();
        
        // 11. 取消表格中所有段落的首行缩进
        this.handler.removeTableIndent();
        
        // 12. 设置页码（奇偶页不同，"— 1 —"格式）
        await this.handler.setPageNumber();
        
        // 13. 设置页面格式
        this._setPageFormat();
        
        // 14. 保存文档
        return await this.handler.save();
    }
    
    /**
     * 删除自动编号
     * @param {Array} paragraphs - 段落数组
     */
    _removeNumbering(paragraphs) {
        for (const para of paragraphs) {
            this.handler.removeNumbering(para.index);
        }
    }
    
    /**
     * 处理文本
     * @param {Array} paragraphs - 段落数组
     */
    _processText(paragraphs) {
        for (const para of paragraphs) {
            // 跳过图片段落
            if (para.hasImage) continue;
            
            let text = para.text;
            
            // 空格处理
            text = processSpaces(text);
            
            // 标点符号替换
            text = replacePunctuation(text);
            
            // 更新段落文本
            this.handler.setParagraphText(para.index, text);
        }
    }
    
    /**
     * 应用样式
     * @param {Array} paragraphs - 段落数组
     */
    _applyStyles(paragraphs) {
        for (const para of paragraphs) {
            // 图片段落特殊处理
            if (para.hasImage) {
                this.handler.setLineSpacing(para.index, 0); // 单倍行距
                this.handler.setFirstLineIndent(para.index, 0); // 不缩进
                continue;
            }
            
            // 空段落跳过
            if (para.text.trim().length <= 1) continue;
            
            // 识别标题级别
            const heading = identifyHeadingLevel(para.text, para.index, paragraphs);
            
            if (heading) {
                // 设置标题样式
                switch (heading.level) {
                    case 1:
                        this.handler.setParagraphStyle(para.index, 'Heading1');
                        this.handler.setParagraphFont(para.index, CONFIG.fonts.h1, CONFIG.fontSize.normal, false);
                        break;
                    case 2:
                        this.handler.setParagraphStyle(para.index, 'Heading2');
                        this.handler.setParagraphFont(para.index, CONFIG.fonts.h2, CONFIG.fontSize.normal, false);
                        break;
                    case 3:
                        this.handler.setParagraphStyle(para.index, 'Heading3');
                        this.handler.setParagraphFont(para.index, CONFIG.fonts.h3, CONFIG.fontSize.normal, true);
                        break;
                    case 4:
                        this.handler.setParagraphStyle(para.index, 'Heading4');
                        this.handler.setParagraphFont(para.index, CONFIG.fonts.h4, CONFIG.fontSize.normal, true);
                        break;
                }
                
                // 设置行距和缩进
                this.handler.setLineSpacing(para.index, CONFIG.lineSpacing);
                this.handler.setFirstLineIndent(para.index, charactersToPoints(2));
            } else {
                // 正文样式
                this.handler.setParagraphStyle(para.index, 'Normal');
                this.handler.setParagraphFont(para.index, CONFIG.fonts.body, CONFIG.fontSize.normal, false);
                this.handler.setLineSpacing(para.index, CONFIG.lineSpacing);
                this.handler.setFirstLineIndent(para.index, charactersToPoints(2));
            }
        }
    }
    
    /**
     * 格式化首段标题（对应 FindFirstNonEmptyRange 逻辑）
     * 收集从第一个非空段落到遇到空白段落之前的所有连续段落（最多5段），应用标题格式
     * @param {Array} paragraphs - 段落数组
     */
    _formatFirstParagraph(paragraphs) {
        // 查找第一个非空段落
        let startIndex = -1;
        for (let i = 0; i < paragraphs.length; i++) {
            const para = paragraphs[i];
            if (para.text.trim().length > 1 && !para.hasImage) {
                startIndex = i;
                break;
            }
        }
        
        if (startIndex === -1) return;
        
        // 收集所有连续的非空段落到遇到空白段落为止（不设上限，收集完成后判断）
        const titleIndices = [startIndex];
        for (let i = startIndex + 1; i < paragraphs.length; i++) {
            const para = paragraphs[i];
            if (para.text.trim().length <= 1 || para.hasImage) {
                break; // 遇到空白段落或图片段落，停止收集
            }
            titleIndices.push(i);
        }
        
        // 确定要格式化的段落索引列表：超过5段则只取第一段
        let formatIndices;
        if (titleIndices.length > 5) {
            formatIndices = [startIndex];
        } else {
            formatIndices = titleIndices;
        }
        
        // 应用标题格式到所有收集的段落
        for (const idx of formatIndices) {
            this.handler.setParagraphFont(idx, CONFIG.fonts.title, CONFIG.fontSize.title1, false);
            this.handler.setParagraphAlignment(idx, 'center');
            this.handler.setLineSpacing(idx, CONFIG.lineSpacing);
            this.handler.setFirstLineIndent(idx, 0);
        }
        
        // 检查标题后是否已有空段落
        const lastTitleIndex = formatIndices[formatIndices.length - 1];
        const nextIndex = lastTitleIndex + 1;
        
        let needInsert = true;
        if (nextIndex < paragraphs.length) {
            const nextPara = paragraphs[nextIndex];
            if (nextPara.text.trim().length <= 1) {
                needInsert = false; // 后面已经是空段落
            }
        }
        
        // 只有标题后没有空段落时才插入
        if (needInsert) {
            this.handler.insertEmptyParagraphAfter(lastTitleIndex);
            console.log('标题后已插入空行');
        }
    }
    
    /**
     * 加粗特定前缀（仅加粗匹配的前缀文本本身，而非整个段落）
     * @param {Array} paragraphs - 段落数组
     */
    _applyBoldPrefixes(paragraphs) {
        for (const para of paragraphs) {
            if (para.hasImage) continue;
            if (para.text.trim().length <= 1) continue;
            
            const text = para.text;
            let hasMatch = false;
            for (const prefix of CONFIG.boldPrefixes) {
                if (text.includes(prefix)) {
                    hasMatch = true;
                    break;
                }
            }
            
            if (hasMatch) {
                this.handler.setBoldPrefixes(para.index, CONFIG.boldPrefixes);
            }
        }
    }
    
    /**
     * 设置页面格式
     */
    _setPageFormat() {
        this.handler.setPageMargin(
            CONFIG.pageMargin.top,
            CONFIG.pageMargin.bottom,
            CONFIG.pageMargin.left,
            CONFIG.pageMargin.right
        );
    }
}
