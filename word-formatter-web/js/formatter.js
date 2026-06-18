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
        
        // 2. 获取所有段落
        const paragraphs = this.handler.getParagraphs();
        
        // 3. 删除自动编号
        this._removeNumbering(paragraphs);
        
        // 4. 处理文本（标点替换、空格处理）
        this._processText(paragraphs);
        
        // 5. 识别标题并设置样式
        this._applyStyles(paragraphs);
        
        // 6. 设置首段标题格式
        this._formatFirstParagraph(paragraphs);
        
        // 7. 加粗特定前缀
        this._applyBoldPrefixes(paragraphs);
        
        // 8. 设置页面格式
        this._setPageFormat();
        
        // 9. 保存文档
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
     * 格式化首段标题
     * @param {Array} paragraphs - 段落数组
     */
    _formatFirstParagraph(paragraphs) {
        // 查找第一个非空段落
        let firstPara = null;
        for (const para of paragraphs) {
            if (para.text.trim().length > 1 && !para.hasImage) {
                firstPara = para;
                break;
            }
        }
        
        if (!firstPara) return;
        
        // 设置标题格式：方正小标宋简体、二号、居中
        this.handler.setParagraphFont(firstPara.index, CONFIG.fonts.title, CONFIG.fontSize.title1, false);
        this.handler.setParagraphAlignment(firstPara.index, 'center');
        this.handler.setLineSpacing(firstPara.index, CONFIG.lineSpacing);
        this.handler.setFirstLineIndent(firstPara.index, 0);
        
        // 检查标题后是否已有空段落
        const nextIndex = firstPara.index + 1;
        if (nextIndex < paragraphs.length) {
            const nextPara = paragraphs[nextIndex];
            if (nextPara.text.trim().length > 1) {
                // 这里需要插入空段落，但在纯前端环境中比较复杂
                // 暂时跳过，用户可以手动添加
                console.log('提示：标题后需要手动添加空行');
            }
        }
    }
    
    /**
     * 加粗特定前缀
     * @param {Array} paragraphs - 段落数组
     */
    _applyBoldPrefixes(paragraphs) {
        for (const para of paragraphs) {
            if (para.hasImage) continue;
            if (para.text.trim().length <= 1) continue;
            
            // 检查是否包含特定前缀
            if (containsAny(para.text, CONFIG.boldPrefixes)) {
                this.handler.setParagraphFont(para.index, null, null, true);
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
