/**
 * 文档处理器
 * 基于jszip和DOMParser操作.docx文件
 */

class DocxHandler {
    constructor() {
        this.zip = null;
        this.documentXml = null;
        this.xmlDoc = null;
        this.paragraphs = [];
    }
    
    /**
     * 加载.docx文件
     * @param {File} file - 文件对象
     * @returns {DocxHandler} - 返回自身以支持链式调用
     */
    async load(file) {
        const arrayBuffer = await file.arrayBuffer();
        this.zip = await JSZip.loadAsync(arrayBuffer);
        
        // 读取document.xml
        const documentXmlFile = this.zip.file('word/document.xml');
        if (!documentXmlFile) {
            throw new Error('无效的.docx文件：缺少document.xml');
        }
        
        this.documentXml = await documentXmlFile.async('string');
        
        // 解析XML
        const parser = new DOMParser();
        this.xmlDoc = parser.parseFromString(this.documentXml, 'text/xml');
        
        // 提取段落
        this._extractParagraphs();
        
        return this;
    }
    
    /**
     * 提取段落
     */
    _extractParagraphs() {
        this.paragraphs = [];
        const body = this.xmlDoc.getElementsByTagName('w:body')[0];
        if (!body) return;
        
        const pElements = body.getElementsByTagName('w:p');
        
        for (let i = 0; i < pElements.length; i++) {
            const pElement = pElements[i];
            const text = this._extractText(pElement);
            const xml = new XMLSerializer().serializeToString(pElement);
            
            this.paragraphs.push({
                index: i,
                text: text,
                hasImage: this._hasImage(xml),
                element: pElement
            });
        }
    }
    
    /**
     * 提取段落文本
     * @param {Element} pElement - 段落XML元素
     * @returns {string} - 段落文本
     */
    _extractText(pElement) {
        let text = '';
        const tElements = pElement.getElementsByTagName('w:t');
        
        for (let i = 0; i < tElements.length; i++) {
            text += tElements[i].textContent || '';
        }
        
        return text;
    }
    
    /**
     * 检查是否包含图片
     * @param {string} xml - XML字符串
     * @returns {boolean} - 是否包含图片
     */
    _hasImage(xml) {
        return xml && (
            xml.includes('<w:drawing') || 
            xml.includes('<w:pict') ||
            xml.includes('<v:imagedata')
        );
    }
    
    /**
     * 获取所有段落
     * @returns {Array} - 段落数组
     */
    getParagraphs() {
        return this.paragraphs;
    }
    
    /**
     * 设置段落文本
     * @param {number} index - 段落索引
     * @param {string} newText - 新文本
     */
    setParagraphText(index, newText) {
        if (index >= this.paragraphs.length) return;
        
        const pElement = this.paragraphs[index].element;
        
        // 获取所有w:r元素
        const rElements = pElement.getElementsByTagName('w:r');
        
        if (rElements.length > 0) {
            const firstR = rElements[0];
            
            // 清除第一个run中的所有w:t元素
            const tElements = firstR.getElementsByTagName('w:t');
            while (tElements.length > 0) {
                tElements[0].parentNode.removeChild(tElements[0]);
            }
            
            // 添加新的w:t元素
            const newT = this.xmlDoc.createElement('w:t');
            newT.setAttribute('xml:space', 'preserve');
            newT.textContent = newText;
            firstR.appendChild(newT);
            
            // 删除其他run
            for (let i = rElements.length - 1; i > 0; i--) {
                rElements[i].parentNode.removeChild(rElements[i]);
            }
        }
    }
    
    /**
     * 设置段落样式
     * @param {number} index - 段落索引
     * @param {string} styleName - 样式名称
     */
    setParagraphStyle(index, styleName) {
        if (index >= this.paragraphs.length) return;
        
        const pElement = this.paragraphs[index].element;
        
        // 获取或创建w:pPr元素
        let pPr = pElement.getElementsByTagName('w:pPr')[0];
        if (!pPr) {
            pPr = this.xmlDoc.createElement('w:pPr');
            pElement.insertBefore(pPr, pElement.firstChild);
        }
        
        // 获取或创建w:pStyle元素
        let pStyle = pPr.getElementsByTagName('w:pStyle')[0];
        if (!pStyle) {
            pStyle = this.xmlDoc.createElement('w:pStyle');
            pPr.insertBefore(pStyle, pPr.firstChild);
        }
        
        // 设置样式值
        pStyle.setAttribute('w:val', styleName);
    }
    
    /**
     * 设置段落字体
     * @param {number} index - 段落索引
     * @param {string} fontName - 字体名称
     * @param {number} fontSize - 字号（磅）
     * @param {boolean} bold - 是否加粗
     */
    setParagraphFont(index, fontName, fontSize, bold) {
        if (index >= this.paragraphs.length) return;
        
        const pElement = this.paragraphs[index].element;
        const rElements = pElement.getElementsByTagName('w:r');
        
        for (let i = 0; i < rElements.length; i++) {
            const rElement = rElements[i];
            
            // 跳过包含图片的run
            if (rElement.getElementsByTagName('w:drawing').length > 0 || 
                rElement.getElementsByTagName('v:shape').length > 0) {
                continue;
            }
            
            // 获取或创建w:rPr元素
            let rPr = rElement.getElementsByTagName('w:rPr')[0];
            if (!rPr) {
                rPr = this.xmlDoc.createElement('w:rPr');
                rElement.insertBefore(rPr, rElement.firstChild);
            }
            
            // 设置字体
            if (fontName) {
                let rFonts = rPr.getElementsByTagName('w:rFonts')[0];
                if (!rFonts) {
                    rFonts = this.xmlDoc.createElement('w:rFonts');
                    rPr.insertBefore(rFonts, rPr.firstChild);
                }
                rFonts.setAttribute('w:ascii', fontName);
                rFonts.setAttribute('w:eastAsia', fontName);
                rFonts.setAttribute('w:hAnsi', fontName);
            }
            
            // 设置字号
            if (fontSize) {
                let sz = rPr.getElementsByTagName('w:sz')[0];
                if (!sz) {
                    sz = this.xmlDoc.createElement('w:sz');
                    rPr.appendChild(sz);
                }
                // Word中字号单位是半磅
                sz.setAttribute('w:val', (fontSize * 2).toString());
                
                let szCs = rPr.getElementsByTagName('w:szCs')[0];
                if (!szCs) {
                    szCs = this.xmlDoc.createElement('w:szCs');
                    rPr.appendChild(szCs);
                }
                szCs.setAttribute('w:val', (fontSize * 2).toString());
            }
            
            // 设置加粗
            if (bold !== undefined) {
                if (bold) {
                    if (!rPr.getElementsByTagName('w:b')[0]) {
                        const b = this.xmlDoc.createElement('w:b');
                        rPr.appendChild(b);
                    }
                    if (!rPr.getElementsByTagName('w:bCs')[0]) {
                        const bCs = this.xmlDoc.createElement('w:bCs');
                        rPr.appendChild(bCs);
                    }
                } else {
                    const b = rPr.getElementsByTagName('w:b')[0];
                    if (b) b.parentNode.removeChild(b);
                    const bCs = rPr.getElementsByTagName('w:bCs')[0];
                    if (bCs) bCs.parentNode.removeChild(bCs);
                }
            }
        }
    }
    
    /**
     * 设置页面边距
     * @param {number} top - 上边距（厘米）
     * @param {number} bottom - 下边距（厘米）
     * @param {number} left - 左边距（厘米）
     * @param {number} right - 右边距（厘米）
     */
    setPageMargin(top, bottom, left, right) {
        const body = this.xmlDoc.getElementsByTagName('w:body')[0];
        if (!body) return;
        
        // 获取或创建w:sectPr元素
        let sectPr = body.getElementsByTagName('w:sectPr')[0];
        if (!sectPr) {
            sectPr = this.xmlDoc.createElement('w:sectPr');
            body.appendChild(sectPr);
        }
        
        // 获取或创建w:pgMar元素
        let pgMar = sectPr.getElementsByTagName('w:pgMar')[0];
        if (!pgMar) {
            pgMar = this.xmlDoc.createElement('w:pgMar');
            sectPr.appendChild(pgMar);
        }
        
        // 设置边距（单位：twips，1厘米=567twips）
        pgMar.setAttribute('w:top', cmToTwips(top).toString());
        pgMar.setAttribute('w:bottom', cmToTwips(bottom).toString());
        pgMar.setAttribute('w:left', cmToTwips(left).toString());
        pgMar.setAttribute('w:right', cmToTwips(right).toString());
    }
    
    /**
     * 设置行距
     * @param {number} index - 段落索引
     * @param {number} lineSpacing - 行距（磅）
     */
    setLineSpacing(index, lineSpacing) {
        if (index >= this.paragraphs.length) return;
        
        const pElement = this.paragraphs[index].element;
        
        // 获取或创建w:pPr元素
        let pPr = pElement.getElementsByTagName('w:pPr')[0];
        if (!pPr) {
            pPr = this.xmlDoc.createElement('w:pPr');
            pElement.insertBefore(pPr, pElement.firstChild);
        }
        
        // 获取或创建w:spacing元素
        let spacing = pPr.getElementsByTagName('w:spacing')[0];
        if (!spacing) {
            spacing = this.xmlDoc.createElement('w:spacing');
            pPr.appendChild(spacing);
        }
        
        // 设置行距（单位：twips的1/20，1磅=20twips）
        spacing.setAttribute('w:line', ptToTwips(lineSpacing).toString());
        spacing.setAttribute('w:lineRule', 'exact');
        
        // 设置段前段后间距为0
        spacing.setAttribute('w:before', '0');
        spacing.setAttribute('w:after', '0');
    }
    
    /**
     * 设置首行缩进
     * @param {number} index - 段落索引
     * @param {number} indentPoints - 缩进量（磅）
     */
    setFirstLineIndent(index, indentPoints) {
        if (index >= this.paragraphs.length) return;
        
        const pElement = this.paragraphs[index].element;
        
        // 获取或创建w:pPr元素
        let pPr = pElement.getElementsByTagName('w:pPr')[0];
        if (!pPr) {
            pPr = this.xmlDoc.createElement('w:pPr');
            pElement.insertBefore(pPr, pElement.firstChild);
        }
        
        // 获取或创建w:ind元素
        let ind = pPr.getElementsByTagName('w:ind')[0];
        if (!ind) {
            ind = this.xmlDoc.createElement('w:ind');
            pPr.appendChild(ind);
        }
        
        // 设置首行缩进（单位：twips，1磅=20twips）
        ind.setAttribute('w:firstLine', ptToTwips(indentPoints).toString());
    }
    
    /**
     * 设置段落对齐方式
     * @param {number} index - 段落索引
     * @param {string} alignment - 对齐方式（left, center, right, both）
     */
    setParagraphAlignment(index, alignment) {
        if (index >= this.paragraphs.length) return;
        
        const pElement = this.paragraphs[index].element;
        
        // 获取或创建w:pPr元素
        let pPr = pElement.getElementsByTagName('w:pPr')[0];
        if (!pPr) {
            pPr = this.xmlDoc.createElement('w:pPr');
            pElement.insertBefore(pPr, pElement.firstChild);
        }
        
        // 获取或创建w:jc元素
        let jc = pPr.getElementsByTagName('w:jc')[0];
        if (!jc) {
            jc = this.xmlDoc.createElement('w:jc');
            pPr.appendChild(jc);
        }
        
        // 设置对齐方式
        jc.setAttribute('w:val', alignment);
    }
    
    /**
     * 删除自动编号
     * @param {number} index - 段落索引
     */
    removeNumbering(index) {
        if (index >= this.paragraphs.length) return;
        
        const pElement = this.paragraphs[index].element;
        
        // 获取w:pPr元素
        const pPr = pElement.getElementsByTagName('w:pPr')[0];
        if (!pPr) return;
        
        // 删除w:numPr元素
        const numPr = pPr.getElementsByTagName('w:numPr')[0];
        if (numPr) {
            numPr.parentNode.removeChild(numPr);
        }
    }
    
    /**
     * 获取样式XML
     * @returns {string|null} - 样式XML字符串
     */
    async getStylesXml() {
        const stylesFile = this.zip.file('word/styles.xml');
        if (!stylesFile) return null;
        return await stylesFile.async('string');
    }
    
    /**
     * 保存文档
     * @returns {Blob} - 文档Blob
     */
    async save() {
        // 序列化XML
        const serializer = new XMLSerializer();
        const newDocumentXml = serializer.serializeToString(this.xmlDoc);
        
        // 更新zip中的document.xml
        this.zip.file('word/document.xml', newDocumentXml);
        
        // 生成新的.docx文件
        return await this.zip.generateAsync({ type: 'blob' });
    }
    
    /**
     * 获取图片
     * @returns {Object} - 图片映射
     */
    async getImages() {
        const images = {};
        
        // 遍历word/media目录
        const mediaFiles = Object.keys(this.zip.files).filter(
            name => name.startsWith('word/media/')
        );
        
        for (const fileName of mediaFiles) {
            const data = await this.zip.file(fileName).async('base64');
            const ext = fileName.split('.').pop().toLowerCase();
            let mimeType = 'image/png';
            
            if (ext === 'jpg' || ext === 'jpeg') {
                mimeType = 'image/jpeg';
            } else if (ext === 'gif') {
                mimeType = 'image/gif';
            } else if (ext === 'bmp') {
                mimeType = 'image/bmp';
            }
            
            images[fileName] = `data:${mimeType};base64,${data}`;
        }
        
        return images;
    }
}
