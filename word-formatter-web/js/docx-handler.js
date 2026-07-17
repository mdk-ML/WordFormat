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
     * 将手动换行符（软回车 <w:br/>）转换为段落标记（硬回车 <w:p>）
     * 对应 Format2GBT9704_2012.js 中 "^l" → "^p" 的替换
     */
    convertSoftReturnsToParagraphs() {
        const body = this.xmlDoc.getElementsByTagName('w:body')[0];
        if (!body) return;
        
        const pElements = body.getElementsByTagName('w:p');
        const paragraphsToProcess = [];
        for (let i = 0; i < pElements.length; i++) {
            paragraphsToProcess.push(pElements[i]);
        }
        
        for (const pElement of paragraphsToProcess) {
            const brElements = pElement.querySelectorAll('w\\:br, br');
            if (brElements.length === 0) continue;
            
            // 过滤：只处理手动换行（非分页符/分栏符）
            let hasSoftBreak = false;
            for (const br of brElements) {
                const type = br.getAttribute('w:type');
                if (!type || type === 'textWrapping') {
                    hasSoftBreak = true;
                    break;
                }
            }
            if (!hasSoftBreak) continue;
            
            // 克隆节点，用于拆分为多个段落
            const pClone = pElement.cloneNode(true);
            const parent = pElement.parentNode;
            const nextSibling = pElement.nextSibling;
            
            // 收集所有 run 的文本 + br 的位置
            const segments = [];
            let currentSegment = [];
            
            const allChildren = Array.from(pClone.childNodes);
            for (const child of allChildren) {
                const tagName = child.nodeName ? child.nodeName.toLowerCase() : '';
                if (tagName === 'w:br' || child.nodeName === 'br') {
                    const type = child.getAttribute ? child.getAttribute('w:type') : null;
                    if (!type || type === 'textWrapping') {
                        // 软回车：结束当前段落，开始新段落
                        if (currentSegment.length > 0) {
                            segments.push(currentSegment);
                            currentSegment = [];
                        }
                        child.parentNode.removeChild(child);
                        continue;
                    }
                }
                currentSegment.push(child);
            }
            if (currentSegment.length > 0) {
                segments.push(currentSegment);
            }
            
            if (segments.length <= 1) continue; // 没有有效拆分
            
            // 清空原段落并用第一段内容填充
            while (pElement.firstChild) {
                pElement.removeChild(pElement.firstChild);
            }
            for (const node of segments[0]) {
                pElement.appendChild(node.cloneNode(true));
            }
            
            // 插入后续段落
            for (let s = 1; s < segments.length; s++) {
                const newP = this.xmlDoc.createElement('w:p');
                // 复制段落属性（如果有）
                for (const node of segments[s]) {
                    newP.appendChild(node.cloneNode(true));
                }
                parent.insertBefore(newP, nextSibling);
            }
        }
        
        // 重新提取段落
        this._extractParagraphs();
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
     * 在指定段落之后插入一个空段落
     * @param {number} index - 段落索引
     */
    insertEmptyParagraphAfter(index) {
        if (index >= this.paragraphs.length) return null;
        
        const pElement = this.paragraphs[index].element;
        const parent = pElement.parentNode;
        const nextSibling = pElement.nextSibling;
        
        // 创建新段落
        const newP = this.xmlDoc.createElement('w:p');
        const pPr = this.xmlDoc.createElement('w:pPr');
        newP.appendChild(pPr);
        
        // 添加段落样式为 Normal
        const pStyle = this.xmlDoc.createElement('w:pStyle');
        pStyle.setAttribute('w:val', 'Normal');
        pPr.appendChild(pStyle);
        
        // 设置固定行距28磅
        const spacing = this.xmlDoc.createElement('w:spacing');
        spacing.setAttribute('w:line', '560');   // 28磅 * 20 = 560 twips
        spacing.setAttribute('w:lineRule', 'exact');
        spacing.setAttribute('w:before', '0');
        spacing.setAttribute('w:after', '0');
        pPr.appendChild(spacing);
        
        // 添加一个空的 run
        const r = this.xmlDoc.createElement('w:r');
        const t = this.xmlDoc.createElement('w:t');
        t.setAttribute('xml:space', 'preserve');
        t.textContent = '';
        r.appendChild(t);
        newP.appendChild(r);
        
        parent.insertBefore(newP, nextSibling);
        
        // 更新段落缓存
        this._extractParagraphs();
        
        return newP;
    }
    
    /**
     * 设置全部字体为Times New Roman（仅影响数字和英文，中文字体保留原样）
     * 对应 Format2GBT9704_2012.js 中 doc.Content.Font.Name = "Times New Roman"
     */
    setAllFontsTimesNewRoman() {
        for (const para of this.paragraphs) {
            const rElements = para.element.getElementsByTagName('w:r');
            for (let i = 0; i < rElements.length; i++) {
                const rElement = rElements[i];
                
                let rPr = rElement.getElementsByTagName('w:rPr')[0];
                if (!rPr) {
                    rPr = this.xmlDoc.createElement('w:rPr');
                    rElement.insertBefore(rPr, rElement.firstChild);
                }
                
                let rFonts = rPr.getElementsByTagName('w:rFonts')[0];
                if (!rFonts) {
                    rFonts = this.xmlDoc.createElement('w:rFonts');
                    rPr.insertBefore(rFonts, rPr.firstChild);
                }
                
                // 设置 ASCII 和 hAnsi 字体为 Times New Roman
                // eastAsia 保留原有中文字体
                rFonts.setAttribute('w:ascii', 'Times New Roman');
                rFonts.setAttribute('w:hAnsi', 'Times New Roman');
            }
        }
    }
    
    /**
     * 取消表格中所有段落的首行缩进
     * 对应 Format2GBT9704_2012.js 表格处理逻辑
     */
    removeTableIndent() {
        const tables = this.xmlDoc.getElementsByTagName('w:tbl');
        for (let t = 0; t < tables.length; t++) {
            const cells = tables[t].getElementsByTagName('w:tc');
            for (let c = 0; c < cells.length; c++) {
                const cell = cells[c];
                const pElements = cell.getElementsByTagName('w:p');
                for (let p = 0; p < pElements.length; p++) {
                    const pElement = pElements[p];
                    
                    let pPr = pElement.getElementsByTagName('w:pPr')[0];
                    if (!pPr) {
                        pPr = this.xmlDoc.createElement('w:pPr');
                        pElement.insertBefore(pPr, pElement.firstChild);
                    }
                    
                    let ind = pPr.getElementsByTagName('w:ind')[0];
                    if (!ind) {
                        ind = this.xmlDoc.createElement('w:ind');
                        pPr.appendChild(ind);
                    }
                    
                    ind.setAttribute('w:firstLine', '0');
                }
            }
        }
    }
    
    /**
     * 在段落中查找并加粗特定前缀文本（仅加粗匹配的前缀本身，而非整个段落）
     * 对应 Format2GBT9704_2012.js 中用 Find/Replace 仅替换前缀字体为粗体
     * @param {number} index - 段落索引
     * @param {Array} prefixes - 要加粗的前缀列表
     */
    setBoldPrefixes(index, prefixes) {
        if (index >= this.paragraphs.length) return;
        
        const pElement = this.paragraphs[index].element;
        const rElements = pElement.getElementsByTagName('w:r');
        if (rElements.length === 0) return;
        
        // 获取段落文本
        const text = this.paragraphs[index].text;
        
        // 查找所有匹配的前缀及其位置
        const matches = [];
        for (const prefix of prefixes) {
            let pos = 0;
            while ((pos = text.indexOf(prefix, pos)) !== -1) {
                matches.push({ start: pos, end: pos + prefix.length });
                pos += prefix.length;
            }
        }
        
        if (matches.length === 0) return;
        
        // 合并重叠的匹配区间
        matches.sort((a, b) => a.start - b.start);
        const mergedMatches = [matches[0]];
        for (let i = 1; i < matches.length; i++) {
            const last = mergedMatches[mergedMatches.length - 1];
            if (matches[i].start <= last.end) {
                last.end = Math.max(last.end, matches[i].end);
            } else {
                mergedMatches.push(matches[i]);
            }
        }
        
        // 由于 setParagraphText 已将文本合并到第一个 run，我们需要拆分 run
        // 获取第一个 run
        const firstR = rElements[0];
        
        // 检查第一个 run 中是否有 w:t
        const tElements = firstR.getElementsByTagName('w:t');
        if (tElements.length === 0) return;
        const tElement = tElements[0];
        const fullText = tElement.textContent || '';
        
        if (fullText !== text) {
            // 文本不匹配，可能已经被修改过，跳过
            return;
        }
        
        // 删除其他 run
        for (let i = rElements.length - 1; i > 0; i--) {
            rElements[i].parentNode.removeChild(rElements[i]);
        }
        
        // 提取原始 run 的字体属性（保留字体名称和字号）
        const originalRPr = firstR.getElementsByTagName('w:rPr')[0];
        const fontProps = {};
        if (originalRPr) {
            const rFonts = originalRPr.getElementsByTagName('w:rFonts')[0];
            if (rFonts) {
                fontProps.ascii = rFonts.getAttribute('w:ascii');
                fontProps.eastAsia = rFonts.getAttribute('w:eastAsia');
                fontProps.hAnsi = rFonts.getAttribute('w:hAnsi');
            }
            const sz = originalRPr.getElementsByTagName('w:sz')[0];
            if (sz) fontProps.sz = sz.getAttribute('w:val');
            const szCs = originalRPr.getElementsByTagName('w:szCs')[0];
            if (szCs) fontProps.szCs = szCs.getAttribute('w:val');
        }
        
        // 按匹配位置拆分文本并创建 run
        let lastPos = 0;
        const fragment = this.xmlDoc.createDocumentFragment();
        
        for (const match of mergedMatches) {
            // 匹配前的文本（不加粗）
            if (match.start > lastPos) {
                const beforeText = fullText.slice(lastPos, match.start);
                const beforeR = this._createRun(beforeText, false, fontProps);
                fragment.appendChild(beforeR);
            }
            
            // 匹配的文本（加粗）
            const matchText = fullText.slice(match.start, match.end);
            const matchR = this._createRun(matchText, true, fontProps);
            fragment.appendChild(matchR);
            
            lastPos = match.end;
        }
        
        // 剩余文本（不加粗）
        if (lastPos < fullText.length) {
            const afterText = fullText.slice(lastPos);
            const afterR = this._createRun(afterText, false, fontProps);
            fragment.appendChild(afterR);
        }
        
        // 替换原段落内容
        // 保留段落属性
        const pPr = pElement.getElementsByTagName('w:pPr')[0];
        // 清空段落
        while (pElement.firstChild) {
            pElement.removeChild(pElement.firstChild);
        }
        // 恢复段落属性
        if (pPr) {
            pElement.appendChild(pPr);
        }
        // 添加新的 run
        const children = Array.from(fragment.childNodes);
        for (const child of children) {
            pElement.appendChild(child);
        }
    }
    
    /**
     * 创建带有指定文本和加粗属性的 run 元素
     * @param {string} text - 文本内容
     * @param {boolean} bold - 是否加粗
     * @param {Object} [fontProps] - 可选的字体属性 {ascii, eastAsia, hAnsi, sz, szCs}
     * @returns {Element} - 创建的 w:r 元素
     */
    _createRun(text, bold, fontProps) {
        const r = this.xmlDoc.createElement('w:r');
        const rPr = this.xmlDoc.createElement('w:rPr');
        
        // 复制字体属性（保留原有字体名称和字号）
        if (fontProps) {
            if (fontProps.ascii || fontProps.eastAsia || fontProps.hAnsi) {
                const rFonts = this.xmlDoc.createElement('w:rFonts');
                if (fontProps.ascii) rFonts.setAttribute('w:ascii', fontProps.ascii);
                if (fontProps.eastAsia) rFonts.setAttribute('w:eastAsia', fontProps.eastAsia);
                if (fontProps.hAnsi) rFonts.setAttribute('w:hAnsi', fontProps.hAnsi);
                rPr.appendChild(rFonts);
            }
            if (fontProps.sz) {
                const sz = this.xmlDoc.createElement('w:sz');
                sz.setAttribute('w:val', fontProps.sz);
                rPr.appendChild(sz);
            }
            if (fontProps.szCs) {
                const szCs = this.xmlDoc.createElement('w:szCs');
                szCs.setAttribute('w:val', fontProps.szCs);
                rPr.appendChild(szCs);
            }
        }
        
        if (bold) {
            const b = this.xmlDoc.createElement('w:b');
            rPr.appendChild(b);
            const bCs = this.xmlDoc.createElement('w:bCs');
            rPr.appendChild(bCs);
        }
        
        r.appendChild(rPr);
        
        const t = this.xmlDoc.createElement('w:t');
        t.setAttribute('xml:space', 'preserve');
        t.textContent = text;
        r.appendChild(t);
        
        return r;
    }
    
    /**
     * 创建带字段的页码 run
     * @param {string} fontName - 字体名称
     * @param {number} fontSizeHalfPt - 字号（半磅）
     * @returns {Array} - run 元素数组
     */
    _createPageFieldRun(fontName, fontSizeHalfPt) {
        const runs = [];
        
        // 创建 rPr 模板
        const rPrXml = `
            <w:rPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                <w:rFonts w:ascii="${fontName}" w:eastAsia="${fontName}" w:hAnsi="${fontName}"/>
                <w:sz w:val="${fontSizeHalfPt}"/>
                <w:szCs w:val="${fontSizeHalfPt}"/>
            </w:rPr>`;
        const parser = new DOMParser();
        const rPrDoc = parser.parseFromString(rPrXml, 'text/xml');
        const rPr = rPrDoc.documentElement;
        
        // — （左空格）
        const r1 = this.xmlDoc.createElement('w:r');
        r1.appendChild(rPr.cloneNode(true));
        const t1 = this.xmlDoc.createElement('w:t');
        t1.setAttribute('xml:space', 'preserve');
        t1.textContent = '— ';
        r1.appendChild(t1);
        runs.push(r1);
        
        // 域开始
        const r2 = this.xmlDoc.createElement('w:r');
        r2.appendChild(rPr.cloneNode(true));
        const fldBegin = this.xmlDoc.createElement('w:fldChar');
        fldBegin.setAttribute('w:fldCharType', 'begin');
        r2.appendChild(fldBegin);
        runs.push(r2);
        
        // 域指令 PAGE
        const r3 = this.xmlDoc.createElement('w:r');
        r3.appendChild(rPr.cloneNode(true));
        const instrText = this.xmlDoc.createElement('w:instrText');
        instrText.textContent = ' PAGE ';
        r3.appendChild(instrText);
        runs.push(r3);
        
        // 域结束
        const r4 = this.xmlDoc.createElement('w:r');
        r4.appendChild(rPr.cloneNode(true));
        const fldEnd = this.xmlDoc.createElement('w:fldChar');
        fldEnd.setAttribute('w:fldCharType', 'end');
        r4.appendChild(fldEnd);
        runs.push(r4);
        
        // — （右空格）
        const r5 = this.xmlDoc.createElement('w:r');
        r5.appendChild(rPr.cloneNode(true));
        const t5 = this.xmlDoc.createElement('w:t');
        t5.setAttribute('xml:space', 'preserve');
        t5.textContent = ' —';
        r5.appendChild(t5);
        runs.push(r5);
        
        return runs;
    }
    
    /**
     * 设置页码
     * 对应 Format2GBT9704_2012.js 中的页码设置逻辑：
     * 奇偶页不同，"— 1 —"格式，奇数页右侧，偶数页左侧，4号宋体
     */
    async setPageNumber() {
        const fontName = '宋体';
        const fontSizeHalfPt = 28; // 4号 = 14pt，半磅 = 28
        
        // 1. 创建页脚 XML
        // 奇数页页脚（右对齐）
        const oddFooterXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:p>
    <w:pPr>
      <w:jc w:val="right"/>
      <w:rPr>
        <w:rFonts w:ascii="${fontName}" w:eastAsia="${fontName}" w:hAnsi="${fontName}"/>
        <w:sz w:val="${fontSizeHalfPt}"/>
        <w:szCs w:val="${fontSizeHalfPt}"/>
      </w:rPr>
    </w:pPr>
  </w:p>
</w:ftr>`;
        
        // 偶数页页脚（左对齐）
        const evenFooterXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:p>
    <w:pPr>
      <w:jc w:val="left"/>
      <w:rPr>
        <w:rFonts w:ascii="${fontName}" w:eastAsia="${fontName}" w:hAnsi="${fontName}"/>
        <w:sz w:val="${fontSizeHalfPt}"/>
        <w:szCs w:val="${fontSizeHalfPt}"/>
      </w:rPr>
    </w:pPr>
  </w:p>
</w:ftr>`;
        
        // 解析 footer XML 并添加 PAGE 字段
        const parser = new DOMParser();
        
        // 奇数页页脚
        const oddFooterDoc = parser.parseFromString(oddFooterXml, 'text/xml');
        const oddP = oddFooterDoc.documentElement.getElementsByTagName('w:p')[0];
        const oddRuns = this._createPageFieldRun(fontName, fontSizeHalfPt);
        for (const run of oddRuns) {
            // Import node from the document's ownerDocument context
            const importedRun = oddFooterDoc.importNode(run, true);
            oddP.appendChild(importedRun);
        }
        
        // 偶数页页脚
        const evenFooterDoc = parser.parseFromString(evenFooterXml, 'text/xml');
        const evenP = evenFooterDoc.documentElement.getElementsByTagName('w:p')[0];
        const evenRuns = this._createPageFieldRun(fontName, fontSizeHalfPt);
        for (const run of evenRuns) {
            const importedRun = evenFooterDoc.importNode(run, true);
            evenP.appendChild(importedRun);
        }
        
        // 序列化为字符串
        const serializer = new XMLSerializer();
        const oddFooterStr = serializer.serializeToString(oddFooterDoc.documentElement);
        const evenFooterStr = serializer.serializeToString(evenFooterDoc.documentElement);
        
        // 2. 添加页脚文件到 zip
        this.zip.file('word/footer1.xml', oddFooterStr);
        this.zip.file('word/footer2.xml', evenFooterStr);
        
        // 3. 更新 [Content_Types].xml
        const contentTypesStr = await this.zip.file('[Content_Types].xml').async('string');
        const contentTypesDoc = parser.parseFromString(contentTypesStr, 'text/xml');
        const typesRoot = contentTypesDoc.documentElement;
        
        // 检查是否已存在 footer 的 Override
        const existingOverrides = typesRoot.getElementsByTagName('Override');
        let hasOddFooter = false;
        let hasEvenFooter = false;
        for (let i = 0; i < existingOverrides.length; i++) {
            const partName = existingOverrides[i].getAttribute('PartName');
            if (partName === '/word/footer1.xml') hasOddFooter = true;
            if (partName === '/word/footer2.xml') hasEvenFooter = true;
        }
        
        if (!hasOddFooter) {
            const override = contentTypesDoc.createElement('Override');
            override.setAttribute('PartName', '/word/footer1.xml');
            override.setAttribute('ContentType', 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml');
            typesRoot.appendChild(override);
        }
        if (!hasEvenFooter) {
            const override = contentTypesDoc.createElement('Override');
            override.setAttribute('PartName', '/word/footer2.xml');
            override.setAttribute('ContentType', 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml');
            typesRoot.appendChild(override);
        }
        this.zip.file('[Content_Types].xml', serializer.serializeToString(contentTypesDoc.documentElement));
        
        // 4. 更新 document.xml.rels
        const relsStr = await this.zip.file('word/_rels/document.xml.rels').async('string');
        const relsDoc = parser.parseFromString(relsStr, 'text/xml');
        const relsRoot = relsDoc.documentElement;
        
        const ns = 'http://schemas.openxmlformats.org/package/2006/relationships';
        const footerRelType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer';
        
        // 查找最大的 rId
        const existingRels = relsRoot.getElementsByTagNameNS(ns, 'Relationship');
        let maxId = 0;
        for (let i = 0; i < existingRels.length; i++) {
            const id = existingRels[i].getAttribute('Id');
            if (id) {
                const num = parseInt(id.replace('rId', ''), 10);
                if (num > maxId) maxId = num;
            }
        }
        
        let hasOddFooterRel = false;
        let hasEvenFooterRel = false;
        for (let i = 0; i < existingRels.length; i++) {
            const type = existingRels[i].getAttribute('Type');
            const target = existingRels[i].getAttribute('Target');
            if (type === footerRelType && target === 'footer1.xml') hasOddFooterRel = true;
            if (type === footerRelType && target === 'footer2.xml') hasEvenFooterRel = true;
        }
        
        if (!hasOddFooterRel) {
            const rel = relsDoc.createElementNS(ns, 'Relationship');
            rel.setAttribute('Id', `rId${maxId + 1}`);
            rel.setAttribute('Type', footerRelType);
            rel.setAttribute('Target', 'footer1.xml');
            relsRoot.appendChild(rel);
            maxId++;
        }
        if (!hasEvenFooterRel) {
            const rel = relsDoc.createElementNS(ns, 'Relationship');
            rel.setAttribute('Id', `rId${maxId + 1}`);
            rel.setAttribute('Type', footerRelType);
            rel.setAttribute('Target', 'footer2.xml');
            relsRoot.appendChild(rel);
            maxId++;
        }
        this.zip.file('word/_rels/document.xml.rels', serializer.serializeToString(relsDoc.documentElement));
        
        // 5. 更新 document.xml 中的 sectPr 添加 footerReference
        const body = this.xmlDoc.getElementsByTagName('w:body')[0];
        if (!body) return;
        
        let sectPr = body.getElementsByTagName('w:sectPr')[0];
        if (!sectPr) {
            sectPr = this.xmlDoc.createElement('w:sectPr');
            body.appendChild(sectPr);
        }
        
        // 获取 footer 的 rId
        const updatedRelsDoc = parser.parseFromString(
            await this.zip.file('word/_rels/document.xml.rels').async('string'),
            'text/xml'
        );
        const updatedRels = updatedRelsDoc.documentElement.getElementsByTagNameNS(ns, 'Relationship');
        let oddRId = null;
        let evenRId = null;
        for (let i = 0; i < updatedRels.length; i++) {
            const type = updatedRels[i].getAttribute('Type');
            const target = updatedRels[i].getAttribute('Target');
            if (type === footerRelType && target === 'footer1.xml') {
                oddRId = updatedRels[i].getAttribute('Id');
            }
            if (type === footerRelType && target === 'footer2.xml') {
                evenRId = updatedRels[i].getAttribute('Id');
            }
        }
        
        // 删除已有的 footerReference
        const existingFooterRefs = sectPr.getElementsByTagName('w:footerReference');
        while (existingFooterRefs.length > 0) {
            existingFooterRefs[0].parentNode.removeChild(existingFooterRefs[0]);
        }
        
        // 添加奇数页页脚引用
        if (oddRId) {
            const footerRefOdd = this.xmlDoc.createElement('w:footerReference');
            footerRefOdd.setAttribute('w:type', 'default');
            footerRefOdd.setAttribute('r:id', oddRId);
            sectPr.insertBefore(footerRefOdd, sectPr.firstChild);
        }
        
        // 添加偶数页页脚引用
        if (evenRId) {
            const footerRefEven = this.xmlDoc.createElement('w:footerReference');
            footerRefEven.setAttribute('w:type', 'even');
            footerRefEven.setAttribute('r:id', evenRId);
            sectPr.insertBefore(footerRefEven, sectPr.firstChild);
        }
        
        // 6. 启用奇偶页不同（对应 doc.PageSetup.OddAndEvenPagesHeaderFooter = true）
        const settingsStr = await this.zip.file('word/settings.xml').async('string');
        const settingsDoc = parser.parseFromString(settingsStr, 'text/xml');
        const settingsRoot = settingsDoc.documentElement;
        const existingEvenOdd = settingsRoot.getElementsByTagName('w:evenAndOddHeaders');
        if (existingEvenOdd.length === 0) {
            const evenOdd = settingsDoc.createElement('w:evenAndOddHeaders');
            settingsRoot.appendChild(evenOdd);
        }
        this.zip.file('word/settings.xml', serializer.serializeToString(settingsDoc.documentElement));
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
     * 修改样式定义（Normal/Heading1-4）
     * 对应 Format2GBT9704_2012.js 中设置 Normal 样式和各标题样式的逻辑
     */
    async updateStyleDefinitions() {
        const stylesStr = await this.zip.file('word/styles.xml').async('string');
        const parser = new DOMParser();
        const stylesDoc = parser.parseFromString(stylesStr, 'text/xml');
        const stylesRoot = stylesDoc.documentElement;
        
        // 样式配置
        const styleConfigs = [
            { id: 'Normal', font: '仿宋_GB2312', size: 32, bold: false, indent: charactersToPoints(2) },
            { id: 'Heading1', font: '黑体', size: 32, bold: false, indent: charactersToPoints(2) },
            { id: 'Heading2', font: '楷体_GB2312', size: 32, bold: false, indent: charactersToPoints(2) },
            { id: 'Heading3', font: '仿宋_GB2312', size: 32, bold: true, indent: charactersToPoints(2) },
            { id: 'Heading4', font: '仿宋_GB2312', size: 32, bold: true, indent: charactersToPoints(2) }
        ];
        
        for (const config of styleConfigs) {
            // 查找样式元素
            const styleElements = stylesRoot.getElementsByTagName('w:style');
            let styleEl = null;
            for (let i = 0; i < styleElements.length; i++) {
                if (styleElements[i].getAttribute('w:styleId') === config.id) {
                    styleEl = styleElements[i];
                    break;
                }
            }
            if (!styleEl) continue;
            
            // 更新字体
            let rPr = styleEl.getElementsByTagName('w:rPr')[0];
            if (!rPr) {
                rPr = stylesDoc.createElement('w:rPr');
                // 插入到 style 的第一个子元素位置
                const firstChild = styleEl.firstChild;
                if (firstChild) {
                    styleEl.insertBefore(rPr, firstChild);
                } else {
                    styleEl.appendChild(rPr);
                }
            }
            
            // 字体名称
            let rFonts = rPr.getElementsByTagName('w:rFonts')[0];
            if (!rFonts) {
                rFonts = stylesDoc.createElement('w:rFonts');
                rPr.insertBefore(rFonts, rPr.firstChild);
            }
            rFonts.setAttribute('w:ascii', config.font);
            rFonts.setAttribute('w:eastAsia', config.font);
            rFonts.setAttribute('w:hAnsi', config.font);
            
            // 字号
            let sz = rPr.getElementsByTagName('w:sz')[0];
            if (!sz) {
                sz = stylesDoc.createElement('w:sz');
                rPr.appendChild(sz);
            }
            sz.setAttribute('w:val', config.size.toString());
            
            let szCs = rPr.getElementsByTagName('w:szCs')[0];
            if (!szCs) {
                szCs = stylesDoc.createElement('w:szCs');
                rPr.appendChild(szCs);
            }
            szCs.setAttribute('w:val', config.size.toString());
            
            // 加粗
            if (config.bold) {
                if (!rPr.getElementsByTagName('w:b')[0]) {
                    const b = stylesDoc.createElement('w:b');
                    rPr.appendChild(b);
                }
                if (!rPr.getElementsByTagName('w:bCs')[0]) {
                    const bCs = stylesDoc.createElement('w:bCs');
                    rPr.appendChild(bCs);
                }
            } else {
                const b = rPr.getElementsByTagName('w:b')[0];
                if (b) b.parentNode.removeChild(b);
                const bCs = rPr.getElementsByTagName('w:bCs')[0];
                if (bCs) bCs.parentNode.removeChild(bCs);
            }
            
            // 更新段落属性
            let pPr = styleEl.getElementsByTagName('w:pPr')[0];
            if (!pPr) {
                pPr = stylesDoc.createElement('w:pPr');
                styleEl.insertBefore(pPr, styleEl.firstChild);
            }
            
            // 行距
            let spacing = pPr.getElementsByTagName('w:spacing')[0];
            if (!spacing) {
                spacing = stylesDoc.createElement('w:spacing');
                pPr.appendChild(spacing);
            }
            spacing.setAttribute('w:line', ptToTwips(CONFIG.lineSpacing).toString());
            spacing.setAttribute('w:lineRule', 'exact');
            spacing.setAttribute('w:before', '0');
            spacing.setAttribute('w:after', '0');
            spacing.setAttribute('w:beforeAutospacing', '0');
            spacing.setAttribute('w:afterAutospacing', '0');
            
            // 首行缩进
            let ind = pPr.getElementsByTagName('w:ind')[0];
            if (!ind) {
                ind = stylesDoc.createElement('w:ind');
                pPr.appendChild(ind);
            }
            ind.setAttribute('w:firstLine', ptToTwips(config.indent).toString());
        }
        
        // 写回 zip
        const serializer = new XMLSerializer();
        this.zip.file('word/styles.xml', serializer.serializeToString(stylesDoc.documentElement));
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
