/**
 * 标题识别逻辑
 * 从Format2GBT9704_2012.js提取的标题识别逻辑
 */

/**
 * 识别标题级别
 * @param {string} text - 段落文本
 * @param {number} index - 段落索引
 * @param {Array} allParagraphs - 所有段落
 * @returns {Object|null} - { level, prefix } 或 null（正文）
 */
function identifyHeadingLevel(text, index, allParagraphs) {
    const trimmedText = text.trim();
    if (!trimmedText) return null;
    
    // 检查一级标题
    for (const prefix of CONFIG.headings.level1) {
        if (trimmedText.startsWith(prefix)) {
            return { level: 1, prefix: prefix };
        }
    }
    
    // 检查二级标题
    for (const prefix of CONFIG.headings.level2) {
        if (trimmedText.startsWith(prefix)) {
            return { level: 2, prefix: prefix };
        }
    }
    
    // 检查三级标题（需要连续判断）
    for (const prefix of CONFIG.headings.level3) {
        if (trimmedText.startsWith(prefix)) {
            if (!isConsecutiveHeading(index, allParagraphs, CONFIG.headings.level3)) {
                return { level: 3, prefix: prefix };
            }
            break;
        }
    }
    
    // 检查四级标题（需要连续判断）
    for (const prefix of CONFIG.headings.level4) {
        if (trimmedText.startsWith(prefix)) {
            if (!isConsecutiveHeading(index, allParagraphs, CONFIG.headings.level4)) {
                return { level: 4, prefix: prefix };
            }
            break;
        }
    }
    
    return null; // 正文
}

/**
 * 判断是否为连续标题（排除正文）
 * @param {number} index - 当前段落索引
 * @param {Array} paragraphs - 所有段落
 * @param {Array} prefixes - 标题前缀数组
 * @returns {boolean} - 是否为连续标题
 */
function isConsecutiveHeading(index, paragraphs, prefixes) {
    // 检查下一段
    if (index < paragraphs.length - 1) {
        const nextText = paragraphs[index + 1].text.trim();
        for (const prefix of prefixes) {
            if (nextText.startsWith(prefix)) {
                return true;
            }
        }
    }
    
    // 检查上一段
    if (index > 0) {
        const prevText = paragraphs[index - 1].text.trim();
        for (const prefix of prefixes) {
            if (prevText.startsWith(prefix)) {
                return true;
            }
        }
    }
    
    return false;
}

/**
 * 检查段落是否包含图片
 * @param {string} paragraphXml - 段落XML字符串
 * @returns {boolean} - 是否包含图片
 */
function hasImage(paragraphXml) {
    if (!paragraphXml) return false;
    return paragraphXml.includes('<w:drawing') || 
           paragraphXml.includes('<w:pict') ||
           paragraphXml.includes('<w:imagedata');
}

/**
 * 检查段落是否包含嵌入型图片
 * @param {Object} pElement - 段落XML元素
 * @returns {boolean} - 是否包含嵌入型图片
 */
function hasInlineImage(pElement) {
    if (!pElement) return false;
    const xml = pElement.toString();
    return hasImage(xml);
}

/**
 * 检查文本是否以特定前缀开头
 * @param {string} text - 文本
 * @param {Array} prefixes - 前缀数组
 * @returns {boolean} - 是否匹配
 */
function startsWithAny(text, prefixes) {
    for (const prefix of prefixes) {
        if (text.startsWith(prefix)) {
            return true;
        }
    }
    return false;
}

/**
 * 检查文本是否包含特定关键词
 * @param {string} text - 文本
 * @param {Array} keywords - 关键词数组
 * @returns {boolean} - 是否包含
 */
function containsAny(text, keywords) {
    for (const keyword of keywords) {
        if (text.includes(keyword)) {
            return true;
        }
    }
    return false;
}
