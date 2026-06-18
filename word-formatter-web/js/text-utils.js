/**
 * 文本处理工具
 * 从Format2GBT9704_2012.js提取的纯文本处理逻辑
 */

/**
 * 替换标点符号
 * @param {string} text - 原始文本
 * @returns {string} - 替换后的文本
 */
function replacePunctuation(text) {
    // 括号
    text = text.replace(/\(/g, '（');
    text = text.replace(/\)/g, '）');
    // 逗号
    text = text.replace(/,/g, '，');
    // 冒号
    text = text.replace(/:/g, '：');
    // 问号
    text = text.replace(/\?/g, '？');
    // 分号
    text = text.replace(/;/g, '；');
    // 句号（智能判断）
    text = replacePeriod(text);
    // 引号（区分左右）
    text = replaceQuotes(text);
    return text;
}

/**
 * 智能替换句号
 * 只替换前后都没有数字的句号
 * @param {string} text - 原始文本
 * @returns {string} - 替换后的文本
 */
function replacePeriod(text) {
    // 前后都没有数字的句号
    text = text.replace(/([!0-9])\.([!0-9])/g, '$1。$2');
    // 段落末尾的句号
    if (text.endsWith('.')) {
        text = text.slice(0, -1) + '。';
    }
    return text;
}

/**
 * 替换引号（区分左右）
 * @param {string} text - 原始文本
 * @returns {string} - 替换后的文本
 */
function replaceQuotes(text) {
    // 双引号
    let isFirstDouble = true;
    text = text.replace(/"/g, () => {
        const quote = isFirstDouble ? '\u201C' : '\u201D';
        isFirstDouble = !isFirstDouble;
        return quote;
    });
    // 单引号
    let isFirstSingle = true;
    text = text.replace(/'/g, () => {
        const quote = isFirstSingle ? '\u2018' : '\u2019';
        isFirstSingle = !isFirstSingle;
        return quote;
    });
    return text;
}

/**
 * 处理空格
 * 保留英文之间的空格，删除其他空格
 * @param {string} text - 原始文本
 * @returns {string} - 处理后的文本
 */
function processSpaces(text) {
    // 保留英文之间的空格
    text = text.replace(/([a-zA-Z]) ([a-zA-Z])/g, '$1§TEMP§$2');
    // 删除其他空格
    text = text.replace(/ /g, '');
    // 恢复英文之间的空格
    text = text.replace(/§TEMP§/g, ' ');
    return text;
}

/**
 * 2字符缩进（32磅）
 * @param {number} charCount - 字符数
 * @returns {number} - 磅值
 */
function charactersToPoints(charCount) {
    return charCount * 16; // 1字符 = 16磅
}

/**
 * 厘米转twips
 * @param {number} cm - 厘米值
 * @returns {number} - twips值
 */
function cmToTwips(cm) {
    return Math.round(cm * 567);
}

/**
 * 磅转twips
 * @param {number} pt - 磅值
 * @returns {number} - twips值
 */
function ptToTwips(pt) {
    return Math.round(pt * 20);
}

/**
 * 厘米转磅
 * @param {number} cm - 厘米值
 * @returns {number} - 磅值
 */
function cmToPt(cm) {
    return Math.round(cm * 28.35);
}
