/**
 * Web应用入口
 * Word格式化工具（GB/T 9704-2012）
 */

const formatter = new Formatter();

// 文件列表
let files = [];
let results = [];

/**
 * 初始化应用
 */
function initApp() {
    const dropZone = document.getElementById('dropZone');
    const fileInput = document.getElementById('fileInput');
    
    // 拖拽上传
    dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        e.stopPropagation();
        dropZone.classList.add('dragover');
    });
    
    dropZone.addEventListener('dragleave', (e) => {
        e.preventDefault();
        e.stopPropagation();
        dropZone.classList.remove('dragover');
    });
    
    dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        e.stopPropagation();
        dropZone.classList.remove('dragover');
        handleFiles(e.dataTransfer.files);
    });
    
    // 点击选择
    dropZone.addEventListener('click', () => {
        fileInput.click();
    });
    
    fileInput.addEventListener('change', (e) => {
        handleFiles(e.target.files);
        // 清空input，允许重复选择同一文件
        fileInput.value = '';
    });
}

/**
 * 处理文件
 * @param {FileList} fileList - 文件列表
 */
function handleFiles(fileList) {
    for (const file of fileList) {
        // 验证文件格式
        if (!file.name.endsWith('.docx')) {
            addFileToList(file.name, 'error', '不支持的文件格式（仅支持.docx）');
            continue;
        }
        
        // 验证文件大小（10MB）
        if (file.size > 10 * 1024 * 1024) {
            addFileToList(file.name, 'error', '文件过大（最大10MB）');
            continue;
        }
        
        // 检查是否已存在
        if (files.some(f => f.name === file.name)) {
            addFileToList(file.name, 'warning', '文件已存在');
            continue;
        }
        
        files.push(file);
        addFileToList(file.name, 'pending', '等待处理');
    }
    
    // 更新按钮状态
    updateButtonStates();
}

/**
 * 添加文件到列表
 * @param {string} filename - 文件名
 * @param {string} status - 状态
 * @param {string} message - 消息
 */
function addFileToList(filename, status, message) {
    const fileList = document.getElementById('fileList');
    
    // 检查是否已存在
    const existingItem = document.getElementById(`file-${filename}`);
    if (existingItem) {
        // 更新状态
        const statusEl = existingItem.querySelector('.file-status');
        statusEl.className = `file-status ${status}`;
        statusEl.textContent = message;
        return;
    }
    
    // 创建新项
    const item = document.createElement('div');
    item.className = 'file-item';
    item.id = `file-${filename}`;
    item.innerHTML = `
        <span class="file-name">${filename}</span>
        <span class="file-status ${status}">${message}</span>
    `;
    fileList.appendChild(item);
}

/**
 * 更新文件状态
 * @param {string} filename - 文件名
 * @param {string} status - 状态
 * @param {string} message - 消息
 */
function updateFileStatus(filename, status, message) {
    const item = document.getElementById(`file-${filename}`);
    if (item) {
        const statusEl = item.querySelector('.file-status');
        statusEl.className = `file-status ${status}`;
        statusEl.textContent = message;
    }
}

/**
 * 更新按钮状态
 */
function updateButtonStates() {
    const processBtn = document.getElementById('processBtn');
    const downloadBtn = document.getElementById('downloadBtn');
    const clearBtn = document.getElementById('clearBtn');
    
    processBtn.disabled = files.length === 0;
    downloadBtn.disabled = results.filter(r => r.status === 'success').length === 0;
    clearBtn.disabled = files.length === 0;
}

/**
 * 开始处理
 */
async function startProcessing() {
    if (files.length === 0) {
        alert('请先上传文件');
        return;
    }
    
    const processBtn = document.getElementById('processBtn');
    processBtn.disabled = true;
    processBtn.textContent = '处理中...';
    
    results = [];
    
    for (const file of files) {
        updateFileStatus(file.name, 'processing', '处理中...');
        
        try {
            const blob = await formatter.format(file);
            const filename = file.name.replace('.docx', '_formatted.docx');
            results.push({ filename, blob, status: 'success' });
            updateFileStatus(file.name, 'success', '处理完成');
        } catch (error) {
            console.error(`处理失败: ${file.name}`, error);
            updateFileStatus(file.name, 'error', '处理失败');
            results.push({ filename: file.name, error: error.message, status: 'error' });
        }
    }
    
    processBtn.disabled = false;
    processBtn.textContent = '开始处理';
    
    // 更新按钮状态
    updateButtonStates();
    
    // 显示完成消息
    const successCount = results.filter(r => r.status === 'success').length;
    const errorCount = results.filter(r => r.status === 'error').length;
    
    if (errorCount === 0) {
        alert(`处理完成！共 ${successCount} 个文件。`);
    } else {
        alert(`处理完成！成功 ${successCount} 个，失败 ${errorCount} 个。`);
    }
}

/**
 * 下载文件
 */
async function downloadFiles() {
    const successResults = results.filter(r => r.status === 'success');
    
    if (successResults.length === 0) {
        alert('没有可下载的文件');
        return;
    }
    
    if (successResults.length === 1) {
        // 单文件直接下载
        saveAs(successResults[0].blob, successResults[0].filename);
    } else {
        // 多文件打包为zip
        const zip = new JSZip();
        for (const result of successResults) {
            zip.file(result.filename, result.blob);
        }
        
        const zipBlob = await zip.generateAsync({ type: 'blob' });
        saveAs(zipBlob, 'formatted_documents.zip');
    }
}

/**
 * 清空列表
 */
function clearList() {
    files = [];
    results = [];
    document.getElementById('fileList').innerHTML = '';
    updateButtonStates();
}

/**
 * 页面加载完成后初始化
 */
window.addEventListener('DOMContentLoaded', initApp);
