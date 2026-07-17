/**
 * Web应用入口
 * Word格式化工具（GB/T 9704-2012）
 * 导入文档后自动执行并自动下载
 */

const formatter = new Formatter();

// 存储已处理文件的结果（用于重新下载）
const processedResults = {};

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
    
    // 键盘访问
    dropZone.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' || e.key === ' ') {
            e.preventDefault();
            fileInput.click();
        }
    });
    
    fileInput.addEventListener('change', (e) => {
        handleFiles(e.target.files);
        // 清空input，允许重复选择同一文件
        fileInput.value = '';
    });
}

/**
 * 处理文件 — 逐个处理并逐个下载
 * @param {FileList} fileList - 文件列表
 */
function handleFiles(fileList) {
    for (const file of fileList) {
        // 验证文件格式
        if (!file.name.endsWith('.docx')) {
            addFileToList(file.name, 'error', '仅支持.docx格式');
            continue;
        }
        
        addFileToList(file.name, 'pending', '等待处理');
        
        // 立即开始处理
        processSingleFile(file);
    }
}

/**
 * 处理单个文件并自动下载
 * @param {File} file - 文件对象
 */
async function processSingleFile(file) {
    updateFileStatus(file.name, 'processing', '处理中...');
    
    try {
        const blob = await formatter.format(file);
        const filename = file.name.replace('.docx', '已格式化.docx');
        
        // 保存结果以便重新下载
        processedResults[file.name] = { blob, filename };
        
        updateFileStatus(file.name, 'success', '处理完成');
        
        // 自动下载
        saveAs(blob, filename);
    } catch (error) {
        console.error(`处理失败: ${file.name}`, error);
        updateFileStatus(file.name, 'error', '处理失败');
    }
}

/**
 * 添加文件到列表，点击可重新下载（处理成功时）
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
    
    // 点击已处理成功的文件可重新下载
    item.addEventListener('click', () => {
        const result = processedResults[filename];
        if (result) {
            saveAs(result.blob, result.filename);
        }
    });
    
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
 * 页面加载完成后初始化
 */
window.addEventListener('DOMContentLoaded', initApp);
