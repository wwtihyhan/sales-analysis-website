// 全局变量
let currentReportId = null;
const API_BASE = window.location.origin;

// DOM加载完成后初始化
document.addEventListener('DOMContentLoaded', () => {
    initUpload();
    loadHistory();
});

// 初始化上传功能
function initUpload() {
    const uploadArea = document.getElementById('uploadArea');
    const fileInput = document.getElementById('fileInput');
    const uploadBtn = document.getElementById('uploadBtn');
    
    // 点击上传区域或按钮触发文件选择
    uploadArea.addEventListener('click', (e) => {
        if (e.target.id !== 'uploadBtn') {
            fileInput.click();
        }
    });
    
    uploadBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        fileInput.click();
    });
    
    // 文件选择完成
    fileInput.addEventListener('change', (e) => {
        if (e.target.files.length > 0) {
            uploadFile(e.target.files[0]);
        }
    });
    
    // 拖拽上传
    uploadArea.addEventListener('dragover', (e) => {
        e.preventDefault();
        uploadArea.style.borderColor = '#764ba2';
        uploadArea.style.background = '#f0f2ff';
    });
    
    uploadArea.addEventListener('dragleave', (e) => {
        e.preventDefault();
        uploadArea.style.borderColor = '#667eea';
        uploadArea.style.background = '#f8f9ff';
    });
    
    uploadArea.addEventListener('drop', (e) => {
        e.preventDefault();
        uploadArea.style.borderColor = '#667eea';
        uploadArea.style.background = '#f8f9ff';
        
        if (e.dataTransfer.files.length > 0) {
            uploadFile(e.dataTransfer.files[0]);
        }
    });
}

// 上传文件
async function uploadFile(file) {
    // 验证文件类型
    if (!file.name.match(/\.(xls|xlsx)$/)) {
        alert('只支持 .xls 或 .xlsx 格式的文件！');
        return;
    }
    
    // 显示进度条
    document.getElementById('uploadProgress').style.display = 'block';
    document.getElementById('uploadArea').style.display = 'none';
    
    const formData = new FormData();
    formData.append('file', file);
    
    try {
        // 模拟进度
        let progress = 0;
        const progressInterval = setInterval(() => {
            progress += 10;
            if (progress > 90) clearInterval(progressInterval);
            document.getElementById('progressFill').style.width = progress + '%';
            document.getElementById('progressText').textContent = '正在分析数据...';
        }, 200);
        
        // 发送请求
        const response = await fetch(`${API_BASE}/api/upload`, {
            method: 'POST',
            body: formData
        });
        
        clearInterval(progressInterval);
        
        if (!response.ok) {
            const error = await response.json();
            throw new Error(error.detail || '上传失败');
        }
        
        const result = await response.json();
        
        // 更新进度条
        document.getElementById('progressFill').style.width = '100%';
        document.getElementById('progressText').textContent = '分析完成！';
        
        // 显示报告
        currentReportId = result.id;
        showReport(result.id);
        
        // 重新加载历史记录
        loadHistory();
        
    } catch (error) {
        alert('上传失败：' + error.message);
        resetUpload();
    }
}

// 显示报告
function showReport(reportId) {
    document.getElementById('reportSection').style.display = 'block';
    document.getElementById('reportFrame').src = `${API_BASE}/api/report/${reportId}`;
    
    // 设置下载按钮
    document.getElementById('downloadBtn').onclick = () => {
        window.open(`${API_BASE}/api/report/${reportId}/download`, '_blank');
    };
    
    // 设置全屏查看按钮
    document.getElementById('viewFullBtn').onclick = () => {
        window.open(`${API_BASE}/api/report/${reportId}`, '_blank');
    };
}

// 重置上传区域
function resetUpload() {
    document.getElementById('uploadProgress').style.display = 'none';
    document.getElementById('uploadArea').style.display = 'block';
    document.getElementById('progressFill').style.width = '0%';
    document.getElementById('fileInput').value = '';
}

// 加载历史记录
async function loadHistory(page = 1) {
    try {
        const response = await fetch(`${API_BASE}/api/records?page=${page}&limit=10`);
        const data = await response.json();
        
        renderHistory(data.records);
        renderPagination(data.total, page, 10);
        
    } catch (error) {
        console.error('加载历史记录失败：', error);
    }
}

// 渲染历史记录列表
function renderHistory(records) {
    const historyList = document.getElementById('historyList');
    
    if (records.length === 0) {
        historyList.innerHTML = '<p class="empty-text">暂无历史记录</p>';
        return;
    }
    
    historyList.innerHTML = records.map(record => `
        <div class="history-item" onclick="showReport('${record.id}')">
            <div class="history-info">
                <h3>📄 ${record.file_name}</h3>
                <p>上传时间：${record.upload_time}</p>
            </div>
            <div class="history-stats">
                <div class="revenue">营业：¥${record.total_revenue.toFixed(2)}</div>
                <div class="profit">利润：¥${record.total_profit.toFixed(2)}</div>
            </div>
        </div>
    `).join('');
}

// 渲染分页
function renderPagination(total, currentPage, limit) {
    const pagination = document.getElementById('pagination');
    const totalPages = Math.ceil(total / limit);
    
    if (totalPages <= 1) {
        pagination.innerHTML = '';
        return;
    }
    
    let html = '';
    
    // 上一页
    if (currentPage > 1) {
        html += `<button class="page-btn" onclick="loadHistory(${currentPage - 1})">上一页</button>`;
    }
    
    // 页码
    for (let i = 1; i <= totalPages; i++) {
        if (i === currentPage) {
            html += `<button class="page-btn active">${i}</button>`;
        } else {
            html += `<button class="page-btn" onclick="loadHistory(${i})">${i}</button>`;
        }
    }
    
    // 下一页
    if (currentPage < totalPages) {
        html += `<button class="page-btn" onclick="loadHistory(${currentPage + 1})">下一页</button>`;
    }
    
    pagination.innerHTML = html;
}
