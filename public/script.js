// 初始化文件选择事件
document.getElementById('wordFile').addEventListener('change', function(e) {
  document.getElementById('wordFileName').textContent = 
    e.target.files[0] ? e.target.files[0].name : '未选择文件';
});

document.getElementById('excelFile').addEventListener('change', function(e) {
  document.getElementById('excelFileName').textContent = 
    e.target.files[0] ? e.target.files[0].name : '未选择文件';
});

document.getElementById('powerpointFile').addEventListener('change', function(e) {
  document.getElementById('powerpointFileName').textContent = 
    e.target.files[0] ? e.target.files[0].name : '未选择文件';
});

async function convertDocument(type, fileInput, statusElement) {
  const file = fileInput.files[0];
  
  if (!file) {
    showStatus(statusElement, '请先选择文件', 'error');
    return;
  }
  
  try {
    showStatus(statusElement, '转换中...请稍候', 'processing');
    
    const formData = new FormData();
    formData.append('file', file);
    
    // 使用相对路径，自动适应部署环境
    const response = await fetch(`/convert-${type}`, {
      method: 'POST',
      body: formData
    });
    
    if (!response.ok) {
      const error = await response.json();
      throw new Error(error.error || '转换失败');
    }
    
    const blob = await response.blob();
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    
    // 保持原始文件名不变
    const originalFilename = file.name;
    const filenameWithoutExtension = originalFilename.lastIndexOf('.') > 0 
      ? originalFilename.substring(0, originalFilename.lastIndexOf('.'))
      : originalFilename;
    const newFilename = `${filenameWithoutExtension}.pdf`;
    
    a.download = newFilename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    
    // 释放URL对象
    setTimeout(() => window.URL.revokeObjectURL(url), 100);
    
    showStatus(statusElement, `转换成功！已下载 ${newFilename}`, 'success');
  } catch (error) {
    console.error(`${type}转换错误:`, error);
    showStatus(statusElement, `错误: ${error.message}`, 'error');
  }
}

function showStatus(element, message, type) {
  element.textContent = message;
  element.className = 'status';
  element.classList.add(type);
}

function convertWord() {
  convertDocument('word', 
    document.getElementById('wordFile'), 
    document.getElementById('wordStatus'));
}

function convertExcel() {
  convertDocument('excel', 
    document.getElementById('excelFile'), 
    document.getElementById('excelStatus'));
}

function convertPowerpoint() {
  convertDocument('powerpoint', 
    document.getElementById('powerpointFile'), 
    document.getElementById('powerpointStatus'));
}

// 添加拖放文件支持
document.querySelectorAll('.file-area').forEach(area => {
  area.addEventListener('dragover', (e) => {
    e.preventDefault();
    area.classList.add('dragover');
  });
  
  area.addEventListener('dragleave', () => {
    area.classList.remove('dragover');
  });
  
  area.addEventListener('drop', (e) => {
    e.preventDefault();
    area.classList.remove('dragover');
    
    if (e.dataTransfer.files.length) {
      const input = area.querySelector('input[type="file"]');
      input.files = e.dataTransfer.files;
      
      // 触发change事件
      const event = new Event('change', { bubbles: true });
      input.dispatchEvent(event);
    }
  });
});