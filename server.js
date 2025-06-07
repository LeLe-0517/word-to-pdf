const express = require('express');
const multer = require('multer');
const fs = require('fs');
const path = require('path');
const cors = require('cors');
const wordToPdf = require('./word-to-pdf');
const excelToPdf = require('./excel-to-pdf');
const powerpointToPdf = require('./powerpoint-to-pdf');

const app = express();
const PORT = process.env.PORT || 3000;

// 启用CORS以支持跨域访问
app.use(cors());

// 设置静态文件目录（前端文件）
app.use(express.static(path.join(__dirname, 'public')));

// 确保上传目录存在
const uploadDir = path.join(__dirname, 'uploads');
if (!fs.existsSync(uploadDir)) {
  fs.mkdirSync(uploadDir, { recursive: true });
}

// 配置文件上传
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, uploadDir);
  },
  filename: (req, file, cb) => {
    const uniqueSuffix = Date.now() + '-' + Math.round(Math.random() * 1E9);
    cb(null, uniqueSuffix + path.extname(file.originalname));
  }
});

const upload = multer({ 
  storage: storage,
  limits: {
    fileSize: 100 * 1024 * 1024 // 100MB文件大小限制
  }
});

// 处理根路径请求
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// 转换Word文档
app.post('/convert-word', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: '未上传文件' });
    }
    
    console.log(`开始转换Word文件: ${req.file.originalname}`);
    const pdfBuffer = await wordToPdf(req.file.path);
    
    // 清理上传的文件
    fs.unlinkSync(req.file.path);
    
    // 设置响应头
    res.setHeader('Content-Type', 'application/pdf');
    res.send(pdfBuffer);
    
    console.log(`成功转换Word文件: ${req.file.originalname}`);
  } catch (error) {
    console.error('Word转换错误详情:', error);
    res.status(500).json({ 
      error: 'Word转换失败',
      detail: error.message,
      advice: [
        '1. 确保已安装LibreOffice',
        '2. 检查文档是否损坏',
        '3. 查看服务器日志获取详细信息'
      ]
    });
  }
});

// 转换Excel文档
app.post('/convert-excel', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: '未上传文件' });
    }
    
    console.log(`开始转换Excel文件: ${req.file.originalname}`);
    const pdfBuffer = await excelToPdf(req.file.path);
    
    // 清理上传的文件
    fs.unlinkSync(req.file.path);
    
    res.setHeader('Content-Type', 'application/pdf');
    res.send(pdfBuffer);
    
    console.log(`成功转换Excel文件: ${req.file.originalname}`);
  } catch (error) {
    console.error('Excel转换错误详情:', error);
    res.status(500).json({ 
      error: 'Excel转换失败',
      detail: error.message,
      advice: [
        '1. 确保已安装LibreOffice',
        '2. 检查文档是否包含复杂公式',
        '3. 查看服务器日志获取详细信息'
      ]
    });
  }
});

// 转换PowerPoint文档
app.post('/convert-powerpoint', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: '未上传文件' });
    }
    
    console.log(`开始转换PowerPoint文件: ${req.file.originalname}`);
    const pdfBuffer = await powerpointToPdf(req.file.path);
    
    // 清理上传的文件
    fs.unlinkSync(req.file.path);
    
    res.setHeader('Content-Type', 'application/pdf');
    res.send(pdfBuffer);
    
    console.log(`成功转换PowerPoint文件: ${req.file.originalname}`);
  } catch (error) {
    console.error('PowerPoint转换错误详情:', error);
    res.status(500).json({ 
      error: 'PowerPoint转换失败',
      detail: error.message,
      advice: [
        '1. 确保已安装LibreOffice',
        '2. 检查文档是否包含特殊动画',
        '3. 查看服务器日志获取详细信息'
      ]
    });
  }
});

// 启动服务器
app.listen(PORT, () => {
  console.log(`服务器运行在 http://localhost:${PORT}`);
  console.log('可用接口:');
  console.log(`  GET  /             - 前端页面`);
  console.log(`  POST /convert-word - Word转PDF`);
  console.log(`  POST /convert-excel - Excel转PDF`);
  console.log(`  POST /convert-powerpoint - PowerPoint转PDF`);
  
  // 检查LibreOffice是否可用
  const { exec } = require('child_process');
  const LIBREOFFICE_PATH = process.platform === 'win32' 
    ? '"C:\\Program Files\\LibreOffice\\program\\soffice.exe"'
    : '/usr/bin/soffice'; // 云服务器通常使用此路径
  
  exec(`${LIBREOFFICE_PATH} --version`, (error, stdout, stderr) => {
    if (error) {
      console.error('\x1b[31m%s\x1b[0m', '警告: LibreOffice未安装或路径不正确');
      console.error('请确保LibreOffice已安装并正确配置');
      console.error('Linux服务器安装命令: sudo apt-get install libreoffice');
      console.error('Windows服务器安装: https://www.libreoffice.org/download/download-libreoffice/');
    } else {
      console.log('\x1b[32m%s\x1b[0m', `LibreOffice检测正常: ${stdout.trim()}`);
    }
  });
});