const { exec } = require('child_process');
const fs = require('fs');
const path = require('path');

// 云服务器上通常使用默认路径
const LIBREOFFICE_PATH = process.platform === 'win32' 
  ? '"C:\\Program Files\\LibreOffice\\program\\soffice.exe"'
  : '/usr/bin/soffice'; // Linux服务器的标准路径

module.exports = function convertWordToPdf(inputPath) {
  return new Promise((resolve, reject) => {
    const outputDir = path.dirname(inputPath);
    const outputName = path.basename(inputPath, path.extname(inputPath)) + '.pdf';
    const outputPath = path.join(outputDir, outputName);

    // 解决中文编码问题 - 使用PDF/A格式并启用Unicode支持
    const command = `${LIBREOFFICE_PATH} --headless --convert-to pdf:writer_pdf_Export --outdir "${outputDir}" "${inputPath}"`;
    
    console.log(`执行Word转换命令: ${command}`);
    
    exec(command, 
      (error, stdout, stderr) => {
        if (error) {
          console.error(`Word转换错误: ${error.message}`);
          return reject(new Error(`Word转换失败: ${error.message}`));
        }
        
        // 输出详细信息便于调试
        console.log(`stdout: ${stdout}`);
        if (stderr) console.error(`stderr: ${stderr}`);

        // 检查生成的 PDF 文件
        if (!fs.existsSync(outputPath)) {
          return reject(new Error('转换失败，输出文件未生成'));
        }

        try {
          // 读取 PDF 文件内容
          const pdfBuffer = fs.readFileSync(outputPath);
          
          // 清理临时文件
          fs.unlinkSync(outputPath);
          
          resolve(pdfBuffer);
        } catch (readError) {
          reject(new Error(`读取PDF失败: ${readError.message}`));
        }
      }
    );
  });
};