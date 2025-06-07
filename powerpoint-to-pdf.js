const { exec } = require('child_process');
const fs = require('fs');
const path = require('path');

module.exports = function convertPowerPointToPdf(inputPath) {
  return new Promise((resolve, reject) => {
    const outputDir = path.dirname(inputPath);
    const outputName = path.basename(inputPath, path.extname(inputPath)) + '.pdf';
    const outputPath = path.join(outputDir, outputName);

    // 使用 LibreOffice 转换 PowerPoint 文件
    exec(`soffice --headless --convert-to pdf "${inputPath}" --outdir "${outputDir}"`, 
      (error, stdout, stderr) => {
        if (error) {
          console.error(`转换错误: ${error.message}`);
          return reject(error);
        }
        if (stderr) {
          console.error(`stderr: ${stderr}`);
        }

        // 检查生成的 PDF 文件
        if (!fs.existsSync(outputPath)) {
          return reject(new Error('转换失败，输出文件未生成'));
        }

        // 读取 PDF 文件内容
        const pdfBuffer = fs.readFileSync(outputPath);
        
        // 清理临时文件
        fs.unlinkSync(outputPath);
        
        resolve(pdfBuffer);
      }
    );
  });
};