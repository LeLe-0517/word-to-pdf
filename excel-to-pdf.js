const { PDFDocument, rgb } = require('pdf-lib');
const ExcelJS = require('exceljs');
const fs = require('fs');

async function excelToPdf(filePath) {
  try {
    const pdfDoc = await PDFDocument.create();
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);

    for (const sheet of workbook.worksheets) {
      const page = pdfDoc.addPage([600, 800]); // 设置页面大小
      const fontSize = 10;
      const textMargin = 50;
      let y = 750; // 从页面顶部开始

      // 添加工作表名称
      page.drawText(`${sheet.name}`, {
        x: textMargin,
        y: y,
        size: fontSize + 2,
        color: rgb(0, 0, 0),
      });
      y -= fontSize + 10;

      // 添加单元格数据
      const rows = sheet-usedRange().rows;
      for (const row of rows) {
        if (y <= 50) {
          // 如果当前页面没有空间，添加新页面
          const newPage = pdfDoc.addPage([600, 800]);
          y = 750;
        }
        let x = textMargin;
        for (const cell of row.cells) {
          page.drawText(`${cell.value || ''}`, {
            x: x,
            y: y,
            size: fontSize,
            color: rgb(0, 0, 0),
          });
          x += 100; // 列间距
        }
        y -= fontSize + 5;
      }
    }

    // 保存 PDF 并返回缓冲区
    const pdfBytes = await pdfDoc.save();
    return pdfBytes;
  } catch (error) {
    throw new Error('Excel to PDF conversion failed: ' + error.message);
  }
}

module.exports = excelToPdf;