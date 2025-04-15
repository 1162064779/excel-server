const express = require('express');
const ExcelJS = require('exceljs');
const path = require('path');

const app = express();
const PORT = 3000;

app.use(express.static(path.join(__dirname, 'public')));
app.use(express.json());

app.post('/generate-excel', async (req, res) => {
  const {
    name,
    amount,
    rate,
    lateRate,
    term,
    startDate,
    endDate,
    repayment
  } = req.body;

  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('借款明细');

  // 设置列宽（A-E）
  worksheet.columns = [
    { key: 'A', width: 11.25 },
    { key: 'B', width: 11.25 },
    { key: 'C', width: 11.25 },
    { key: 'D', width: 11.25 },
    { key: 'E', width: 15 },
  ];

// 数据行（每一项为一个 row 数组）
const rows = [
    ['第1笔借款明细表'],
    ['要素表'],
    ['基本要素'],
    ['借款人姓名', '', '', '', name],
    ['借款本金', '', '', '', Number(amount)],
    ['年利率', '', '', '', Number(rate) / 100],        // 百分比格式用小数
    ['逾期年利率', '', '', '', Number(lateRate) / 100], // 百分比格式用小数
    ['期限（月/期）', '', '', '', Number(term)],
    ['起息日', '', '', '', startDate],
    ['到期日', '', '', '', endDate],
    ['还款方式', '', '', '', repayment],
  ];
  
  // 插入行
  rows.forEach((row) => {
    worksheet.addRow(row);
  });
  
  // 合并单元格
  worksheet.mergeCells('A1:E1');
  worksheet.mergeCells('A2:E2');
  for (let i = 3; i <= 11; i++) {
    worksheet.mergeCells(`A${i}:D${i}`);
  }
  
  // 设置样式
  worksheet.eachRow((row, rowNumber) => {
    row.eachCell((cell, colNumber) => {
      // 所有单元格居中
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
  
      // 百分比格式（第6行和第7行，第5列即 E 列）
      if ((rowNumber === 6 || rowNumber === 7) && colNumber === 5) {
        cell.numFmt = '0.00%';
      }
  
      // 金额格式（第5行，第5列）
      if (rowNumber === 5 && colNumber === 5) {
        cell.numFmt = '#,##0.00';
      }
  
      // 其他可以根据需要添加格式
    });
  });

  // 导出为 buffer
  const buffer = await workbook.xlsx.writeBuffer();

  const safeFileName = encodeURIComponent(`test_xp.xlsx`);
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${safeFileName}`);
  res.send(buffer);
});

app.listen(PORT, () => {
  console.log(`✅ 服务运行中：http://localhost:${PORT}`);
});