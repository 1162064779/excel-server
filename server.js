const express = require('express');
const ExcelJS = require('exceljs');
const path = require('path');
const {exec} = require('child_process');
const os = require('os');
const dayjs = require('dayjs');
const customParseFormat = require('dayjs/plugin/customParseFormat');
dayjs.extend(customParseFormat);

function openUrl(url) {
  const platform = os.platform();

  let command;
  if (platform === 'win32') {
    command = `start "" "${url}"`;
  } else if (platform === 'darwin') {
    command = `open "${url}"`;
  } else if (platform === 'linux') {
    command = `xdg-open "${url}"`;
  } else {
    console.error('Unsupported OS');
    return;
  }

  exec(command, (err) => {
    if (err) {
      console.error('Failed to open URL:', err);
    }
  });
}

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
    repayment,
    repaymentType,
    interestDay
  } = req.body;

  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('借款明细');

  // 设置列宽（A-E）
  worksheet.columns = [
    {key: 'A', width: 11.25},
    {key: 'B', width: 11.25},
    {key: 'C', width: 11.25},
    {key: 'D', width: 11.25},
    {key: 'E', width: 15},
  ];

  // 数据行（每一项为一个 row 数组）
  const rows = [
    ['第1笔借款明细表'], ['要素表'], ['基本要素'],
    ['借款人姓名', '', '', '', name], ['借款本金', '', '', '', amount],
    ['年利率', '', '', '', rate / 100],  // 百分比格式用小数
    ['逾期年利率', '', '', '', lateRate / 100],  // 百分比格式用小数
    ['期限（月/期）', '', '', '', term],
    ['起息日', '', '', '', startDate], ['到期日', '', '', '', endDate],
    ['还款方式', '', '', '', repayment],  // 第11行
    ['', '', '', '', ''],                 // 第12行
    ['', '', '', '', ''],                 // 第13行
  ];

  // 插入行
  rows.forEach((row) => {
    worksheet.addRow(row);
  });

  // 合并单元格
  worksheet.mergeCells('A1:E1');
  worksheet.mergeCells('A2:E2');
  for (let i = 3; i <= 10; i++) {
    worksheet.mergeCells(`A${i}:D${i}`);
  }
  // 还款方式占三行（A11:D13, E11:E13）
  worksheet.mergeCells('A11:D13');
  worksheet.mergeCells('E11:E13');

  if (repaymentType == 1) {
    worksheet.insertRow(15, [
      '期数',
      '计算日',
      '起息日',
      '截息日',
      '计息天数',
      '应还本金金额',
      '利随本清利息',
      '应还利息金额',
      '已还本金金额',
      '已还利息金额',
      '累计未还本金金额',
      '累计未还利息金额',
      '复利（以当期未还利息为基数）',
      '当期未还利息金额',
      '复利利息标准（期内基准执行利率；期外逾期执行利率）',
      '复利起止期限',
      '',
      '计息天数',
      '罚息（以当期未还本金为基数）',
      '当期未还本金金额',
      '逾期利息标准',
      '罚息起止期限',
      '',
      '计息天数',
      '逾期利息（罚息+复利）',
      '已还逾期利息',
      '未还逾期利息'
    ]);

    worksheet.mergeCells('P15:Q15');
    worksheet.mergeCells('V15:W15');

    // 日期处理
    const startDateParsed =
        dayjs(startDate, ['YYYY年M月D日', 'YYYY/MM/DD', 'YYYY-MM-DD']);
    const endDateParsed =
        dayjs(endDate, ['YYYY年M月D日', 'YYYY/MM/DD', 'YYYY-MM-DD']);
    const interestDayInt = parseInt(interestDay, 10);

    const startRow = 16;

    for (let i = 0; i <= term; i++) {
      const row = worksheet.getRow(startRow + i);
      row.getCell(1).value = i;  // A列：期数

      if (i === 0) {
        // 第0期不填 C、D
        row.commit();
        continue;
      }

      const currentRow = startRow + i;
      const prevDRow = currentRow - 1;

      // === C列：起息日 ===
      if (i === 1) {
        // 第1期：直接用 startDate
        row.getCell(3).value = startDateParsed.format('YYYY/MM/DD');
      } else {
        // 从第2期开始：用公式引用上一行 D 列
        row.getCell(3).value = {formula: `D${prevDRow}`};
      }

      // === D列：结息日 ===
      if (i === term) {
        // 最后一期，D列为 endDate
        row.getCell(4).value = endDateParsed.format('YYYY/MM/DD');
      } else {
        // 结息日为当前行起息日所在月份的 interestDay
        const monthOffset = i - 1;  // 第1期加0月，第2期加1月...
        const tentative = startDateParsed.add(monthOffset, 'month')
                              .set('date', interestDayInt);

        // 如果 interestDayInt
        // 比起息日早，可能需要再加一个月（比如起息是28号，interestDay是21）
        let adjustedDate = tentative;
        if (tentative.isBefore(
                startDateParsed.add(monthOffset, 'month'), 'day')) {
          adjustedDate = tentative.add(1, 'month');
        }

        row.getCell(4).value = adjustedDate.format('YYYY/MM/DD');
      }

      row.commit();
    }
  }

  // 设置样式
  worksheet.eachRow((row, rowNumber) => {
    row.eachCell((cell, colNumber) => {
      // 所有单元格居中
      cell.alignment = {
        vertical: 'middle',
        horizontal: 'center',
        wrapText: true,
      };

      // 百分比格式（第6行和第7行，第5列即 E 列）
      if ((rowNumber === 6 || rowNumber === 7) && colNumber === 5) {
        cell.numFmt = '0.00%';
      }

      // 金额格式（第5行，第5列）
      if (rowNumber === 5 && colNumber === 5) {
        cell.numFmt = '#,##0.00';
      }
    });
  });

  // 导出为 buffer
  const buffer = await workbook.xlsx.writeBuffer();

  const safeFileName = encodeURIComponent(`test_xp.xlsx`);
  res.setHeader(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader(
      'Content-Disposition', `attachment; filename*=UTF-8''${safeFileName}`);
  res.send(buffer);
});

app.listen(PORT, () => {
  console.log(`✅ 服务运行中：http://localhost:${PORT}`);

  const url = `http://localhost:${PORT}`;
  // ✅ 自动打开默认浏览器访问
  openUrl(url);
});