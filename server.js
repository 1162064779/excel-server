const express = require('express');
const ExcelJS = require('exceljs');
const multer = require('multer');
const path = require('path');
const {exec} = require('child_process');
const os = require('os');
const dayjs = require('dayjs');
const customParseFormat = require('dayjs/plugin/customParseFormat');
dayjs.extend(customParseFormat);

const upload = multer();

const app = express();
const PORT = 3000;

app.use(express.static(path.join(__dirname, 'public')));

// ✅ 生成 Excel 列名，例如 A, B, ..., Z, AA, AB, ..., AZ, BA, BB, ...
function generateExcelColumnNames(count) {
  const columns = [];
  let i = 0;

  while (columns.length < count) {
    let name = '';
    let temp = i;

    do {
      name = String.fromCharCode((temp % 26) + 65) + name;
      temp = Math.floor(temp / 26) - 1;
    } while (temp >= 0);

    columns.push(name);
    i++;
  }

  return columns;
}

/**
 * 生成包含多个 sheet 的 Excel Buffer
 * @param {Array} groups - 传入多个 group，每个 group 有一个 sheet
 * @returns {Promise<Buffer>}
 */
async function generateExcelBuffer(groups = []) {
  const workbook = new ExcelJS.Workbook();

  for (const group of groups) {
    const {
      sheetNumber,
      name,
      amount,
      rate,
      lateRate,
      term,
      startDate,
      endDate,
      repayment,
      repaymentType,
      interestDay,
      intimeTerm,
      currentDate,
      paymentPairs
    } = group;

    const worksheet = workbook.addWorksheet(`sheet${sheetNumber}`);

    worksheet.columns = [
      {key: 'A', width: 11.25}, {key: 'B', width: 11.25},
      {key: 'C', width: 11.25}, {key: 'D', width: 11.25},
      {key: 'E', width: 15},    {key: 'F', width: 11.25},
      {key: 'G', width: 11.25}, {key: 'H', width: 11.25},
      {key: 'I', width: 11.25}, {key: 'J', width: 11.25},
      {key: 'K', width: 11.25}, {key: 'L', width: 11.25},
      {key: 'M', width: 11.25}, {key: 'N', width: 11.25},
      {key: 'O', width: 11.25}, {key: 'P', width: 11.25},
      {key: 'Q', width: 11.25}, {key: 'R', width: 11.25},
      {key: 'S', width: 11.25}, {key: 'T', width: 11.25},
      {key: 'U', width: 11.25}, {key: 'V', width: 11.25},
      {key: 'W', width: 11.25}, {key: 'X', width: 11.25},
      {key: 'Y', width: 11.25}, {key: 'Z', width: 11.25},
      {key: 'AA', width: 11.25}
    ];


    const rows = [
      ['第1笔借款明细表'],
      ['要素表'],
      ['基本要素'],
      ['借款人姓名', '', '', '', name],
      ['借款本金', '', '', '', amount],
      ['年利率', '', '', '', rate / 100],
      ['逾期年利率', '', '', '', lateRate / 100],
      ['期限（月/期）', '', '', '', term],
      ['起息日', '', '', '', startDate],
      ['到期日', '', '', '', endDate],
      ['还款方式', '', '', '', repayment],
      ['', '', '', '', ''],
      ['', '', '', '', ''],
    ];

    rows.forEach((row) => worksheet.addRow(row));

    worksheet.mergeCells('A1:E1');
    worksheet.mergeCells('A2:E2');
    for (let i = 3; i <= 10; i++) {
      worksheet.mergeCells(`A${i}:D${i}`);
    }
    worksheet.mergeCells('A11:D13');
    worksheet.mergeCells('E11:E13');

    worksheet.insertRow(15, [
      '期数',
      '期末本金余额',
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
      '复利（基于未还利息）',
      '当期未还利息金额',
      '复利利息标准（期内基础执行利率；期外逾期执行利率',
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

    // 起息日、到期日、开始行
    const startDateParsed = dayjs(startDate);
    const endDateParsed = dayjs(endDate);
    const currentDateParsed = dayjs(currentDate);
    const startRow = 16;

    // 期数 0，空着
    const row0 = worksheet.getRow(startRow);
    row0.getCell(1).value = 0;
    row0.commit();

    // 检查 intimeTerm
    if (intimeTerm < 0 || intimeTerm > term) {
      throw new Error(`提前还款期限 intimeTerm(${
          intimeTerm}) 超出范围，应在 0-${term} 之间`);
    }

    const originRowNumbers = [];    // 记录初始 period 的真实行号
    const insertedRowNumbers = [];  // 记录插入子期的真实行号

    let firstInterestEndDate;
    // 确定读取 periods 的起始行
    let periodsStartRow = intimeTerm > 0 ? startRow + 2 : startRow + 1;
    let periodsEndRow = -1;

    // 生成提前还款 (只生成一行)
    if (intimeTerm > 0) {
      const row = worksheet.getRow(startRow + 1);

      // 期数列
      if (intimeTerm === 1) {
        row.getCell(1).value = 1;
      } else {
        row.getCell(1).value = `1-${intimeTerm}`;
      }

      // 起息日
      row.getCell(3).value = startDateParsed.format('YYYY/MM/DD');

      // 计算第一段的结息日
      let tentativeEndDate = startDateParsed.add(intimeTerm, 'month');

      // 如果起息日的“日”大于 interestDay，需要额外再推1个月
      if (startDateParsed.date() > interestDay) {
        tentativeEndDate = tentativeEndDate.add(1, 'month');
      }

      // 最后把日子设置成固定的 interestDay
      tentativeEndDate = tentativeEndDate.set('date', interestDay);

      firstInterestEndDate = tentativeEndDate;

      row.getCell(4).value = firstInterestEndDate.format('YYYY/MM/DD');

      row.commit();

      originRowNumbers.push(startRow + 1);
    }

    let isBeforeCurrent = true;

    // 生成剩余期数，从 intimeTerm+1 到 term
    for (let i = intimeTerm + 1; i <= term; i++) {
      const rowIndex = startRow + (i - intimeTerm) + 1;
      const row = worksheet.getRow(rowIndex);

      // 期数列
      row.getCell(1).value = i;

      // 起息日
      if (i === intimeTerm + 1) {
        row.getCell(3).value = firstInterestEndDate.format('YYYY/MM/DD');
      } else {
        const prevEndDate = worksheet.getRow(rowIndex - 1).getCell(4).value;
        row.getCell(3).value = prevEndDate;
      }

      // 结息日

      const prevEndDateStr = worksheet.getRow(rowIndex).getCell(3).value;
      const prevEndDateParsed = dayjs(prevEndDateStr);

      let nextEndDate =
          prevEndDateParsed.add(1, 'month').set('date', interestDay);

      if (nextEndDate.isBefore(prevEndDateParsed.add(1, 'month'), 'day')) {
        nextEndDate = nextEndDate.add(1, 'month');
      }

      if (nextEndDate.isAfter(endDateParsed, 'day')) {
        nextEndDate = endDateParsed;
        row.getCell(4).value = nextEndDate.format('YYYY/MM/DD');
        console.log(`超过总 endDate，使用最终 endDate ${
            nextEndDate.format('YYYY/MM/DD')}`);
        periodsEndRow = rowIndex;
        break;
      } else if (nextEndDate.isAfter(currentDateParsed, 'day')) {
        nextEndDate = currentDateParsed;
        row.getCell(4).value = nextEndDate.format('YYYY/MM/DD');
        console.log(`超过 currentDate，使用当前 currentDate ${
            nextEndDate.format('YYYY/MM/DD')}`);
        periodsEndRow = rowIndex;
        isBeforeCurrent = false;
        break;
      } else {
        row.getCell(4).value = nextEndDate.format('YYYY/MM/DD');
      }
      row.commit();
    }

    // 读取 periods（已有）
    const periods = [];
    for (let i = periodsStartRow; i <= periodsEndRow; i++) {
      const row = worksheet.getRow(i);
      const period = {
        row: i,
        period: row.getCell(1).value,
        start: dayjs(row.getCell(3).value),
        end: dayjs(row.getCell(4).value),
      };
      periods.push(period);
    }

    // 排序 paymentPairs
    paymentPairs.sort(
        (a, b) => dayjs(a.date).isAfter(dayjs(b.date), 'day') ? 1 : -1);

    console.log(`当前 paymentPairs 列表:`);
    paymentPairs.forEach((pair, index) => {
      console.log(
          `  [${index}] 日期: ${dayjs(pair.date).format('YYYY/MM/DD')}, 金额: ${
              pair.value}, 类型: ${pair.type}`);
    });

    const newPeriods = [];
    let paymentIndex = 0;  // paymentPairs处理到的位置

    const interestRowNumbers = [];
    const principalRowNumbers = [];
    const overdueInterestRowNumbers = [];
    const lastPeriodRowNumbers = [];

    for (let i = 0; i < periods.length; i++) {
      const currentPeriod = periods[i];
      const nextPeriod = periods[i + 1];  // 可能没有，记得判断

      console.log(`\n处理 period ${currentPeriod.period}: ${
          currentPeriod.start.format(
              'YYYY/MM/DD')} ~ ${currentPeriod.end.format('YYYY/MM/DD')}`);


      let subIndex = 1;  // 子期编号，例如 (1)、(2)、(3) ...

      // 只在 i == 0 时额外处理
      if (i === 0) {
        const firstPayment = paymentPairs[paymentIndex];
        if (firstPayment) {
          let paymentDate = dayjs(firstPayment.date);

          while (paymentIndex < paymentPairs.length &&
                 paymentDate.isBefore(currentPeriod.end, 'day')) {
            const prevSubPeriodName =
                `${currentPeriod.period - 1}(${subIndex})`;
            const lastEndDate = newPeriods.length > 0 ?
                newPeriods[newPeriods.length - 1].end :
                currentPeriod.start;

            console.log(`插入前置子期 ${prevSubPeriodName}: ${
                lastEndDate.format(
                    'YYYY/MM/DD')} ~ ${paymentDate.format('YYYY/MM/DD')}`);

            const currentPayment = paymentPairs[paymentIndex];

            if (!currentPayment) {
              throw new Error(
                  `paymentPairs[paymentIndex=${paymentIndex}] 为空，无法处理`);
            }

            const newPeriod = {
              period: prevSubPeriodName,
              start: lastEndDate,
              end: paymentDate,
            };

            newPeriods.push(newPeriod);

            const lastNewPeriod = newPeriods[newPeriods.length - 1];
            const rowNumber = startRow + 1 + newPeriods.length;

            if (currentPayment.type === '利息') {
              lastNewPeriod.interest = currentPayment.value;
              interestRowNumbers.push(rowNumber);
            } else if (currentPayment.type === '本金') {
              lastNewPeriod.principal = currentPayment.value;
              principalRowNumbers.push(rowNumber);
            } else if (currentPayment.type === '逾期利息') {
              lastNewPeriod.overdueInterest = currentPayment.value;
              overdueInterestRowNumbers.push(rowNumber);
            } else {
              throw new Error(`未知的 currentPayment.type: ${
                  currentPayment.type}，无法处理！`);
            }

            insertedRowNumbers.push(rowNumber);

            subIndex++;
            paymentIndex++;  // 处理下一个 payment

            // 准备下一个 payment
            if (paymentIndex < paymentPairs.length) {
              const nextPayment = paymentPairs[paymentIndex];
              if (nextPayment) {
                paymentDate = dayjs(nextPayment.date);

                if (!paymentDate.isBefore(currentPeriod.end, 'day')) {
                  break;
                }
              } else {
                break;
              }
            } else {
              break;
            }
          }
        }
      }
      // 先插入原 period
      newPeriods.push({
        period: currentPeriod.period,
        start: currentPeriod.start,
        end: currentPeriod.end,
      });
      originRowNumbers.push(startRow + 1 + newPeriods.length);

      subIndex = 1;  // 子期编号，例如 (1)、(2)、(3) ...

      // 如果还有 paymentPairs 没处理完
      while (paymentIndex < paymentPairs.length) {
        const currentPayment = paymentPairs[paymentIndex];
        const paymentDate = dayjs(currentPayment.date);

        // 处理还款日期和截止日相同的情况
        if (paymentDate.isSame(currentPeriod.end, 'day')) {
          const lastNewPeriod = newPeriods[newPeriods.length - 1];

          if (!lastNewPeriod) {
            throw new Error(
                `当前 newPeriods 为空，无法给最后一个元素赋值！period: ${
                    currentPeriod.period}`);
          }

          if (currentPayment.type === '利息') {
            lastNewPeriod.interest = currentPayment.value;
            interestRowNumbers.push(startRow + 1 + newPeriods.length);
          } else if (currentPayment.type === '本金') {
            lastNewPeriod.principal = currentPayment.value;
            principalRowNumbers.push(startRow + 1 + newPeriods.length);
          } else if (currentPayment.type === '逾期利息') {
            lastNewPeriod.overdueInterest = currentPayment.value;
            overdueInterestRowNumbers.push(startRow + 1 + newPeriods.length);
          } else {
            throw new Error(`未知的 currentPayment.type: ${
                currentPayment.type}，无法处理！`);
          }

          console.log(`更新最后一个 newPeriod [period: ${
              lastNewPeriod.period}]，增加字段 ${currentPayment.type}: ${
              currentPayment.value}`);

          paymentIndex++;
          continue;
        }

        // 判断是否要插入到当前period下面
        if (nextPeriod) {
          if (paymentDate.isBefore(nextPeriod.end, 'day')) {
            // 插入子期
            const lastEndDate =
                newPeriods[newPeriods.length - 1].end;  // 上一个period的end

            const subPeriodName = `${currentPeriod.period}(${subIndex})`;

            console.log(`插入子期 ${subPeriodName}: ${
                lastEndDate.format(
                    'YYYY/MM/DD')} ~ ${paymentDate.format('YYYY/MM/DD')}`);

            const newPeriod = {
              period: subPeriodName,
              start: lastEndDate,
              end: paymentDate,
            };
            newPeriods.push(newPeriod);

            const lastNewPeriod = newPeriods[newPeriods.length - 1];
            const rowNumber = startRow + 1 + newPeriods.length;

            if (currentPayment.type === '利息') {
              lastNewPeriod.interest = currentPayment.value;
              interestRowNumbers.push(rowNumber);
            } else if (currentPayment.type === '本金') {
              lastNewPeriod.principal = currentPayment.value;
              principalRowNumbers.push(rowNumber);
            } else if (currentPayment.type === '逾期利息') {
              lastNewPeriod.overdueInterest = currentPayment.value;
              overdueInterestRowNumbers.push(rowNumber);
            } else {
              throw new Error(`未知的 currentPayment.type: ${
                  currentPayment.type}，无法处理！`);
            }

            insertedRowNumbers.push(rowNumber);

            subIndex++;
            paymentIndex++;  // 移动到下一个 paymentPair
          } else {
            // 当前 paymentPair 不属于这个 period，停止 while，去处理下一个
            // period
            break;
          }
        } else {
          // 如果已经是最后一个 period（没有 nextPeriod了）
          // 那么剩下的 paymentPairs 都归到最后一个 period处理
          const lastEndDate = newPeriods[newPeriods.length - 1].end;

          const subPeriodName = `${currentPeriod.period}(${subIndex})`;

          console.log(`最后插入子期 ${subPeriodName}: ${
              lastEndDate.format(
                  'YYYY/MM/DD')} ~ ${paymentDate.format('YYYY/MM/DD')}`);

          const newPeriod = {
            period: subPeriodName,
            start: lastEndDate,
            end: paymentDate,
          };
          newPeriods.push(newPeriod);

          const lastNewPeriod = newPeriods[newPeriods.length - 1];
          const rowNumber = startRow + 1 + newPeriods.length;

          if (currentPayment.type === '利息') {
            lastNewPeriod.interest = currentPayment.value;
            interestRowNumbers.push(rowNumber);
          } else if (currentPayment.type === '本金') {
            lastNewPeriod.principal = currentPayment.value;
            principalRowNumbers.push(rowNumber);
          } else if (currentPayment.type === '逾期利息') {
            lastNewPeriod.overdueInterest = currentPayment.value;
            overdueInterestRowNumbers.push(rowNumber);
          } else {
            throw new Error(`未知的 currentPayment.type: ${
                currentPayment.type}，无法处理！`);
          }
          lastPeriodRowNumbers.push(rowNumber);
          insertedRowNumbers.push(rowNumber);

          subIndex++;
          paymentIndex++;  // 移动到下一个 paymentPair
        }
      }
    }

    // 重新写回 worksheet
    let currentRowIdx = periodsStartRow;
    for (const newPeriod of newPeriods) {
      const row = worksheet.getRow(currentRowIdx);

      row.getCell(1).value = newPeriod.period;
      row.getCell(3).value = newPeriod.start.format('YYYY/MM/DD');
      row.getCell(4).value = newPeriod.end.format('YYYY/MM/DD');

      if (newPeriod.principal !== undefined) {
        row.getCell(9).value = newPeriod.principal;
      }

      if (newPeriod.interest !== undefined) {
        row.getCell(10).value = newPeriod.interest;
      }

      if (newPeriod.overdueInterest !== undefined) {
        row.getCell(26).value = newPeriod.overdueInterest;
      }

      console.log(`写入 Row ${currentRowIdx}：期数 ${newPeriod.period}，${
          newPeriod.start.format(
              'YYYY/MM/DD')} ~ ${newPeriod.end.format('YYYY/MM/DD')}`);

      row.commit();
      currentRowIdx++;
    }

    // 到期后
    if (isBeforeCurrent) {
      const row = worksheet.getRow(currentRowIdx);

      row.getCell(1).value = '到期后';
      row.getCell(3).value = endDateParsed.format('YYYY/MM/DD');
      row.getCell(4).value = currentDateParsed.format('YYYY/MM/DD');

      console.log(`写入 Row ${currentRowIdx}：到期后，${
          endDateParsed.format(
              'YYYY/MM/DD')} ~ ${currentDateParsed.format('YYYY/MM/DD')}`);

      row.commit();
    }

    console.log('原始 period 行号:', originRowNumbers);
    console.log('插入子期行号:', insertedRowNumbers);

    console.log('利息子期行号:', interestRowNumbers);
    console.log('本金子期行号:', principalRowNumbers);
    console.log('逾期利息子期行号:', overdueInterestRowNumbers);

    const firstRowIdx = startRow;
    const lastRowIdx = currentRowIdx - 1;

    // 单独处理第零行
    {
      const firstRow = worksheet.getRow(firstRowIdx);

      // 读取 E5 的值
      const e5Value = worksheet.getRow(5).getCell(5).value;

      if (e5Value === undefined || e5Value === null) {
        throw new Error(`E5单元格为空，无法赋值到第${firstRowIdx}行的B列`);
      }

      // B列写公式 =E5
      firstRow.getCell(2).value = {formula: `$E$5`};

      firstRow.commit();
    }

    // 处理第二行到倒数第二行
    for (let rowIdx = startRow + 1; rowIdx <= lastRowIdx; rowIdx++) {
      const row = worksheet.getRow(rowIdx);

      // 设置 E 列（第 5 列）为公式：=（D列 - C列的天数差）
      row.getCell(5).value = {formula: `D${rowIdx}-C${rowIdx}`};

      // 在B列写入公式 B(x) = B(x-1) - F(x)
      row.getCell(2).value = {formula: `B${rowIdx - 1}-F${rowIdx}`};

      // 设置 M 列（第 13 列）的公式: M(x) = N(x) * O(x) / 360 * R(x)
      row.getCell(13).value = {formula: `N${rowIdx}*O${rowIdx}/360*R${rowIdx}`};

      // O列写公式 = $E$6
      const oCell = row.getCell(15);  // O列是第15列
      if (lastPeriodRowNumbers.includes(rowIdx)) {
        oCell.value = {formula: `$E$7`};
      } else {
        oCell.value = {formula: `$E$6`};
      }

      // 使用 Excel 公式设置 R 列（第 18 列）：=Qx - Px
      row.getCell(18).value = {formula: `Q${rowIdx}-P${rowIdx}`};

      // 设置 Y 列（第 25 列）的公式：Y(x) = S(x) + M(x)
      row.getCell(25).value = {formula: `S${rowIdx}+M${rowIdx}`};

      // 设置 AA 列（第 27 列）的公式：AA(x) = Y(x) - Z(x)
      row.getCell(27).value = {formula: `Y${rowIdx}-Z${rowIdx}`};

      row.commit();
    }

    // 单独处理最后一行
    {
      const lastRow = worksheet.getRow(lastRowIdx + 1);



      lastRow.commit();
    }


    if (repaymentType === 1) {  // 先息后本

      // 单独处理第一行
      {
        const firstRow = worksheet.getRow(firstRowIdx + 1);
        firstRow.getCell(10).value = {formula: `H${firstRowIdx + 1}`};
        firstRow.commit();
      }

      // 处理第一行到倒数第二行
      for (let rowIdx = startRow + 1; rowIdx <= lastRowIdx; rowIdx++) {
        const row = worksheet.getRow(rowIdx);

        // 如果这一行是 originRowNumbers 中的
        if (originRowNumbers.includes(rowIdx)) {
          // F、I、K列都设为0
          row.getCell(6).value = 0;   // F列
          row.getCell(9).value = 0;   // I列
          row.getCell(11).value = 0;  // K列

          // H列加公式
          row.getCell(8).value = {
            formula: `B${rowIdx}*O${rowIdx}/360*E${rowIdx}`
          };
        }

        // 给L列加公式：L(x) = L(x-1) + H(x) - J(x)
        row.getCell(12).value = {
          formula: `L${rowIdx - 1}+H${rowIdx}-J${rowIdx}`
        };

        // M列加公式：M(x) = N(x) * O(x) / 360 * R(x)
        row.getCell(13).value = {
          formula: `N${rowIdx}*O${rowIdx}/360*R${rowIdx}`
        };

        const prevRow = rowIdx - 1;

        // 设置第 14 列（N列）的公式
        if (lastPeriodRowNumbers.includes(rowIdx)) {
          row.getCell(14).value = {
            formula: `L${prevRow}`  // 公式：N(x) = L(x-1)
          };
        } else if (
            interestRowNumbers.includes(prevRow) &&
            !originRowNumbers.includes(prevRow)) {
          row.getCell(14).value = {
            formula: `N${prevRow}-J${prevRow}`  // 公式：N(x) = N(x-1) - J(x-1)
          };
        } else {
          row.getCell(14).value = {
            formula: `H${prevRow}-J${prevRow}`  // 公式：N(x) = H(x-1) - J(x-1)
          };
        }

        // P列（第16列）和 Q列（第17列）
        if (interestRowNumbers.includes(rowIdx) &&
            !originRowNumbers.includes(rowIdx)) {
          row.getCell(16).value = {
            formula: `C${rowIdx}`  // P(x) = c(x)
          };
          row.getCell(17).value = {
            formula: `D${rowIdx}`  // Q(x) = D(x)
          };
        } else {
          row.getCell(16).value = {
            formula: `C${rowIdx}`  // P(x) = c(x)
          };
          row.getCell(17).value = {
            formula: `$E$10`  // Q(x) = E10
          };
        }

        row.commit();
      }
    }

    worksheet.eachRow((row, rowNumber) => {
      row.eachCell((cell, colNumber) => {
        // 设置通用对齐样式
        cell.alignment = {
          vertical: 'middle',
          horizontal: 'center',
          wrapText: true
        };

        // 设置百分比格式
        if ((rowNumber === 6 || rowNumber === 7) && colNumber === 5) {
          cell.numFmt = '0.00%';
        }

        if (colNumber === 15 || colNumber === 21) {
          cell.numFmt = '0.00%';
        }

        // 两位小数+千分号（行5，列5）
        if (rowNumber === 5 && colNumber === 5) {
          cell.numFmt = '#,##0.00';
        }

        // 两位小数+千分号（统一用 '#,##0.00'）：B, F, H, I, J, K, L, M, N, S,
        // T, X, Y, Z, AA
        const decimalColumns =
            [2, 6, 8, 9, 10, 11, 12, 13, 14, 19, 20, 24, 25, 26, 27];
        if (decimalColumns.includes(colNumber)) {
          cell.numFmt = '#,##0.00';
        }

        // 设置日期格式的列：C, D, P, Q, V, W
        const dateColumns = [3, 4, 16, 17, 22, 23];
        if (dateColumns.includes(colNumber)) {
          cell.numFmt = 'YYYY/MM/DD';
        }
      });
    });
  }

  return await workbook.xlsx.writeBuffer();
}

// =====================
// 文件上传并生成 Excel
// =====================
app.post('/generate-excel', upload.single('file'), async (req, res) => {
  try {
    const {currentDate} = req.body;
    const buffer = req.file.buffer;

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);
    console.log(`工作簿包含的工作表数量: ${workbook.worksheets.length}`);
    if (workbook.worksheets.length === 0) {
      throw new Error('Excel 文件中没有任何工作表。');
    }
    const sheet = workbook.worksheets[0];
    console.log(`读取第一个工作表，名称为: ${sheet.name}`);
    console.log(`工作表行数: ${sheet.rowCount}, 列数: ${sheet.columnCount}`);

    const name = sheet.getCell('B1').value || '';
    const groups = [];
    const maxGroups = 30;  // 最多支持 30 组（你可以修改为任意值）

    // 每组占 3 列，且中间留 1 列空白 => 每组跨度为 4 列
    const totalColumnsNeeded = maxGroups * 4;
    const allColumns = generateExcelColumnNames(totalColumnsNeeded);

    for (let groupIndex = 0; groupIndex < maxGroups; groupIndex++) {
      const baseIdx = 3 + groupIndex * 4;

      const colDate = allColumns[baseIdx];  // 日期列
      const colAmount = allColumns[baseIdx + 1];  // 金额列（也是参数基准列）
      const colType = allColumns[baseIdx + 2];  // 类型列

      console.log(` 正在处理第 ${groupIndex + 1} 组，列分别为：日期列 = ${
          colDate}，金额列 = ${colAmount}，类型列 = ${colType}`);

      // 参数区域（第 4~12 行）以金额列为基准
      const amount = parseFloat(sheet.getCell(`${colAmount}4`).value) || 0;

      //  如果关键字段为空，说明没有这一组了，跳出循环
      if (!amount) {
        break;
      }

      console.log(`第 ${groupIndex + 1} 组 金额 (${colAmount}4): ${amount}`);

      const rate = parseFloat(sheet.getCell(`${colAmount}5`).value) * 100 || 0;
      console.log(`第 ${groupIndex + 1} 组 利率 (${colAmount}5): ${rate}%`);

      const lateRate =
          parseFloat(sheet.getCell(`${colAmount}6`).value) * 100 || 0;
      console.log(
          `第 ${groupIndex + 1} 组 罚息利率 (${colAmount}6): ${lateRate}%`);

      const term = parseInt(sheet.getCell(`${colAmount}7`).value) || 0;
      console.log(`第 ${groupIndex + 1} 组 期限 (${colAmount}7): ${term} 月`);

      const startDate = sheet.getCell(`${colAmount}8`).value || '';
      console.log(
          `第 ${groupIndex + 1} 组 起始日 (${colAmount}8): ${startDate}`);

      const endDate = sheet.getCell(`${colAmount}9`).value || '';
      console.log(`第 ${groupIndex + 1} 组 到期日 (${colAmount}9): ${endDate}`);

      const repaymentCell = sheet.getCell(`${colAmount}10`).value || '';
      console.log(
          `第 ${groupIndex + 1} 组 还款方式原始值 (${colAmount}10):`,
          repaymentCell);

      const repayment =
          (typeof repaymentCell === 'object' && repaymentCell.richText) ?
          repaymentCell.richText.map(part => part.text).join('').trim() :
          String(repaymentCell || '').trim();

      console.log(`第 ${groupIndex + 1} 组 还款方式解析后: ${repayment}`);

      let repaymentType = null;
      if (repayment.includes('每月付息')) {
        repaymentType = 1;
      } else {
        throw new Error(`第 ${
            groupIndex +
            1} 组的还款方式必须为「每月付息」，但实际为：${repayment}`);
      }

      const interestDay = parseInt(sheet.getCell(`${colAmount}11`).value) || 21;
      console.log(
          `第 ${groupIndex + 1} 组 结息日 (${colAmount}11): ${interestDay} 日`);

      const intimeTerm = parseInt(sheet.getCell(`${colAmount}12`).value) || 0;
      console.log(`第 ${groupIndex + 1} 组 提前还款期限 (${colAmount}12): ${
          intimeTerm} 天`);

      //  从第13行开始读取还款明细
      const paymentPairs = [];
      let row = 13;

      while (true) {
        const dateCell = sheet.getCell(`${colDate}${row}`).value;
        const valueCell = sheet.getCell(`${colAmount}${row}`).value;
        const typeCell = sheet.getCell(`${colType}${row}`).value;

        if (valueCell === null || valueCell === undefined || valueCell === '') {
          break;
        }

        paymentPairs.push({
          date: dateCell,
          value: parseFloat(valueCell),
          type: (typeCell || '').toString().trim()
        });

        row++;
      }

      // ✅ 添加当前组数据
      groups.push({
        sheetNumber: groupIndex + 1,
        name,
        amount,
        rate,
        lateRate,
        term,
        startDate,
        endDate,
        repayment,
        interestDay,
        intimeTerm,
        repaymentType,
        paymentPairs
      });
    }

    const outputBuffer = await generateExcelBuffer(groups);

    const fileName = '借款明细.xlsx';
    res.setHeader(
        'Content-Type',
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');

    // 设置 Content-Disposition，兼容大部分浏览器，包括中文文件名
    res.setHeader(
        'Content-Disposition',
        `attachment; filename="${
            encodeURIComponent(
                fileName)}"; filename*=UTF-8''${encodeURIComponent(fileName)}`);

    res.send(outputBuffer);
  } catch (err) {
    console.error('生成失败:', err);
    res.status(500).send(`生成失败: ${err.message}`);
  }
});

// =====================
// 启动服务
// =====================
app.listen(PORT, () => {
  const url = `http://localhost:${PORT}`;
  console.log(`✅ 服务运行中：${url}`);
  openUrl(url);
});

// =====================
// 打开浏览器
// =====================
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
    console.error('不支持的操作系统');
    return;
  }
  exec(command, (err) => {
    if (err) console.error('打开浏览器失败:', err);
  });
}