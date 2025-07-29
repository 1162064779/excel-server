const express = require('express');
const ExcelJS = require('exceljs');
const multer = require('multer');
const path = require('path');
const {exec} = require('child_process');
const os = require('os');
const dayjs = require('dayjs');
const customParseFormat = require('dayjs/plugin/customParseFormat');
const JSZip = require('jszip');
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
async function generateExcelBuffer(groups = [], infoGroup = {}) {
  const workbook = new ExcelJS.Workbook();
  // 设置工作簿的全局计算模式为自动
  workbook.calcProperties.fullCalcOnLoad = true;  // ⚡关键设置

  // 生成总表
  const sumRowNumbers = [];  // 记录每个子表的总结行号
  const summarySheet = workbook.addWorksheet('总表');
  summarySheet.columns = [
    {key: 'A', width: 13.25}, {key: 'B', width: 13.25},
    {key: 'C', width: 13.25}, {key: 'D', width: 13.25},
    {key: 'E', width: 13.25}, {key: 'F', width: 13.25},
    {key: 'G', width: 13.25}, {key: 'H', width: 13.25},
    {key: 'I', width: 13.25}, {key: 'J', width: 13.25},
    {key: 'K', width: 13.25}, {key: 'L', width: 13.25},
    {key: 'M', width: 13.25}, {key: 'N', width: 13.25},
    {key: 'O', width: 13.25}, {key: 'P', width: 13.25},
    {key: 'Q', width: 13.25}, {key: 'R', width: 13.25}, {key: 'S', width: 13.25}
  ];
  // 插入标题行（a = name，b = currentDate）
  const a = groups?.[0]?.name || '未知借款人';
  const rawDate = groups?.[0]?.currentDate || '';
  const rawDateStr = typeof rawDate === 'string' ?
      rawDate :
      dayjs(rawDate).format('YYYY/MM/DD');  // 如果是 Date 对象，用 dayjs 格式化
  const b = formatChineseDate(rawDate);
  const title = `【${a}】 借款计算表【总表】--暂计至${b}`;

  // 添加第一行作为标题
  summarySheet.addRow([title]);
  summarySheet.mergeCells('A1', 'R1');
  summarySheet.getCell('A1')
      .font = {bold: true, size: 16, color: {argb: 'FF000000'}};
  summarySheet.getCell('A1').alignment = {
    vertical: 'middle',
    horizontal: 'center'
  };

  summarySheet.addRow([
    '笔数', '起息日期', '账单到期日', '提前到期日（开庭日）', '借期数',
    '还款方式', '借款本金（元）', '年利率', '逾期利率', '已还本金（元）',
    '应还利息（元）', '已还利息（元）', '已还逾期利息（元）', '未还本金（元）',
    '未还利息（元）', '复利（以未还利息为基数）', '罚息（以未还本金为基数）',
    '未还逾期利息'
  ]);

  // 生成n个子表
  for (const group of groups) {
    let {
      sheetNum,
      sheetName,
      name,
      amount,
      rate,
      lateRate,
      term,
      startDate,
      endDate,
      isLarger30,
      repayment,
      interestDay,
      intimeTerm,
      currentDate,
      repaymentType,
      paymentPairs
    } = group;

    const worksheet = workbook.addWorksheet(sheetName);

    worksheet.columns = [
      {key: 'A', width: 15},    {key: 'B', width: 11.25},
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
      [`第${sheetNum}笔借款明细表`],
      ['要素表'],
      ['基本要素', '', '', '', '借款情况'],
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
      '期初本金余额',
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

    // 不是先息后本时，intime当作0来生成行
    let realIntimeTerm = 0;
    if (repaymentType != 1) {
      realIntimeTerm = intimeTerm;
      intimeTerm = 0;
    }

    let firstInterestEndDate = startDateParsed;
    // 确定读取 periods 的起始行
    let intimeTermOffset = intimeTerm > 0 ? 1 : 0;
    let periodsStartRow = startRow + intimeTermOffset + 1;
    let periodsEndRow = -1;

    // 如果设置第一期大于30天，就再加一个月
    if(!isLarger30){
      term = term + 1;
    }

    if (startDateParsed.date() === interestDay) {
      if(!isLarger30){
        term = term - 1;
        intimeTerm = intimeTerm - 1;
      }
      isLarger30 = false;
    }
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

      let tentativeEndDate =
          startDateParsed.add(intimeTerm - 1, 'month').date(interestDay);
      // 计算“下一个结息日”
      let nextInterestDate = startDateParsed.date(interestDay);
      if (!nextInterestDate.isAfter(startDateParsed)) {
        // 如果当前月的 interestDay 已经过了，则取下个月的 interestDay
        nextInterestDate = nextInterestDate.add(1, 'month');
        tentativeEndDate = tentativeEndDate.add(1, 'month');
      }

      // 判断差值
      const diffDays = nextInterestDate.diff(startDateParsed, 'day');

      if (isLarger30) {
        tentativeEndDate = tentativeEndDate.add(1, 'month');
      }

      firstInterestEndDate = tentativeEndDate.isAfter(endDateParsed) ?
          endDateParsed :
          tentativeEndDate;

      row.getCell(4).value = firstInterestEndDate.format('YYYY/MM/DD');

      row.commit();

      originRowNumbers.push(startRow + 1);
    }

    let isBeforeCurrent = true;

    // 生成剩余期数，从 intimeTerm+1 到 term
    for (let i = intimeTerm + 1; i <= term; i++) {
      const rowIndex = startRow + (i - intimeTerm) + intimeTermOffset;
      const row = worksheet.getRow(rowIndex);

      console.log(`\n 第 ${i} 期（Excel 第 ${rowIndex} 行）`);

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

      let nextEndDate;
      if (i === intimeTerm + 1) {
        nextEndDate = prevEndDateParsed.date(interestDay);
        if (!nextEndDate.isAfter(prevEndDateParsed)) {
          // 如果当前月的 interestDay 已经过了，则取下个月的 interestDay
          nextEndDate = nextEndDate.add(1, 'month');
        }
        if (isLarger30) {
          // 如果设置第一期大于30天，就再加一个月
          nextEndDate = nextEndDate.add(1, 'month');
        }
      } else {
        nextEndDate = prevEndDateParsed.add(1, 'month');
      }

      if (nextEndDate.isAfter(endDateParsed, 'day')) {
        nextEndDate = endDateParsed;
        row.getCell(4).value = nextEndDate.format('YYYY/MM/DD');
        console.log(`超过总 endDate，使用最终 endDate ${
            nextEndDate.format('YYYY/MM/DD')}`);
        periodsEndRow = rowIndex;
        break;
      } else if (
          nextEndDate.isAfter(currentDateParsed, 'day') ||
          nextEndDate.isSame(currentDateParsed, 'day')) {
        nextEndDate = currentDateParsed;
        row.getCell(4).value = nextEndDate.format('YYYY/MM/DD');
        console.log(`超过 currentDate，使用当前 currentDate ${
            nextEndDate.format('YYYY/MM/DD')}`);
        periodsEndRow = rowIndex;
        isBeforeCurrent = false;
        break;
      } else {
        row.getCell(4).value = nextEndDate.format('YYYY/MM/DD');
        periodsEndRow = rowIndex;
      }
      row.commit();
    }

    console.log(` periodsStartRow: ${periodsStartRow}, periodsEndRow: ${
        periodsEndRow}`);
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

    console.log(' Periods:');
    periods.forEach(p => {
      console.log(`Row ${p.row}: ${p.period}, Start: ${
          p.start.format('YYYY-MM-DD')}, End: ${p.end.format('YYYY-MM-DD')}`);
    });

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
      const nextPeriod = periods[i + 1];

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

            if (currentPayment.type === '逾期利息') {
              overdueInterestRowNumbers.push(
                  {date: currentPayment.date, value: currentPayment.value});
              paymentIndex++;
            } else if (
                repaymentType === 1 &&
                currentPayment.type === '本金') {  // 先息后本按期还款中还本金
              if (intimeTerm > 0) {
                const row = worksheet.getRow(startRow + 1);
                row.getCell(9).value =
                    (row.getCell(9).value || 0) + currentPayment.value;
                paymentIndex++;
              } else {
                throw new Error(`插入前置子期但按期还款期数为0`);
              }
            } else {
              const newPeriod = {
                period: prevSubPeriodName,
                start: lastEndDate,
                end: paymentDate,
              };

              newPeriods.push(newPeriod);

              const lastNewPeriod = newPeriods[newPeriods.length - 1];
              const rowNumber = startRow + intimeTermOffset + newPeriods.length;

              if (currentPayment.type === '利息') {
                lastNewPeriod.interest = currentPayment.value;
                interestRowNumbers.push(rowNumber);
              } else if (currentPayment.type === '本金') {
                lastNewPeriod.principal = currentPayment.value;
                principalRowNumbers.push(rowNumber);
              } else if (currentPayment.type === '逾期利息') {
                // lastNewPeriod.overdueInterest = currentPayment.value;
                // overdueInterestRowNumbers.push(rowNumber);
              } else {
                throw new Error(`未知的 currentPayment.type: ${
                    currentPayment.type}，无法处理！`);
              }

              insertedRowNumbers.push(rowNumber);

              subIndex++;
              paymentIndex++;  // 处理下一个 payment
            }

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
        start: newPeriods.length > 0 ? newPeriods[newPeriods.length - 1].end :
                                       currentPeriod.start,
        end: currentPeriod.end,
      });

      originRowNumbers.push(startRow + intimeTermOffset + newPeriods.length);

      subIndex = 1;  // 子期编号，例如 (1)、(2)、(3) ...

      // 如果还有 paymentPairs 没处理完
      while (paymentIndex < paymentPairs.length) {
        const currentPayment = paymentPairs[paymentIndex];
        const paymentDate = dayjs(currentPayment.date);

        // 处理还款日期和截止日相同的情况
        if (paymentDate.isSame(newPeriods[newPeriods.length - 1].end, 'day')) {
          const lastNewPeriod = newPeriods[newPeriods.length - 1];

          if (!lastNewPeriod) {
            throw new Error(
                `当前 newPeriods 为空，无法给最后一个元素赋值！period: ${
                    currentPeriod.period}`);
          }

          if (currentPayment.type === '利息') {
            lastNewPeriod.interest = currentPayment.value;
            interestRowNumbers.push(
                startRow + intimeTermOffset + newPeriods.length);
          } else if (currentPayment.type === '本金') {
            lastNewPeriod.principal = currentPayment.value;
            principalRowNumbers.push(
                startRow + intimeTermOffset + newPeriods.length);
          } else if (currentPayment.type === '逾期利息') {
            lastNewPeriod.overdueInterest = currentPayment.value;
            // overdueInterestRowNumbers.push(startRow + intimeTermOffset +
            // newPeriods.length);
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

            if (currentPayment.type === '逾期利息') {
              overdueInterestRowNumbers.push(
                  {date: currentPayment.date, value: currentPayment.value});
              paymentIndex++;
            } else {
              const newPeriod = {
                period: subPeriodName,
                start: lastEndDate,
                end: paymentDate,
              };
              newPeriods.push(newPeriod);

              const lastNewPeriod = newPeriods[newPeriods.length - 1];
              const rowNumber = startRow + intimeTermOffset + newPeriods.length;

              if (currentPayment.type === '利息') {
                lastNewPeriod.interest = currentPayment.value;
                interestRowNumbers.push(rowNumber);
              } else if (currentPayment.type === '本金') {
                lastNewPeriod.principal = currentPayment.value;
                principalRowNumbers.push(rowNumber);
              } else if (currentPayment.type === '逾期利息') {
                // lastNewPeriod.overdueInterest = currentPayment.value;
                // overdueInterestRowNumbers.push(rowNumber);
              } else {
                throw new Error(`未知的 currentPayment.type: ${
                    currentPayment.type}，无法处理！`);
              }

              insertedRowNumbers.push(rowNumber);

              subIndex++;
              paymentIndex++;  // 移动到下一个 paymentPair
            }
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

          if (currentPayment.type === '逾期利息') {
            overdueInterestRowNumbers.push(
                {date: currentPayment.date, value: currentPayment.value});
            paymentIndex++;
          } else {
            const newPeriod = {
              period: subPeriodName,
              start: lastEndDate,
              end: paymentDate,
            };
            newPeriods.push(newPeriod);

            const lastNewPeriod = newPeriods[newPeriods.length - 1];
            const rowNumber = startRow + intimeTermOffset + newPeriods.length;

            if (currentPayment.type === '利息') {
              lastNewPeriod.interest = currentPayment.value;
              interestRowNumbers.push(rowNumber);
            } else if (currentPayment.type === '本金') {
              lastNewPeriod.principal = currentPayment.value;
              principalRowNumbers.push(rowNumber);
            } else if (currentPayment.type === '逾期利息') {
              // lastNewPeriod.overdueInterest = currentPayment.value;
              // overdueInterestRowNumbers.push(rowNumber);
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
    }

    // 重新写回 worksheet
    let currentRowIdx = periodsStartRow;
    for (let i = 0; i < newPeriods.length; i++) {
      const newPeriod = newPeriods[i];
      const row = worksheet.getRow(currentRowIdx);

      row.getCell(1).value = newPeriod.period;

      if (intimeTerm === 0 && i === 0) {
        // 第一个 newPeriod 且 intimeTerm 为 0，用 start 时间
        row.getCell(3).value = newPeriod.start.format('YYYY/MM/DD');
      } else {
        if (insertedRowNumbers.includes(currentRowIdx)) {
          row.getCell(3).value = {formula: `D${currentRowIdx - 1}`};
        } else {
          let prevRowIdx = currentRowIdx - 1;
          while (insertedRowNumbers.includes(prevRowIdx) && prevRowIdx > 0) {
            prevRowIdx--;
          }
          row.getCell(3).value = {formula: `D${prevRowIdx}`};
        }
      }
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

    // 最后一笔还款后
    if (isBeforeCurrent) {
      const currentRow = worksheet.getRow(currentRowIdx);

      // 设置行头与日期
      currentRow.getCell(1).value = '最后一笔还款后';

      // 找到上一行中不属于 insertedRowNumbers 的行号
      let prevRowIdx = currentRowIdx - 1;
      while (insertedRowNumbers.includes(prevRowIdx) && prevRowIdx > 0) {
        prevRowIdx--;
      }
      // 设置公式引用该行的 D 列
      currentRow.getCell(3).value = {formula: `D${prevRowIdx}`};
      currentRow.getCell(4).value = currentDateParsed.format('YYYY/MM/DD');

      console.log(`写入 Row ${currentRowIdx}：最后一笔还款后，${
          endDateParsed.format(
              'YYYY/MM/DD')} ~ ${currentDateParsed.format('YYYY/MM/DD')}`);

      // 设置公式
      const r = currentRowIdx;
      const rPrev = r - 1;

      currentRow.getCell(2).value = {
        formula: `B${rPrev} - F${rPrev}`
      };  // B = B-1 - F-1

      currentRow.getCell(5).value = {formula: `D${r} - C${r}`};  // E = D - C

      currentRow.getCell(6).value = {formula: `B${r}`};   // F = B
      currentRow.getCell(20).value = {formula: `K${r}`};  // T = K

      currentRow.getCell(8).value = {
        formula: `B${r} * $E$6 / 360 * E${r}`
      };  // H = B * E6 / 360 * E
      currentRow.getCell(11).value = {
        formula: `K${rPrev} + F${r} - I${r}`
      };  // K = K-1 + F - I
      currentRow.getCell(12).value = {
        formula: `L${rPrev} + H${r} - J${r}`
      };  // L = L-1 + H - J
      currentRow.getCell(13).value = {
        formula: `N${r} * O${r} / 360 * R${r}`
      };  // M = N * O / 360 * R
      currentRow.getCell(14).value = {formula: `L${rPrev}`};      // N = L-1
      currentRow.getCell(15).value = {formula: `$E$7`};           // O = E7
      currentRow.getCell(16).value = {formula: `D${rPrev}`};      // P = D-1
      currentRow.getCell(17).value = {formula: `D${r}`};          // Q = D
      currentRow.getCell(18).value = {formula: `Q${r} - P${r}`};  // R = Q - P
      currentRow.getCell(19).value = {
        formula: `T${r} * U${r} / 360 * X${r}`
      };  // S = T * U / 360 * X
      currentRow.getCell(21).value = {formula: `$E$7`};           // U = E7
      currentRow.getCell(22).value = {formula: `P${r}`};          // V = P
      currentRow.getCell(23).value = {formula: `Q${r}`};          // W = Q
      currentRow.getCell(24).value = {formula: `W${r} - V${r}`};  // X = W - V
      currentRow.getCell(25).value = {formula: `S${r} + M${r}`};  // Y = S + M
      currentRow.getCell(27).value = {formula: `Y${r} - Z${r}`};  // AA = Y - Z

      currentRow.commit();
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
      // B列写公式 =E5
      firstRow.getCell(2).value = {formula: `$E$5`};

      firstRow.commit();
    }

    // 处理第二行到倒数第二行
    for (let rowIdx = startRow + 1; rowIdx <= lastRowIdx; rowIdx++) {
      const row = worksheet.getRow(rowIdx);

      // 设置 E 列（第 5 列）为公式：=（D列 - C列的天数差）
      row.getCell(5).value = {formula: `D${rowIdx}-C${rowIdx}`};

      // 在B列写入公式 B(x) = B(x-1) - F(x-1)
      row.getCell(2).value = {formula: `B${rowIdx - 1}-F${rowIdx - 1}`};

      // 设置 K 列：K(x) = K(x-1) + F(x) - I(x)
      row.getCell(11).value = {formula: `K${rowIdx - 1}+F${rowIdx}-I${rowIdx}`};

      // 给L列加公式：L(x) = L(x-1) + H(x) - J(x)
      row.getCell(12).value = {formula: `L${rowIdx - 1}+H${rowIdx}-J${rowIdx}`};

      // M列加公式：M(x) = N(x) * O(x) / 360 * R(x)
      row.getCell(13).value = {formula: `N${rowIdx}*O${rowIdx}/360*R${rowIdx}`};

      const prevRow = rowIdx - 1;

      // 设置第 14 列（N列）的公式
      if (lastPeriodRowNumbers.includes(rowIdx)) {
        row.getCell(14).value = {
          formula: `L${prevRow}`  // 公式：N(x) = L(x-1)
        };
      } else if (
          insertedRowNumbers.includes(prevRow) &&
          !originRowNumbers.includes(prevRow)) {
        row.getCell(14).value = {
          formula: `N${prevRow}-J${prevRow}`  // 公式：N(x) = N(x-1) - J(x-1)
        };
      } else {
        row.getCell(14).value = {
          formula: `H${prevRow}-J${prevRow}`  // 公式：N(x) = H(x-1) - J(x-1)
        };
      }
      // P列（第16列）
      if (rowIdx === startRow + 1) {
        row.getCell(16).value = {
          formula: `C${rowIdx}`  // P(x) = C(x)
        };
      } else {
        row.getCell(16).value = {
          formula: `D${rowIdx - 1}`  // P(x) = D(x-1)
        };
      }

      // Q列（第17列）
      if (insertedRowNumbers.includes(rowIdx) &&
          !originRowNumbers.includes(rowIdx)) {
        row.getCell(17).value = {
          formula: `D${rowIdx}`  // Q(x) = D(x)
        };
      } else {
        if (isBeforeCurrent) {
          row.getCell(17).value = {
            formula: `$E$10`  // Q(x) = E10
          };
        } else {
          row.getCell(17).value = currentDateParsed.format('YYYY/MM/DD');
        }
      }

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

    if (repaymentType === 1) {  // 先息后本

      // 单独处理第一行
      if (intimeTerm > 0) {
        const firstRow = worksheet.getRow(firstRowIdx + 1);
        firstRow.getCell(10).value = {
          formula: `H${firstRowIdx + 1}`
        };  // 已还利息等于应还利息
        firstRow.commit();
      }

      // 处理第一行到倒数第二行
      for (let rowIdx = startRow + 1; rowIdx <= lastRowIdx; rowIdx++) {
        const row = worksheet.getRow(rowIdx);

        // 如果这一行是 originRowNumbers 中的
        if (originRowNumbers.includes(rowIdx)) {
          const currentIndex = originRowNumbers.indexOf(rowIdx);
          let formula;

          // 情况1：当前是第一个元素，无前驱行
          if (currentIndex === 0) {
            formula = `B${rowIdx}*$E$6/360*E${rowIdx}`;

            // 情况2：有前驱行
          } else {
            const prevRow = originRowNumbers[currentIndex - 1];
            const gap = rowIdx - prevRow;

            // 子情况1：行号连续
            if (gap === 1) {
              formula = `B${rowIdx}*$E$6/360*E${rowIdx}`;

              // 子情况2：行号不连续
            } else {
              // 生成中间行号数组（例如 18→22 时的 19,20,21）
              const middleRows = Array.from(
                  {length: rowIdx - prevRow - 1}, (_, i) => prevRow + 1 + i);

              // 构造中间行的累加部分（Bx*E6*Ex）
              const sumFormula =
                  middleRows.map(x => `B${x}*$E$6/360*E${x}`).join('+');

              // 构造当前行的减数部分（E22 - E19 - E20 - E21）
              const eSubtraction = middleRows.map(x => `E${x}`).join('-');

              // 合并完整公式
              formula = `${sumFormula ? sumFormula + '+' : ''}B${
                  rowIdx}*$E$6/360*(E${rowIdx}-${eSubtraction})`;
            }
          }

          // 写入单元格
          row.getCell(8).value = {formula};
        }
        row.getCell(6).value = {formula: `I${rowIdx}`};  // F列 = I列
        if (lastPeriodRowNumbers.includes(rowIdx)) {
          row.getCell(6).value = 0;  // F列 = 0
        }

        const targetRow = lastPeriodRowNumbers.length >
                0 ?  // 结息日之后的第一行，或者最后一笔还款后前一行
            lastPeriodRowNumbers[0] :
            currentRowIdx - 1;

        if (rowIdx === targetRow) {
          row.getCell(6).value = {formula: `B${rowIdx}`};
        }

        if (lastPeriodRowNumbers.includes(rowIdx)) {
          // 当到期后，当期未还本金会变成累计未还本金
          if (rowIdx === targetRow) {
            // t = f
            row.getCell(20).value = {formula: `F${rowIdx}`};
          } else {
            // t = k-1
            row.getCell(20).value = {formula: `K${rowIdx - 1}`};
          }
          // U = $E$7
          row.getCell(21).value = {formula: `$E$7`};
          // S = T * U / 360 * X
          row.getCell(19).value = {
            formula: `T${rowIdx} * U${rowIdx} / 360 * X${rowIdx}`
          };
          // V = P(x)
          row.getCell(22).value = {formula: `P${rowIdx}`};
          // W = Q(x) ）
          row.getCell(23).value = {formula: `Q${rowIdx}`};
          // X = W - V
          row.getCell(24).value = {formula: `W${rowIdx} - V${rowIdx}`};
        }

        row.commit();
      }
    } else {                     // 等额本金和等额本息
      let originRowCounter = 0;  // 记录当前是 origin 中的第几个
      for (let rowIdx = startRow + 1; rowIdx <= lastRowIdx; rowIdx++) {
        const row = worksheet.getRow(rowIdx);

        if (originRowCounter < realIntimeTerm) {
          // 前 realIntimeTerm 个 origin 行：I列 ：i(x) = f(x)
          row.getCell(9).value = {formula: `F${rowIdx}`};
          // J列 = H{rowIdx} → j(x) = h(x)
          row.getCell(10).value = {formula: `H${rowIdx}`};
        }
        // 如果这一行是 originRowNumbers 中的
        if (originRowNumbers.includes(rowIdx)) {
          originRowCounter++;
          if (repaymentType === 2) {  // 等额本金
            // F列 = ROUND(E5/E8, 2)
            row.getCell(6).value = {
              formula: `ROUND($E$5/$E$8, 2)`
            };
            // H列 = ROUND(B行号 * $E$6 / 360 * E行号, 2)
            row.getCell(8).value = {
              formula: `ROUND(B${rowIdx}*$E$6/360*E${rowIdx}, 2)`
            };
          } else if (repaymentType === 3) {  // 等额本息
            // F列 = ROUND(PPMT(...), 2)
            row.getCell(6).value = {
              formula: `ROUND(PPMT($E$6/12, A${rowIdx}, $E$8, -$E$5, 0), 2)`
            };
            // H列 = ROUND(IPMT(...), 2)
            row.getCell(8).value = {
              formula: `ROUND(IPMT($E$6/12, A${rowIdx}, $E$8, -$E$5, 0), 2)`
            };
          }
        }

        if (lastPeriodRowNumbers.includes(rowIdx)) {
          // 当到期后，当期未还本金会变成累计未还本金
          row.getCell(20).value = {formula: `K${rowIdx - 1}`};
        } else if (
            insertedRowNumbers.includes(rowIdx - 1) &&
            !originRowNumbers.includes(rowIdx - 1)) {
          // 上一行还了本金且上一行不是原始行的时候，T = T(x-1) - I(x-1)
          row.getCell(20).value = {formula: `T${rowIdx - 1}-I${rowIdx - 1}`};
        } else {
          // T = F(x-1) - I(x-1)
          row.getCell(20).value = {formula: `F${rowIdx - 1} - I${rowIdx - 1}`};
        }

        // U = $E$7
        row.getCell(21).value = {formula: `$E$7`};
        // S = T * U / 360 * X
        row.getCell(19).value = {
          formula: `T${rowIdx} * U${rowIdx} / 360 * X${rowIdx}`
        };
        // V = P(x)
        row.getCell(22).value = {formula: `P${rowIdx}`};
        // W = Q(x) ）
        row.getCell(23).value = {formula: `Q${rowIdx}`};
        // X = W - V
        row.getCell(24).value = {formula: `W${rowIdx} - V${rowIdx}`};
      }
    }

    // 插入多出来的逾期利息，遍历每一个逾期记录，插入到正确位置
    for (const item of overdueInterestRowNumbers) {
      const insertDate = dayjs(item.date);

      let insertRowIndex = currentRowIdx;
      let foundExactMatch = false;

      for (let i = startRow + 1; i <= currentRowIdx - 1; i++) {
        const row = worksheet.getRow(i);
        const cellDate = row.getCell(4).value;

        if (cellDate) {
          const cellDateDayjs = dayjs(cellDate);

          if (cellDateDayjs.isSame(insertDate, 'day')) {
            row.getCell(26).value = item.value;
            row.commit();
            foundExactMatch = true;
            break;
          }

          if (cellDateDayjs.isAfter(insertDate)) {
            insertRowIndex = i;
            break;
          }
        }
      }

      if (!foundExactMatch) {
        // 插入一行
        worksheet.spliceRows(insertRowIndex, 0, []);

        // 插入后修复公式
        for (let rowIdx = startRow + 1; rowIdx <= currentRowIdx + 1; rowIdx++) {
          const row = worksheet.getRow(rowIdx);

          for (let colIdx = 1; colIdx <= row.cellCount; colIdx++) {
            const cell = row.getCell(colIdx);

            let formulaText = null;

            // ExcelJS 可能存在 formula 或 value.formula 两种写法
            if (cell.formula) {
              formulaText = cell.formula;
            } else if (
                cell.value && typeof cell.value === 'object' &&
                cell.value.formula) {
              formulaText = cell.value.formula;
            }

            if (formulaText) {
              const updatedFormula =
                  formulaText.replace(/\$?[A-Z]+\$?(\d+)/g, (match) => {
                    const col = match.match(/[A-Z]+/)[0];
                    const rowNumber = parseInt(match.match(/\d+/)[0], 10);

                    if (rowNumber >= insertRowIndex) {
                      const newRowNum = rowNumber + 1;
                      return match.replace(
                          rowNumber.toString(), newRowNum.toString());
                    }
                    return match;
                  });

              // 写入新公式
              cell.value = {formula: updatedFormula};
            }
          }
          row.commit();
        }

        const newRow = worksheet.getRow(insertRowIndex);

        const prevLabel = worksheet.getRow(insertRowIndex - 1).getCell(1).value;
        const nextLabel = getNextOverdueLabel(prevLabel);
        newRow.getCell(1).value = nextLabel;

        newRow.getCell(3).value = item.date;
        newRow.getCell(4).value = item.date;
        newRow.getCell(26).value = item.value;
        newRow.getCell(27).value = {
          formula: `Y${insertRowIndex}-Z${insertRowIndex}`
        };
        newRow.commit();

        currentRowIdx++;  // 扩展数据区
      }
    }

    // 给到期后的还款项改名
    const matchedRows = [];

    for (let rowIdx = startRow; rowIdx <= currentRowIdx; rowIdx++) {
      const row = worksheet.getRow(rowIdx);
      const cellValue = row.getCell(1).value;  // A列 = 第1列

      if (typeof cellValue === 'string') {
        const regex = new RegExp(`^${term}\\(\\d+\\)$`);

        if (regex.test(cellValue.trim())) {
          matchedRows.push(rowIdx);
        }
      }
    }

    // 替换为 到期后还款1、2、3...
    matchedRows.forEach((rowIdx, index) => {
      const row = worksheet.getRow(rowIdx);
      row.getCell(1).value = `到期后还款${index + 1}`;
      row.commit();  // 提交修改
    });
    // sum行
    const sumRowIndex = currentRowIdx + 10;
    const sumRow = worksheet.getRow(sumRowIndex);

    sumRow.getCell(1).value = '合计';

    let sumIdx = currentRowIdx;
    if (!isBeforeCurrent) {
      sumIdx = currentRowIdx - 1;  // 没有最后一行最后一笔还款后
      const lastRow = worksheet.getRow(sumIdx);
      lastRow.getCell(2).value = {
        // B = B-1 - F-1
        formula: `B${sumIdx - 1}-F${sumIdx - 1}`
      };
      lastRow.getCell(6).value = {formula: `B${sumIdx}`};  // F = B
      lastRow.commit();
    }

    // 需要求和的列
    const sumCols = [
      {letter: 'H', index: 8}, {letter: 'I', index: 9},
      {letter: 'J', index: 10}, {letter: 'M', index: 13},
      {letter: 'S', index: 19}, {letter: 'Y', index: 25},
      {letter: 'Z', index: 26}, {letter: 'AA', index: 27}
    ];

    // 从startRow到currentRowIdx求和
    for (const {letter, index} of sumCols) {
      let formula;

      if (letter === 'M') {
        formula = `IF(SUM(${letter}${startRow}:${letter}${sumIdx}) >= Z${
            sumRowIndex}, SUM(${letter}${startRow}:${letter}${sumIdx}) - Z${
            sumRowIndex}, 0)`;
      } else if (letter === 'S') {
        formula = `IF(SUM(M${startRow}:M${sumIdx}) >= Z${sumRowIndex}, SUM(${
            letter}${startRow}:${letter}${sumIdx}), SUM(${letter}${startRow}:${
            letter}${sumIdx}) - Z${sumRowIndex} + SUM(M${startRow}:M${
            sumIdx}))`;
      } else {
        formula = `SUM(${letter}${startRow}:${letter}${sumIdx})`;
      }

      sumRow.getCell(index).value = {formula};
    }

    // K L的值从原来最后一行拿
    sumRow.getCell(11).value = {formula: `K${sumIdx}`};
    sumRow.getCell(12).value = {formula: `L${sumIdx}`};

    sumRow.commit();
    sumRowNumbers.push(sumRowIndex);
    // 最后表格
    const displayStartRow = sumRowIndex + 1;

    const finalRows = [
      {label: '拖欠本金', formula: `K${sumRowIndex}`},
      {label: '拖欠正常利息', formula: `L${sumRowIndex}`},
      {label: '复利', formula: `M${sumRowIndex}`},
      {label: '罚息', formula: `S${sumRowIndex}`}, {
        label: '合计',
        formula: `K${sumRowIndex} + L${sumRowIndex} + M${sumRowIndex} + S${
            sumRowIndex}`
      },
      {label: '本息合计', formula: `K${sumRowIndex} + L${sumRowIndex}`},
      {label: '未还逾期利息', formula: `AA${sumRowIndex}`}
    ];

    finalRows.forEach((item, idx) => {
      const rowIdx = displayStartRow + idx;
      const row = worksheet.getRow(rowIdx);

      // 合并 A-J（1-10），K-AA（11-27）
      worksheet.mergeCells(rowIdx, 1, rowIdx, 10);   // A-J
      worksheet.mergeCells(rowIdx, 11, rowIdx, 27);  // K-AA

      // 写入内容和公式
      row.getCell(1).value = item.label;
      row.getCell(11).value = {formula: item.formula};

      // 设置样式
      row.getCell(1).font = {bold: true};
      row.getCell(11).font = {bold: true};

      row.getCell(1).alignment = {vertical: 'middle', horizontal: 'center'};

      row.getCell(11).alignment = {vertical: 'middle', horizontal: 'left'};

      row.commit();
    });

    worksheet.eachRow((row, rowNumber) => {
      row.eachCell((cell, colNumber) => {
        // 设置通用对齐样式
        cell.alignment = {
          vertical: cell.alignment?.vertical || 'middle',
          horizontal: cell.alignment?.horizontal || 'center',
          wrapText: true
        };

        // cell.border = {
        //   top: {style: 'thick'},
        //   left: {style: 'thick'},
        //   bottom: {style: 'thick'},
        //   right: {style: 'thick'}
        // };

        // 设置百分比格式
        if ((rowNumber === 6 || rowNumber === 7) && colNumber === 5) {
          cell.numFmt = '0.######%';
        }

        if (colNumber === 15 || colNumber === 21) {
          cell.numFmt = '0.######%';
        }

        // 两位小数+千分号（行5，列5）
        if (rowNumber === 5 && colNumber === 5) {
          cell.numFmt = '#,##0.00';
        }

        // 两位小数+千分号（统一用 '#,##0.00'）：B, F, H, I, J, K, L, M, N, S,
        // T, X, Y, Z, AA
        const decimalColumns =
            [2, 6, 8, 9, 10, 11, 12, 13, 14, 19, 20, 25, 26, 27];
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

    // 设置边框
    const realLastRowIdx = displayStartRow + finalRows.length - 1;
    // 设置边框样式：medium + 黑色
    const borderStyle = {
      top: {style: 'medium', color: {argb: 'FF000000'}},
      bottom: {style: 'medium', color: {argb: 'FF000000'}},
      left: {style: 'medium', color: {argb: 'FF000000'}},
      right: {style: 'medium', color: {argb: 'FF000000'}}
    };

    // 1. 设置 A1:E13 的边框
    for (let rowIdx = 1; rowIdx <= 13; rowIdx++) {
      const row = worksheet.getRow(rowIdx);
      for (let colIdx = 1; colIdx <= 5; colIdx++) {  // A-E = 1-5
        const cell = row.getCell(colIdx);
        cell.border = borderStyle;
      }
    }

    // 2. 设置 A15:AA{realLastRowIdx} 的边框
    for (let rowIdx = 15; rowIdx <= realLastRowIdx; rowIdx++) {
      const row = worksheet.getRow(rowIdx);
      for (let colIdx = 1; colIdx <= 27; colIdx++) {  // A-AA = 1-27
        const cell = row.getCell(colIdx);
        cell.border = borderStyle;
      }
    }
  }

  // 总表

  const n = groups.length;
  let startIdx = 3;
  for (let i = 1; i <= n; i++) {
    const rowIdx = i + startIdx - 1;
    const row = summarySheet.getRow(rowIdx);
    const sumIdx = sumRowNumbers[i - 1];

    const sheetName = groups[i - 1].sheetName;

    // 如果 sheetName 含有空格或中文，需加英文单引号
    const quotedSheetName =
        /[\s\u4e00-\u9fa5]/.test(sheetName) ? `'${sheetName}'` : sheetName;

    // A列：序号
    row.getCell(1).value = i;

    // B列：引用子表 E9
    row.getCell(2).value = {formula: `${quotedSheetName}!$E$9`};

    // C列：引用子表 E10
    row.getCell(3).value = {formula: `${quotedSheetName}!$E$10`};

    // D列：比较 C列 和 rawDate
    const cellC = `C${rowIdx}`;
    row.getCell(4).value = {
      formula: `IF(${cellC}>DATEVALUE("${rawDateStr}"), "加速到期", "正常到期")`
    };

    // E列：子表 E8
    row.getCell(5).value = {formula: `${quotedSheetName}!$E$8`};
    // F列：子表 E11
    row.getCell(6).value = {formula: `${quotedSheetName}!$E$11`};
    // G列：子表 E5
    row.getCell(7).value = {formula: `${quotedSheetName}!$E$5`};
    // H列：子表 E6
    row.getCell(8).value = {formula: `${quotedSheetName}!$E$6`};
    // I列：子表 E7
    row.getCell(9).value = {formula: `${quotedSheetName}!$E$7`};
    // J列（第10列）：子表 I列（第9列）第 sumIdx 行
    row.getCell(10).value = {formula: `${quotedSheetName}!I${sumIdx}`};
    // K列（第11列）：子表 H列（第8列）第 sumIdx 行
    row.getCell(11).value = {formula: `${quotedSheetName}!H${sumIdx}`};
    // L列（第12列）：子表 J列（第10列）第 sumIdx 行
    row.getCell(12).value = {formula: `${quotedSheetName}!J${sumIdx}`};
    // M列（第13列）：子表 Z列（第26列）第 sumIdx 行
    row.getCell(13).value = {formula: `${quotedSheetName}!Z${sumIdx}`};
    // N列（第14列）：子表 K列（第11列）第 sumIdx 行
    row.getCell(14).value = {formula: `${quotedSheetName}!K${sumIdx}`};
    // O列（第15列）：子表 L列（第12列）第 sumIdx 行
    row.getCell(15).value = {formula: `${quotedSheetName}!L${sumIdx}`};
    // P列（第16列）：子表 M列（第13列）第 sumIdx 行
    row.getCell(16).value = {formula: `${quotedSheetName}!M${sumIdx}`};
    // Q列（第17列）：子表 S列（第19列）第 sumIdx 行
    row.getCell(17).value = {formula: `${quotedSheetName}!S${sumIdx}`};
    // R列（第18列）：子表 AA列（第27列）第 sumIdx 行
    row.getCell(18).value = {formula: `${quotedSheetName}!AA${sumIdx}`};

    row.commit();
  }

  // 合计行索引
  const ssumIdx = startIdx + n + 5;

  // 获取合计行
  const sumRow = summarySheet.getRow(ssumIdx);

  // A列：写“合计”
  sumRow.getCell(1).value = '合计';

  // 要求和的列号
  const sumCols = [7, 10, 11, 12, 13, 14, 15, 16, 17, 18];

  // 所有列中，用到的最大列号（用于加粗整行）
  const allCols = [1, ...sumCols];

  for (const col of allCols) {
    const cell = sumRow.getCell(col);

    // 如果是求和列，设置公式
    if (col !== 1) {
      const colLetter = getExcelColLetter(col);
      cell.value = {
        formula: `SUM(${colLetter}${startIdx}:${colLetter}${startIdx + n - 1})`
      };
    }

    // 设置加粗字体
    cell.font = {bold: true};
  }

  // 总表特殊单元格
  sumRow.getCell(19).value = {
    formula: `N${ssumIdx} + O${ssumIdx}+ R${ssumIdx}`
  };
  sumRow.getCell(19).font = {bold: true};

  // 提交合计行
  sumRow.commit();

  // 第一步：构建分类 Map（保持不变）
  const rateGroups = new Map();

  for (let i = 0; i < n; i++) {
    const rowIdx = startIdx + i;
    const row = summarySheet.getRow(rowIdx);
    if (!row) {
      console.warn(`跳过不存在行: ${rowIdx}`);
      continue;
    }

    // 处理I列单元格
    const cellI = row.getCell(9);
    let actualValue;

    if (cellI.value && typeof cellI.value === 'object' && cellI.value.formula) {
      // 解析公式
      const formula = cellI.value.formula;
      const formulaMatch = formula.match(/^'?(.*?)'?!\$?([A-Z]+\$?\d+)$/);

      if (formulaMatch) {
        // 处理带转义的工作表名
        const sheetName = formulaMatch[1].replace(/''/g, '\'');
        const cellAddress = formulaMatch[2].replace(/\$/g, '');

        // 获取目标工作表
        const targetSheet = workbook.getWorksheet(sheetName);
        if (targetSheet) {
          const targetCell = targetSheet.getCell(cellAddress);
          actualValue = targetCell.value;
        } else {
          console.error(`工作表不存在: ${sheetName} (行 ${rowIdx})`);
          actualValue = '无效引用';
        }
      } else {
        console.error(`公式格式错误: ${formula} (行 ${rowIdx})`);
        actualValue = '格式错误';
      }
    } else {
      // 直接处理非公式值
      actualValue = cellI.value;
    }

    // 处理空值
    if (actualValue === null || actualValue === undefined) {
      actualValue = '空值';
    }

    // 更新分组
    const groupKey =
        typeof actualValue === 'number' ? actualValue : String(actualValue);

    if (!rateGroups.has(groupKey)) {
      rateGroups.set(groupKey, []);
    }
    rateGroups.get(groupKey).push(rowIdx);

    // 调试日志
    console.log(`行 ${rowIdx} 归类到:`, groupKey);
  }

  // 打印最终分组结果
  console.log('\n最终分组结果:');
  for (const [key, rows] of rateGroups.entries()) {
    console.log(`[${key}] => 行号: ${rows.join(', ')}`);
  }

  // 第二步：生成公式并写入
  let outputIdx = ssumIdx + 2;

  const summaryValues = [];  // 额外记录 sum 和 rate 的值，用于信息汇总表

  for (const [rate, rowIndices] of rateGroups.entries()) {
    // --- 错误处理：跳过空分组 ---
    if (rowIndices.length === 0) {
      continue;  // ⚡ 跳过无行的分组
    }

    // --- 性能优化：使用 SUM 函数 ---
    const formulaParts = [];
    for (const idx of rowIndices) {
      formulaParts.push(`N${idx}`, `O${idx}`);  // 收集所有 N/O 列单元格
    }
    const sumFormula = `SUM(${formulaParts.join(',')})`;  // 如 SUM(N1,O1,N3,O3)

    // --- 获取利率单元格引用（取第一个行号） ---
    const firstRowIdx = rowIndices[0];
    const rateCell = `I${firstRowIdx}`;  // 如 I1

    // 构建 Excel 公式（注意 TEXT 保留两位小数，百分比）
    const formulaText = `"以未还本息" & TEXT(${
        sumFormula}, "0.00") & "元为基数，按照年利率" & TEXT(${
        rateCell}, "0.######%") & "计算"`;

    // --- 写入公式到 Excel ---
    const outputRow = summarySheet.getRow(outputIdx++);
    // 合并 B → R（第2列 → 第18列）
    summarySheet.mergeCells(outputRow.number, 2, outputRow.number, 18);
    outputRow.getCell(2).value = {
      formula: `=${formulaText}`  // 必须以等号开头
    };
    outputRow.commit();

    summaryValues.push(
        {sumFormula, rate});  // 额外记录 sum 和 rate 的值，用于信息汇总表
  }


  // 设置总表格式
  summarySheet.eachRow((row, rowNumber) => {
    row.eachCell((cell, colNumber) => {
      // 设置通用对齐样式（居中 + 自动换行）
      cell.alignment = {
        vertical: cell.alignment?.vertical || 'middle',
        horizontal: cell.alignment?.horizontal || 'center',
        wrapText: true
      };
      // cell.border = {
      //   top: {style: 'thick'},
      //   left: {style: 'thick'},
      //   bottom: {style: 'thick'},
      //   right: {style: 'thick'}
      // };

      // 设置百分比格式：H、I（列号 8、9）
      const percentColumns = [8, 9];
      if (percentColumns.includes(colNumber)) {
        cell.numFmt = '0.######%';
      }

      // 设置千分位 + 两位小数：G、J、K、L、M、N、O、P、Q、R
      const decimalColumns = [7, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19];
      if (decimalColumns.includes(colNumber)) {
        cell.numFmt = '#,##0.00';
      }

      // 设置日期格式：B、C（列号 2、3）
      const dateColumns = [2, 3];
      if (dateColumns.includes(colNumber)) {
        cell.numFmt = 'yyyy/mm/dd';
      }
    });
  });

  // 加边框
  // 设置边框样式：medium + 黑色
  const borderStyle = {
    top: {style: 'medium', color: {argb: 'FF000000'}},
    bottom: {style: 'medium', color: {argb: 'FF000000'}},
    left: {style: 'medium', color: {argb: 'FF000000'}},
    right: {style: 'medium', color: {argb: 'FF000000'}}
  };
  // 1. summarySheet：第 1 行 到 ssumIdx 行，列 A-R（1-18）
  for (let rowIdx = 1; rowIdx <= ssumIdx; rowIdx++) {
    const row = summarySheet.getRow(rowIdx);
    for (let colIdx = 1; colIdx <= 18; colIdx++) {
      const cell = row.getCell(colIdx);
      cell.border = borderStyle;
    }
  }

  // 2. summarySheet：从 ssumIdx + 2 到 outputIdx - 1 行，列 B-R（2-18）
  for (let rowIdx = ssumIdx + 2; rowIdx < outputIdx; rowIdx++) {
    const row = summarySheet.getRow(rowIdx);
    for (let colIdx = 2; colIdx <= 18; colIdx++) {
      const cell = row.getCell(colIdx);
      cell.border = borderStyle;
    }
  }

  // 设置背景色：橙黄色（N~R列 = 14~18）
  const highlightColor = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: {argb: 'FFF9A825'}  // 橙黄色
  };

  // 行号数组：第1行和第ssumIdx行
  const highlightRows = [2, ssumIdx];

  // 应用于 summarySheet 的 N~R 列（14~18）
  for (const rowIdx of highlightRows) {
    const row = summarySheet.getRow(rowIdx);
    for (let colIdx = 14; colIdx <= 18; colIdx++) {
      const cell = row.getCell(colIdx);
      cell.fill = highlightColor;
    }
  }

  // 设置亮黄色背景色（纯黄）
  const yellowFill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: {argb: 'FFFFFF00'}  // 亮黄
  };

  // 设置 summarySheet 的 B~R 列（2~18）背景色
  for (let rowIdx = ssumIdx + 2; rowIdx < outputIdx; rowIdx++) {
    const row = summarySheet.getRow(rowIdx);
    for (let colIdx = 2; colIdx <= 18; colIdx++) {
      const cell = row.getCell(colIdx);
      cell.fill = yellowFill;
    }
  }

  const infoIdx = outputIdx + 5;
  const headerRow = summarySheet.getRow(infoIdx);
  const row2 = summarySheet.getRow(infoIdx + 1);

  // 解析身份证信息
  function parseIdNumber(idNumber) {
    if (!/^\d{17}[\dXx]$/.test(idNumber)) return {gender: '', birth: ''};
    const birth = idNumber.slice(6, 14);
    const birthFormatted = `${birth.slice(0, 4)}-${birth.slice(4, 6)}`;
    const genderCode = Number(idNumber[16]);
    const gender = genderCode % 2 === 0 ? '女' : '男';
    return {gender, birth: birthFormatted};
  }
  const {gender, birth} = parseIdNumber(infoGroup.idNumber);

  // 构建汇总字段数组
  const summaryFields = [
    {title: '借款人', value: a}, {title: '性别', value: gender},
    {title: '出生年月', value: birth},
    {title: '身份证号', value: infoGroup.idNumber},
    {title: '户籍地', value: infoGroup.registeredAddress},
    {title: '联系地址', value: infoGroup.contactAddress},
    {title: '被告手机号', value: infoGroup.defendantPhoneNumber},
    {title: '暂计日期', value: dayjs(rawDate).format('YYYY/MM/DD')}, {
      title: '开口计算开始日',
      value: dayjs(rawDate).add(1, 'day').format('YYYY/MM/DD')
    },
    {title: '本金', value: {formula: `N${ssumIdx}`}},
    {title: '利息', value: {formula: `O${ssumIdx}`}},
    {title: '逾期利息', value: {formula: `R${ssumIdx}`}},
    {title: '标的额', value: {formula: `S${ssumIdx}`}}
  ];

  const decimalFields =
      ['本金', '利息', '逾期利息', '标的额', '律师费', '贷款总额', '已还本金'];

  summaryFields.forEach((item, index) => {
    const col = index + 2;
    const headerCell = headerRow.getCell(col);
    const dataCell = row2.getCell(col);

    headerCell.value = item.title;
    dataCell.value = item.value;

    // 如果是需要保留两位小数的字段
    if (decimalFields.includes(item.title)) {
      dataCell.numFmt = '0.00';
    }
  });

  // 写入动态列（本息、逾期利息）
  const summaryStartCol = summaryFields.length + 2;
  summaryValues.forEach((item, i) => {
    const baseCol = summaryStartCol + i * 2;
    headerRow.getCell(baseCol).value = `本息${i + 1}`;
    headerRow.getCell(baseCol + 1).value = `逾期利息${i + 1}`;
    const sumCell = row2.getCell(baseCol);
    sumCell.value = {formula: item.sumFormula};
    sumCell.numFmt = '0.00';
    const rateCell = row2.getCell(baseCol + 1);
    rateCell.value = item.rate;
    rateCell.numFmt = '0.######%';
  });

  // 写入附加字段
  const additionalFields = [
    {title: '律师费', value: infoGroup.lawyerFee}, {
      title: '合同签署日',
      value: dayjs(infoGroup.contractSignDate).format('YYYY/MM/DD')
    },
    {title: '合同名称', value: infoGroup.contractName}, {
      title: '首次放款日',
      value: dayjs(infoGroup.firstLoanDate).format('YYYY/MM/DD')
    },
    {title: '贷款总额', value: infoGroup.totalLoanAmount},
    {title: '已还本金', value: {formula: `J${ssumIdx}`}},
    {title: '产品名称', value: infoGroup.productName}
  ];

  const additionalStartCol = summaryStartCol + summaryValues.length * 2;
  additionalFields.forEach((item, i) => {
    const col = additionalStartCol + i;
    const headerCell = headerRow.getCell(col);
    const dataCell = row2.getCell(col);

    headerCell.value = item.title;
    dataCell.value = item.value;

    // 设置数值格式：保留两位小数
    if (decimalFields.includes(item.title)) {
      dataCell.numFmt = '0.00';
    }
  });
  const finalCol = additionalStartCol + additionalFields.length;

  for (let col = 2; col < finalCol; col++) {
    const headerCell = headerRow.getCell(col);
    const dataCell = row2.getCell(col);

    // 表头行样式（背景、加粗、居中、边框）
    headerCell.fill = yellowFill;
    headerCell.font = {bold: true};
    headerCell
        .alignment = {vertical: 'middle', horizontal: 'center', wrapText: true};
    headerCell.border = borderStyle;

    // 数据行样式（居中、边框）
    dataCell
        .alignment = {vertical: 'middle', horizontal: 'center', wrapText: true};
    dataCell.border = borderStyle;
  }

  // 设置列宽
  for (let col = 1; col <= finalCol; col++) {
    summarySheet.getColumn(col).width = 17.5;
  }
  headerRow.commit();
  row2.commit();


  const excelbuffer = await workbook.xlsx.writeBuffer();
  const bufferName = `${a}截止${b}`;
  return {excelbuffer, bufferName};
}

// =====================
// 文件上传并生成 Excel
// =====================
app.post('/generate-excel', upload.single('file'), async (req, res) => {
  // const now = new Date();
  // const deadline = new Date('2025-07-10');

  // if (now > deadline) {
  //   return res.status(403).json(
  //       {message: '超过试用日期，无法继续使用此功能。'});
  // }
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
    const infoGroup = {
      idNumber: sheet.getCell('B4').value || '',
      registeredAddress: sheet.getCell('B5').value || '',
      contactAddress: sheet.getCell('B6').value || '',
      defendantPhoneNumber: sheet.getCell('B7').value || '',
      lawyerFee: sheet.getCell('B8').value || '',
      contractSignDate: sheet.getCell('B9').value || '',
      contractName: sheet.getCell('B10').value || '',
      firstLoanDate: sheet.getCell('B11').value || '',
      totalLoanAmount: sheet.getCell('B12').value || '',
      productName: sheet.getCell('B13').value || ''
    };
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

      const rawIsLarger30 = sheet.getCell(`${colAmount}10`).value;
      const isLarger30 = rawIsLarger30 === '是';
      console.log(`第 ${groupIndex + 1} 组 是否大于30天 (${colAmount}10): ${
          isLarger30 ? '是（true）' : '否（false）'}`);

      const repaymentCell = sheet.getCell(`${colAmount}11`).value || '';
      console.log(
          `第 ${groupIndex + 1} 组 还款方式原始值 (${colAmount}11):`,
          repaymentCell);

      const repayment =
          (typeof repaymentCell === 'object' && repaymentCell.richText) ?
          repaymentCell.richText.map(part => part.text).join('').trim() :
          String(repaymentCell || '').trim();

      console.log(`第 ${groupIndex + 1} 组 还款方式解析后: ${repayment}`);

      let repaymentType = null;
      if (repayment.includes('先息后本')) {
        repaymentType = 1;
      } else if (repayment.includes('等额本金')) {
        repaymentType = 2;
      } else if (repayment.includes('等额本息')) {
        repaymentType = 3;
      } else {
        throw new Error(`第 ${
            groupIndex +
            1} 组的还款方式必须为先息后本、等额本金、等额本息，但实际为：${
            repayment}`);
      }

      const interestDay = parseInt(sheet.getCell(`${colAmount}12`).value) || 21;
      console.log(
          `第 ${groupIndex + 1} 组 结息日 (${colAmount}12): ${interestDay} 日`);

      const cell = sheet.getCell(`${colAmount}14`);
      // 先拿到单元格的“原始值”(raw)
      const raw = (typeof cell.value === 'object' && cell.value !== null &&
                   'formula' in cell.value) ?
          cell.value.result  // 公式缓存
          :
          cell.value;  // 普通单元格

      // 把 undefined / null / NaN 统统替换成 0
      const intimeTerm = Number.isFinite(Number(raw)) ? Number(raw) : 0;
      console.log(
          `第 ${groupIndex + 1} 组 提前还款期限 (${colAmount}14): ${
              intimeTerm} 天`,
      );

      //  从第13行开始读取还款明细
      const paymentPairs = [];
      let row = 15;

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
        sheetNum: groups.length + 1,
        sheetName: `借款${groups.length + 1}`,
        name,
        amount,
        rate,
        lateRate,
        term,
        startDate,
        endDate,
        isLarger30,
        repayment,
        interestDay,
        intimeTerm,
        currentDate,
        repaymentType,
        paymentPairs
      });
    }

    const {excelbuffer, bufferName} =
        await generateExcelBuffer(groups, infoGroup);

    // 将文件内容转成 base64 字符串
    const base64 = excelbuffer.toString('base64');

    res.json({
      buffer: base64,
      bufferName: bufferName + '.xlsx',  // 加上扩展名
    });
  } catch (err) {
    console.error('生成失败:', err);
    res.status(500).json({message: `生成失败: ${err.message}`});
  }
});

app.post('/generate-excel-zip', upload.any(), async (req, res) => {
  // const now = new Date();
  // const deadline = new Date('2025-07-10');

  // if (now > deadline) {
  //   return res.status(403).json(
  //       {message: '超过试用日期，无法继续使用此功能。'});
  // }

  try {
    const currentDate = req.body.currentDate;
    const files = req.files;

    if (!files || files.length === 0) {
      return res.status(400).json({message: '未上传任何文件'});
    }

    const zip = new JSZip();

    for (const file of files) {
      try {
        const buffer = file.buffer;
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(buffer);
        const sheet = workbook.worksheets[0];
        const name = sheet.getCell('B1').value || '';

        const infoGroup = {
          idNumber: sheet.getCell('B4').value || '',
          registeredAddress: sheet.getCell('B5').value || '',
          contactAddress: sheet.getCell('B6').value || '',
          defendantPhoneNumber: sheet.getCell('B7').value || '',
          lawyerFee: sheet.getCell('B8').value || '',
          contractSignDate: sheet.getCell('B9').value || '',
          contractName: sheet.getCell('B10').value || '',
          firstLoanDate: sheet.getCell('B11').value || '',
          totalLoanAmount: sheet.getCell('B12').value || '',
          productName: sheet.getCell('B13').value || ''
        };

        const groups = [];
        const maxGroups = 30;
        const allColumns = generateExcelColumnNames(maxGroups * 4);

        for (let groupIndex = 0; groupIndex < maxGroups; groupIndex++) {
          const baseIdx = 3 + groupIndex * 4;

          const colDate = allColumns[baseIdx];  // 日期列
          const colAmount =
              allColumns[baseIdx + 1];  // 金额列（也是参数基准列）
          const colType = allColumns[baseIdx + 2];  // 类型列

          console.log(` 正在处理第 ${groupIndex + 1} 组，列分别为：日期列 = ${
              colDate}，金额列 = ${colAmount}，类型列 = ${colType}`);

          // 参数区域（第 4~12 行）以金额列为基准
          const amount = parseFloat(sheet.getCell(`${colAmount}4`).value) || 0;

          //  如果关键字段为空，说明没有这一组了，跳出循环
          if (!amount) {
            break;
          }

          console.log(
              `第 ${groupIndex + 1} 组 金额 (${colAmount}4): ${amount}`);

          const rate =
              parseFloat(sheet.getCell(`${colAmount}5`).value) * 100 || 0;
          console.log(`第 ${groupIndex + 1} 组 利率 (${colAmount}5): ${rate}%`);

          const lateRate =
              parseFloat(sheet.getCell(`${colAmount}6`).value) * 100 || 0;
          console.log(
              `第 ${groupIndex + 1} 组 罚息利率 (${colAmount}6): ${lateRate}%`);

          const term = parseInt(sheet.getCell(`${colAmount}7`).value) || 0;
          console.log(
              `第 ${groupIndex + 1} 组 期限 (${colAmount}7): ${term} 月`);

          const startDate = sheet.getCell(`${colAmount}8`).value || '';
          console.log(
              `第 ${groupIndex + 1} 组 起始日 (${colAmount}8): ${startDate}`);

          const endDate = sheet.getCell(`${colAmount}9`).value || '';
          console.log(
              `第 ${groupIndex + 1} 组 到期日 (${colAmount}9): ${endDate}`);

          const rawIsLarger30 = sheet.getCell(`${colAmount}10`).value;
          const isLarger30 = rawIsLarger30 === '是';
          console.log(`第 ${groupIndex + 1} 组 是否大于30天 (${colAmount}10): ${
              isLarger30 ? '是（true）' : '否（false）'}`);

          const repaymentCell = sheet.getCell(`${colAmount}11`).value || '';
          console.log(
              `第 ${groupIndex + 1} 组 还款方式原始值 (${colAmount}11):`,
              repaymentCell);

          const repayment =
              (typeof repaymentCell === 'object' && repaymentCell.richText) ?
              repaymentCell.richText.map(part => part.text).join('').trim() :
              String(repaymentCell || '').trim();

          console.log(`第 ${groupIndex + 1} 组 还款方式解析后: ${repayment}`);

          let repaymentType = null;
          if (repayment.includes('先息后本')) {
            repaymentType = 1;
          } else if (repayment.includes('等额本金')) {
            repaymentType = 2;
          } else if (repayment.includes('等额本息')) {
            repaymentType = 3;
          } else {
            throw new Error(`第 ${
                groupIndex +
                1} 组的还款方式必须为先息后本、等额本金、等额本息，但实际为：${
                repayment}`);
          }

          const interestDay =
              parseInt(sheet.getCell(`${colAmount}12`).value) || 21;
          console.log(`第 ${groupIndex + 1} 组 结息日 (${colAmount}12): ${
              interestDay} 日`);

          const cell = sheet.getCell(`${colAmount}14`);
          // ① 先拿到单元格的“原始值”(raw)
          const raw = (typeof cell.value === 'object' && cell.value !== null &&
                       'formula' in cell.value) ?
              cell.value.result  // 公式缓存
              :
              cell.value;  // 普通单元格

          // ② 把 undefined / null / NaN 统统替换成 0
          const intimeTerm = Number.isFinite(Number(raw)) ? Number(raw) : 0;
          console.log(
              `第 ${groupIndex + 1} 组 提前还款期限 (${colAmount}14): ${
                  intimeTerm} 天`,
          );

          //  从第15行开始读取还款明细
          const paymentPairs = [];
          let row = 15;

          while (true) {
            const dateCell = sheet.getCell(`${colDate}${row}`).value;
            const valueCell = sheet.getCell(`${colAmount}${row}`).value;
            const typeCell = sheet.getCell(`${colType}${row}`).value;

            if (valueCell === null || valueCell === undefined ||
                valueCell === '') {
              break;
            }

            paymentPairs.push({
              date: dateCell,
              value: parseFloat(valueCell),
              type: (typeCell || '').toString().trim()
            });

            row++;
          }

          groups.push({
            sheetNum: groups.length + 1,
            sheetName: `借款${groups.length + 1}`,
            name,
            amount,
            rate,
            lateRate,
            term,
            startDate,
            endDate,
            isLarger30,
            repayment,
            interestDay,
            intimeTerm,
            currentDate,
            repaymentType,
            paymentPairs
          });
        }

        const {excelbuffer, bufferName} =
            await generateExcelBuffer(groups, infoGroup);
        zip.file(bufferName + '.xlsx', excelbuffer);

      } catch (err) {
        console.error(`处理文件 ${file.originalname} 出错:`, err.message);
        // 可以继续处理其他文件，不中断整个流程
      }
    }

    const zipBuffer = await zip.generateAsync({type: 'nodebuffer'});
    res.setHeader('Content-Type', 'application/zip');
    res.setHeader(
        'Content-Disposition',
        `attachment; filename="download.zip"; filename*=UTF-8''${
            encodeURIComponent('借款明细打包.zip')}`);
    res.send(zipBuffer);

  } catch (err) {
    console.error('批量生成失败:', err);
    res.status(500).json({message: `批量生成失败: ${err.message}`});
  }
});

// =====================
// 下载输入表模板
// =====================
app.get('/download-template', (req, res) => {
  const filePath = path.join(__dirname, 'public', 'template', '输入表.xlsx');
  res.download(filePath, '输入表.xlsx', (err) => {
    if (err) {
      console.error('下载模板失败:', err);
      res.status(500).send('下载失败');
    }
  });
});

// =====================
// 启动服务
// =====================
app.listen(PORT, () => {
  const url = `http://localhost:${PORT}`;
  console.log(`✅ 服务运行中：${url}`);
  // openUrl(url);
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

function getNextOverdueLabel(prevValue) {
  const str = String(prevValue).trim();

  // case: "9(1)" → 提取主编号 & 次编号
  const matchParen = str.match(/^(\d+)\((\d+)\)$/);
  if (matchParen) {
    const main = matchParen[1];
    const sub = parseInt(matchParen[2], 10) + 1;
    return `${main}(${sub})`;
  }

  // case: "9" → 变成 "9(1)"
  const matchNumberOnly = str.match(/^\d+$/);
  if (matchNumberOnly) {
    return `${str}(1)`;
  }

  // case: "1-7" → 取最后一个数字作为主编号
  const matchRange = str.match(/^(\d+)-(\d+)$/);
  if (matchRange) {
    const main = matchRange[2];
    return `${main}(1)`;
  }

  // fallback: unknown format, return as is
  return `${str}(1)`;
}

function formatChineseDate(dateStr) {
  if (!dateStr) return '未知日期';

  const parts = dateStr.split(/[\/\-]/);  // 支持 "2025/2/10" 或 "2025-2-10"
  if (parts.length !== 3) return '未知日期';

  const [year, month, day] = parts;
  return `${year}年${parseInt(month)}月${parseInt(day)}日`;
}

function getExcelColLetter(n) {
  let s = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}