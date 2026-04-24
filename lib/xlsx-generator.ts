import ExcelJS from 'exceljs';
import type { Employee } from './types';

const HEADERS = ['工号', '姓名', '职位', '入职日期', '薪资', '绩效等级'];
const COLUMN_WIDTHS = [8, 14, 18, 14, 14, 10];

const HEADER_BG = 'FF1F4E79';
const ALT_ROW_BG = 'FFD6E4F0';

const PERFORMANCE_FILLS: Record<string, { bg: string; font?: Partial<ExcelJS.Font> }> = {
  A: { bg: 'FF92D050' },
  B: { bg: 'FF00B0F0' },
  C: { bg: 'FFFFC000' },
  D: { bg: 'FFFF0000', font: { color: { argb: 'FFFFFFFF' }, bold: true } },
};

function salaryFontColor(salary: number): string {
  if (salary >= 20000) return 'FF00B050';
  if (salary >= 15000) return 'FF00B0F0';
  return 'FFFF0000';
}

function groupByDepartment(employees: Employee[]) {
  const map = new Map<string, Employee[]>();
  for (const emp of employees) {
    const list = map.get(emp.department) ?? [];
    list.push(emp);
    map.set(emp.department, list);
  }
  return Array.from(map, ([department, employees]) => ({ department, employees }));
}

const thinBorder: Partial<ExcelJS.Borders> = {
  top: { style: 'thin' },
  left: { style: 'thin' },
  bottom: { style: 'thin' },
  right: { style: 'thin' },
};

function applyHeaderStyle(row: ExcelJS.Row) {
  row.height = 30;
  row.eachCell({ includeEmpty: true }, (cell) => {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: HEADER_BG } };
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
    cell.border = thinBorder as ExcelJS.Borders;
  });
}

function applyDataStyles(ws: ExcelJS.Worksheet, employees: Employee[]) {
  employees.forEach((emp, i) => {
    const row = ws.addRow([
      emp.id,
      emp.name,
      emp.position,
      new Date(emp.hireDate),
      emp.salary,
      emp.performance,
    ]);

    row.getCell(4).numFmt = 'YYYY-MM-DD';
    row.getCell(5).numFmt = '#,##0';

    const isAlt = i % 2 === 1;
    for (let col = 1; col <= HEADERS.length; col++) {
      const cell = row.getCell(col);
      cell.border = thinBorder as ExcelJS.Borders;
      if (isAlt) {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: ALT_ROW_BG } };
      }
    }

    // Alignment: name(B) and title(C) left, rest center
    row.getCell(1).alignment = { vertical: 'middle', horizontal: 'center' };
    row.getCell(2).alignment = { vertical: 'middle', horizontal: 'left' };
    row.getCell(3).alignment = { vertical: 'middle', horizontal: 'left' };
    row.getCell(4).alignment = { vertical: 'middle', horizontal: 'center' };
    row.getCell(5).alignment = { vertical: 'middle', horizontal: 'center' };
    row.getCell(6).alignment = { vertical: 'middle', horizontal: 'center' };

    // Performance grade fill
    const perfStyle = PERFORMANCE_FILLS[emp.performance];
    if (perfStyle) {
      const gradeCell = row.getCell(6);
      gradeCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: perfStyle.bg } };
      if (perfStyle.font) gradeCell.font = perfStyle.font;
    }

    // Salary font color
    row.getCell(5).font = { color: { argb: salaryFontColor(emp.salary) } };
  });
}

export async function generateXlsx(employees: Employee[]) {
  const workbook = new ExcelJS.Workbook();
  const sheets = groupByDepartment(employees);

  for (const { department, employees: emps } of sheets) {
    const ws = workbook.addWorksheet(department);
    ws.columns = COLUMN_WIDTHS.map((width) => ({ width }));

    const headerRow = ws.addRow(HEADERS);
    applyHeaderStyle(headerRow);

    ws.views = [{ state: 'frozen', ySplit: 1 }];
    ws.autoFilter = {
      from: { row: 1, column: 1 },
      to: { row: 1, column: HEADERS.length },
    };

    applyDataStyles(ws, emps);
  }

  return workbook.xlsx.writeBuffer();
}
