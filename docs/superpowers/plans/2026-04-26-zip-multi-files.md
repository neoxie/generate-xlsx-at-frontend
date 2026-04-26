# Zip Multi-Files Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 将单文件多 sheet 导出改为每个部门一个 xlsx 文件，打包成 zip 下载。

**Architecture:** 从 `lib/xlsx-generator.ts` 抽取单部门 workbook 生成逻辑，新增 `generateXlsxZip()` 使用 JSZip 打包多个 xlsx。`app/page.tsx` 切换调用并适配 zip MIME type。

**Tech Stack:** ExcelJS, JSZip, file-saver, TanStack Query

---

### Task 1: Install jszip as direct dependency

**Files:**
- Modify: `package.json`

- [ ] **Step 1: Install jszip**

Run: `npm install jszip`

- [ ] **Step 2: Verify type declarations exist**

Run: `ls node_modules/jszip/index.d.ts`
Expected: file exists (JSZip ships its own types)

- [ ] **Step 3: Type check baseline**

Run: `npx tsc --noEmit`
Expected: no errors (existing code still compiles)

---

### Task 2: Refactor xlsx-generator.ts — extract per-department workbook creation and add generateXlsxZip

**Files:**
- Modify: `lib/xlsx-generator.ts`

- [ ] **Step 1: Add JSZip import and create `createDepartmentWorkbook` function, refactor `generateXlsx`, add `generateXlsxZip`**

Replace the entire file content with:

```typescript
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
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

function createDepartmentWorkbook(employees: Employee[]): ExcelJS.Workbook {
  const workbook = new ExcelJS.Workbook();
  const ws = workbook.addWorksheet('Sheet1');
  ws.columns = COLUMN_WIDTHS.map((width) => ({ width }));

  const headerRow = ws.addRow(HEADERS);
  applyHeaderStyle(headerRow);

  ws.views = [{ state: 'frozen', ySplit: 1 }];
  ws.autoFilter = {
    from: { row: 1, column: 1 },
    to: { row: 1, column: HEADERS.length },
  };

  applyDataStyles(ws, employees);
  return workbook;
}

/** @deprecated Use generateXlsxZip instead for per-department files. */
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

export async function generateXlsxZip(employees: Employee[]): Promise<ArrayBuffer> {
  const zip = new JSZip();
  const sheets = groupByDepartment(employees);

  for (const { department, employees: emps } of sheets) {
    const workbook = createDepartmentWorkbook(emps);
    const buffer = await workbook.xlsx.writeBuffer();
    zip.file(`${department}.xlsx`, buffer);
  }

  return zip.generateAsync({ type: 'arraybuffer' });
}
```

- [ ] **Step 2: Type check**

Run: `npx tsc --noEmit`
Expected: no errors

---

### Task 3: Update page.tsx to use generateXlsxZip

**Files:**
- Modify: `app/page.tsx`

- [ ] **Step 1: Update import, blob creation, and filename**

Change the import on line 6:

```typescript
import { generateXlsxZip } from '@/lib/xlsx-generator';
```

Change lines 31-35 (the buffer + blob + saveAs block):

```typescript
      const buffer = await generateXlsxZip(result.data);
      const blob = new Blob([buffer], { type: 'application/zip' });
      saveAs(blob, `员工信息表_${new Date().toISOString().slice(0, 10)}.zip`);
```

- [ ] **Step 2: Type check**

Run: `npx tsc --noEmit`
Expected: no errors

- [ ] **Step 3: Verify the full page.tsx reads correctly**

The final file should be:

```typescript
'use client';

import { useState } from 'react';
import { useQuery } from '@tanstack/react-query';
import { saveAs } from 'file-saver';
import { generateXlsxZip } from '@/lib/xlsx-generator';
import type { Employee } from '@/lib/types';

export default function Home() {
  const [downloading, setDownloading] = useState(false);
  const [errorMsg, setErrorMsg] = useState('');

  const { refetch } = useQuery({
    queryKey: ['employees'],
    queryFn: async (): Promise<Employee[]> => {
      const res = await fetch('/api/employees');
      if (!res.ok) throw new Error('获取数据失败');
      return res.json();
    },
    enabled: false,
  });

  const handleDownload = async () => {
    setDownloading(true);
    setErrorMsg('');
    try {
      const result = await refetch();
      if (result.isError || !result.data) {
        throw new Error(result.error?.message || '获取数据失败');
      }
      const buffer = await generateXlsxZip(result.data);
      const blob = new Blob([buffer], { type: 'application/zip' });
      saveAs(blob, `员工信息表_${new Date().toISOString().slice(0, 10)}.zip`);
    } catch (err) {
      setErrorMsg(err instanceof Error ? err.message : '文件生成失败');
      setTimeout(() => setErrorMsg(''), 3000);
    } finally {
      setDownloading(false);
    }
  };

  return (
    <main className="min-h-screen flex items-center justify-center bg-gray-50">
      <div className="text-center">
        <h1 className="text-2xl font-bold text-gray-800 mb-6">员工信息导出</h1>
        <button
          onClick={handleDownload}
          disabled={downloading}
          className="px-6 py-3 bg-blue-600 text-white rounded-lg hover:bg-blue-700 disabled:bg-gray-400 disabled:cursor-not-allowed transition-colors cursor-pointer"
        >
          {downloading ? '生成中...' : '导出员工信息表'}
        </button>
        {errorMsg && <p className="mt-4 text-red-600">{errorMsg}</p>}
      </div>
    </main>
  );
}
```

---

### Task 4: End-to-end verification

- [ ] **Step 1: Start dev server**

Run: `npm run dev`

- [ ] **Step 2: Open http://localhost:3000 in browser, click download button**

Expected: a file named `员工信息表_YYYY-MM-DD.zip` downloads.

- [ ] **Step 3: Extract the zip and verify contents**

Expected: zip contains 4 files — `技术部.xlsx`, `市场部.xlsx`, `人事部.xlsx`, `财务部.xlsx`. Each opens correctly in Excel with proper formatting (header style, alternating rows, conditional coloring).
