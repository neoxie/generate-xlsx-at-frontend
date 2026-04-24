# Frontend XLSX Export - Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build an MVP that fetches employee data from a backend API and generates a multi-sheet, richly formatted XLSX file entirely in the browser for download.

**Architecture:** Three-layer separation — API route returns mock JSON data, a pure `xlsx-generator` module builds the ExcelJS workbook with per-department sheets and conditional formatting, and a client page component orchestrates fetch (via TanStack Query) + generate + download.

**Tech Stack:** Next.js 16, React 19, ExcelJS, TanStack Query v5, Tailwind CSS 4

---

## File Structure

| Action | File | Responsibility |
|--------|------|----------------|
| Create | `lib/types.ts` | Shared `Employee` interface |
| Create | `app/api/employees/route.ts` | GET API — returns mock employee JSON |
| Create | `lib/xlsx-generator.ts` | ExcelJS workbook generation with multi-sheet + formatting |
| Create | `lib/query-provider.tsx` | TanStack Query client provider wrapper (client component) |
| Modify | `app/layout.tsx` | Wrap children with `QueryProvider` |
| Modify | `app/page.tsx` | Download button, TanStack Query fetch, error handling |

---

### Task 1: Install dependencies

- [ ] **Step 1: Install exceljs and @tanstack/react-query**

Run:
```bash
npm install exceljs @tanstack/react-query
```

Expected: Both packages added to `dependencies` in `package.json`.

---

### Task 2: Create shared types

**Files:**
- Create: `lib/types.ts`

- [ ] **Step 1: Create `lib/types.ts`**

```typescript
export interface Employee {
  id: number;
  name: string;
  department: string;
  position: string;
  hireDate: string;
  salary: number;
  performance: 'A' | 'B' | 'C' | 'D';
}
```

---

### Task 3: Create API route with mock data

**Files:**
- Create: `app/api/employees/route.ts`

- [ ] **Step 1: Create `app/api/employees/route.ts`**

```typescript
import type { Employee } from '@/lib/types';

const employees: Employee[] = [
  { id: 1, name: '张伟', department: '技术部', position: '高级工程师', hireDate: '2020-03-15', salary: 25000, performance: 'A' },
  { id: 2, name: '李娜', department: '技术部', position: '前端开发', hireDate: '2021-07-01', salary: 18000, performance: 'B' },
  { id: 3, name: '王强', department: '技术部', position: '后端开发', hireDate: '2022-01-10', salary: 16000, performance: 'A' },
  { id: 4, name: '赵敏', department: '技术部', position: '测试工程师', hireDate: '2023-05-20', salary: 14000, performance: 'C' },
  { id: 5, name: '陈磊', department: '技术部', position: '架构师', hireDate: '2019-11-03', salary: 35000, performance: 'A' },
  { id: 6, name: '刘芳', department: '市场部', position: '市场经理', hireDate: '2020-06-15', salary: 22000, performance: 'A' },
  { id: 7, name: '杨洋', department: '市场部', position: '市场专员', hireDate: '2022-09-01', salary: 12000, performance: 'B' },
  { id: 8, name: '吴静', department: '市场部', position: '品牌策划', hireDate: '2021-03-10', salary: 17000, performance: 'B' },
  { id: 9, name: '孙丽', department: '人事部', position: 'HR经理', hireDate: '2019-08-20', salary: 20000, performance: 'A' },
  { id: 10, name: '周杰', department: '人事部', position: '招聘专员', hireDate: '2023-02-14', salary: 13000, performance: 'C' },
  { id: 11, name: '黄鑫', department: '财务部', position: '财务主管', hireDate: '2020-01-06', salary: 21000, performance: 'B' },
  { id: 12, name: '马超', department: '财务部', position: '会计', hireDate: '2022-07-15', salary: 15000, performance: 'B' },
  { id: 13, name: '林佳', department: '财务部', position: '出纳', hireDate: '2024-01-08', salary: 11000, performance: 'D' },
];

export async function GET() {
  return Response.json(employees);
}
```

---

### Task 4: Create xlsx generator module

**Files:**
- Create: `lib/xlsx-generator.ts`

- [ ] **Step 1: Create `lib/xlsx-generator.ts`**

```typescript
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
  for (const cell of row.cells) {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: HEADER_BG } };
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
    cell.border = thinBorder as ExcelJS.Borders;
  }
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
```

---

### Task 5: Set up TanStack Query provider

**Files:**
- Create: `lib/query-provider.tsx`
- Modify: `app/layout.tsx`

- [ ] **Step 1: Create `lib/query-provider.tsx`**

```tsx
'use client';

import { QueryClient, QueryClientProvider } from '@tanstack/react-query';
import { useState, type ReactNode } from 'react';

export function QueryProvider({ children }: { children: ReactNode }) {
  const [queryClient] = useState(
    () =>
      new QueryClient({
        defaultOptions: {
          queries: { staleTime: 60_000 },
        },
      }),
  );

  return <QueryClientProvider client={queryClient}>{children}</QueryClientProvider>;
}
```

- [ ] **Step 2: Update `app/layout.tsx` — wrap children with QueryProvider, update metadata and lang**

Replace the full file content with:

```tsx
import type { Metadata } from 'next';
import { Geist, Geist_Mono } from 'next/font/google';
import { QueryProvider } from '@/lib/query-provider';
import './globals.css';

const geistSans = Geist({
  variable: '--font-geist-sans',
  subsets: ['latin'],
});

const geistMono = Geist_Mono({
  variable: '--font-geist-mono',
  subsets: ['latin'],
});

export const metadata: Metadata = {
  title: '员工信息导出',
  description: 'Frontend XLSX generation demo',
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="zh-CN" className={`${geistSans.variable} ${geistMono.variable} h-full antialiased`}>
      <body className="min-h-full flex flex-col">
        <QueryProvider>{children}</QueryProvider>
      </body>
    </html>
  );
}
```

---

### Task 6: Build the download page

**Files:**
- Modify: `app/page.tsx`

- [ ] **Step 1: Replace `app/page.tsx` with the download page**

```tsx
'use client';

import { useState } from 'react';
import { useQuery } from '@tanstack/react-query';
import { generateXlsx } from '@/lib/xlsx-generator';
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
      const buffer = await generateXlsx(result.data);
      const blob = new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `员工信息表_${new Date().toISOString().slice(0, 10)}.xlsx`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
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

### Task 7: Verify end-to-end

- [ ] **Step 1: Start dev server**

Run:
```bash
npm run dev
```

Expected: Server starts on `http://localhost:3000`.

- [ ] **Step 2: Verify API returns data**

Open browser to `http://localhost:3000/api/employees`.

Expected: JSON array of 13 employee objects.

- [ ] **Step 3: Verify download**

Click "导出员工信息表" button on `http://localhost:3000`.

Expected:
- Button shows "生成中..." then restores
- A file `员工信息表_YYYY-MM-DD.xlsx` downloads
- Open in Excel/Numbers: 4 sheets (技术部, 市场部, 人事部, 财务部)
- Each sheet has: frozen header row with blue background, auto-filter, alternating row colors, date format, salary number format, performance grade cell fills, salary conditional font colors
