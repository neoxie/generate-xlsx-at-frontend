# Frontend XLSX Export MVP - Design Spec

## Overview

A Next.js 16 MVP that generates formatted XLSX files entirely in the browser. A backend API provides employee data, the frontend fetches it via TanStack Query, applies rich formatting via ExcelJS with multiple sheets (one per department), and triggers a download.

## Architecture & File Structure

```
app/
├── api/
│   └── employees/
│       └── route.ts          # GET API - returns mock employee data
├── page.tsx                  # Main page with download button
├── layout.tsx                # Layout (existing)
└── globals.css               # Styles (existing)

lib/
└── xlsx-generator.ts         # ExcelJS generation logic (isolated module)
```

Three layers with single responsibilities:

- **API layer**: `/api/employees` returns JSON array of employees, data-only responsibility
- **Generation layer**: `lib/xlsx-generator.ts` receives data and produces xlsx Blob, pure logic
- **Page layer**: `page.tsx` orchestrates: button click -> TanStack Query fetch -> generate -> download

### Data Structure

```typescript
interface Employee {
  id: number;
  name: string;
  department: string;
  position: string;
  hireDate: string;       // ISO date string
  salary: number;
  performance: 'A' | 'B' | 'C' | 'D';
}
```

## Technology Choices

### XLSX Generation: ExcelJS

Node.js/browser dual-environment XLSX generation library. Selected over SheetJS (community edition lacks full style support) and xlsx-js-style (community fork with uncertain maintenance). ExcelJS provides the most complete free styling API with active maintenance (14k+ GitHub stars).

### Data Fetching: TanStack Query

React Query library for server state management. Provides caching, loading/error states, and retry logic out of the box. Used as the data fetching layer between API and xlsx generation.

## XLSX Format Specification

### Multi-Sheet Structure

Data is split into multiple sheets by department. Each sheet:

- Named after the department (e.g., "技术部", "市场部")
- Contains only employees belonging to that department
- Has identical column structure and formatting

### Header Row (applied to each sheet)

- Deep blue background (#1F4E79), white bold text
- Row height 30, header row frozen
- Auto-filter enabled on header row

### Columns (per sheet, department column excluded since sheet name already identifies dept)

| Column | Content | Width | Format       |
|--------|---------|-------|--------------|
| A      | ID      | 8     | Center       |
| B      | Name    | 14    | Left         |
| C      | Title   | 18    | Left         |
| D      | Date    | 14    | YYYY-MM-DD   |
| E      | Salary  | 14    | #,##0 Center |
| F      | Grade   | 10    | Center       |

### Data Rows

- Alternating row colors: white / light blue (#D6E4F0)
- Thin borders on all cells

### Conditional Coloring - Performance Grade (Column F)

- A -> Green (#92D050)
- B -> Blue (#00B0F0)
- C -> Yellow (#FFC000)
- D -> Red (#FF0000) with white text

### Conditional Coloring - Salary (Column E)

- >= 20000 -> Green text
- >= 15000 and < 20000 -> Blue text
- < 15000 -> Red text

## Interaction & Error Handling

### Page Interaction

- Single button labeled "导出员工信息表"
- TanStack Query manages fetch state: loading, error, success
- Button shows "生成中..." during data fetch + generation
- On success: browser triggers download, filename format `员工信息表_YYYY-MM-DD.xlsx`

### Error Handling

- TanStack Query handles API errors with built-in error state
- API call failure: show red text "获取数据失败，请重试"
- XLSX generation failure: show "文件生成失败，请重试"
- Error messages auto-dismiss after 3 seconds

## Out of Scope (YAGNI)

- No pagination, filtering, or search
- No progress bar
- No export history
- No internationalization
