# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

```bash
npm run dev       # Start dev server on http://localhost:3000
npm run build     # Production build
npm run lint      # ESLint
npx tsc --noEmit  # Type check without emitting
```

No test framework is configured.

## Architecture

Frontend-only XLSX generation app. Employee data flows through three layers:

1. **API layer** (`app/api/employees/route.ts`) — Next.js route handler returning hardcoded mock JSON. Uses standard `Response.json()` (not `NextResponse`).

2. **Generation layer** (`lib/xlsx-generator.ts`) — Pure function `generateXlsx(employees)` that groups data by department into separate ExcelJS worksheets, applies rich formatting (header styling, alternating rows, conditional coloring for salary/performance), and returns a write buffer.

3. **Page layer** (`app/page.tsx` + `lib/query-provider.tsx`) — Client component using TanStack Query (`enabled: false` + `refetch()`) to fetch data on button click, then generates and downloads the XLSX blob.

### Key patterns

- `lib/query-provider.tsx` wraps `QueryClientProvider` as a `"use client"` boundary, injected into the server component `layout.tsx`.
- ExcelJS runs entirely in the browser (client component). `workbook.xlsx.writeBuffer()` produces an ArrayBuffer-compatible buffer passed directly to `new Blob()`.
- `@/*` path alias maps to project root (configured in tsconfig.json).

## Next.js version

This project uses Next.js 16 with React 19. **Breaking changes exist** — read `node_modules/next/dist/docs/` before using unfamiliar APIs. Route handlers use Web standard `Response.json()`, and `context.params` is a Promise.
