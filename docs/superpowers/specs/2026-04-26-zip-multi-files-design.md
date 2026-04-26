# Zip Multi-Files Design

将"单文件多 sheet"导出改为"每个部门一个 xlsx 文件，打包成 zip 下载"。

## 需求

- 每个部门生成独立的 `.xlsx` 文件，文件名 `部门名称.xlsx`
- 每个 xlsx 内部样式与现有完全一致（表头、交替行、冻结首行、自动筛选、薪资/绩效条件格式）
- 所有 xlsx 打包成 zip，zip 文件名 `员工信息表_YYYY-MM-DD.zip`
- 前端生成，不引入后端逻辑

## 方案

使用 JSZip（~45KB gzipped，npm 周下载 ~18M）在前端打包。

## 文件改动

| 文件 | 改动 |
|------|------|
| `lib/xlsx-generator.ts` | 抽取单部门生成逻辑，新增导出 `generateXlsxZip()` |
| `app/page.tsx` | 调用改为 `generateXlsxZip()`，MIME type 改为 zip |
| `package.json` | 新增 `jszip` 依赖 |

不动：`lib/types.ts`、`lib/query-provider.tsx`、`app/api/employees/route.ts`、`app/layout.tsx`

## 详细设计

### `lib/xlsx-generator.ts`

**新增私有函数 `createDepartmentSheet(employees: Employee[]): ExcelJS.Workbook`**

为一个部门的员工创建独立 workbook，应用现有全部样式。逻辑直接从现有 `generateXlsx` 的 for 循环体中提取，包括：

- 设置列宽
- 添加表头行并应用 `applyHeaderStyle`
- 冻结首行（`ws.views`）
- 自动筛选（`ws.autoFilter`）
- 通过 `applyDataStyles` 应用数据行样式

**新增导出 `generateXlsxZip(employees: Employee[]): Promise<ArrayBuffer>`**

1. 按部门分组（复用现有 `groupByDepartment`）
2. 对每个部门调用 `createDepartmentSheet()` 生成 workbook
3. 将每个 workbook 的 `writeBuffer()` 结果写入 JSZip 实例，key 为 `部门名称.xlsx`
4. 返回 `zip.generateAsync({ type: 'arraybuffer' })`

**保留 `generateXlsx`**，标记 `@deprecated`。

### `app/page.tsx`

- 导入改为 `generateXlsxZip`
- Blob MIME type: `application/zip`
- 文件名后缀: `.xlsx` → `.zip`
- 其余逻辑（TanStack Query、错误处理、UI）不变

### 依赖

```bash
npm install jszip
```

JSZip 自带 TypeScript 类型，无需额外 `@types` 包。

## 错误处理

与现有逻辑一致：try-catch 包裹，错误信息显示 3 秒后自动消失。JSZip 和 ExcelJS 的异常均在同一 catch 中处理。
