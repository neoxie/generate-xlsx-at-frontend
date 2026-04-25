# Generate XLSX at Frontend

前端生成 XLSX 文件的演示项目。通过 Next.js API 获取模拟员工数据，在浏览器端使用 ExcelJS 生成带丰富格式的 Excel 文件并触发下载。

## 技术栈

- **Next.js 16** + React 19 + TypeScript
- **ExcelJS** — 浏览器端 Excel 文件生成与格式化
- **TanStack Query** — 数据获取与状态管理
- **Tailwind CSS 4** — 页面样式

## 功能

- 按部门分 Sheet（技术部、市场部、人事部、财务部）
- 表头深色背景 + 白色加粗字体
- 奇偶行交替背景色
- 绩效等级条件着色（A 绿 / B 蓝 / C 黄 / D 红）
- 薪资数字格式化与颜色区分
- 冻结首行 + 自动筛选
- 一键导出 `.xlsx` 文件

## 快速开始

```bash
npm install
npm run dev
```

打开 http://localhost:3000，点击「导出员工信息表」按钮即可下载 Excel 文件。

## 其他命令

```bash
npm run build    # 生产构建
npm run lint     # ESLint 检查
npx tsc --noEmit # 类型检查
```

## 项目结构

```
app/
  api/employees/route.ts   # API 路由，返回模拟员工 JSON
  page.tsx                 # 下载页面（客户端组件）
  layout.tsx               # 根布局
  globals.css              # 全局样式
lib/
  xlsx-generator.ts        # ExcelJS 生成逻辑（按部门分 Sheet、格式化）
  types.ts                 # Employee 类型定义
  query-provider.tsx       # TanStack Query Provider 封装
```

## 架构

数据流经三层：

1. **API 层** — Next.js Route Handler 返回硬编码的员工 JSON
2. **生成层** — 纯函数 `generateXlsx(employees)`，使用 ExcelJS 分组创建工作表并应用格式
3. **页面层** — TanStack Query（`enabled: false` + `refetch()`）按需获取数据，生成 Blob 触发浏览器下载
