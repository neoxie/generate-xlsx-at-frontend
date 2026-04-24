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
