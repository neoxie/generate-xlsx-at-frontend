export interface Employee {
  id: number;
  name: string;
  department: string;
  position: string;
  hireDate: string;
  salary: number;
  performance: 'A' | 'B' | 'C' | 'D';
}
