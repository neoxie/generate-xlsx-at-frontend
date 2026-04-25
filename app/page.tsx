'use client';

import { useState } from 'react';
import { useQuery } from '@tanstack/react-query';
import { saveAs } from 'file-saver';
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
      saveAs(blob, `员工信息表_${new Date().toISOString().slice(0, 10)}.xlsx`);
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
