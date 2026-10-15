import React from 'react';
import ReactDOM from 'react-dom/client';
import App from './ui/App';
import './index.css';

// ตั้งธีมก่อนวาดครั้งแรก (5C): ใน Electron ใช้ค่าที่ main ส่งมา, ในเบราว์เซอร์ดูโหมดของระบบ
{
  const initial = (window as any).electronAPI?.initialTheme;
  const dark =
    initial === 'dark' ||
    (initial === undefined && typeof matchMedia === 'function' && matchMedia('(prefers-color-scheme: dark)').matches);
  document.documentElement.dataset.theme = dark ? 'dark' : 'light';
}

ReactDOM.createRoot(document.getElementById('root') as HTMLElement).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
);
