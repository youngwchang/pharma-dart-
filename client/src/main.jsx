import React from 'react';
import ReactDOM from 'react-dom/client';
import App from './app';

const root = document.getElementById('root');
console.log('root element:', root);

try {
  ReactDOM.createRoot(root).render(
    <React.StrictMode>
      <App />
    </React.StrictMode>
  );
  console.log('React mounted successfully');
} catch(e) {
  console.error('React mount failed:', e);
  root.innerHTML = `<div style="color:red;padding:20px;font-family:monospace">
    <h2>앱 로드 실패</h2>
    <pre>${e.message}</pre>
  </div>`;
}
